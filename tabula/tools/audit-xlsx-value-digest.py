#!/usr/bin/env python3
"""Compare all stored worksheet scalar caches through bounded XML streams; never log cell content."""
import argparse,datetime,decimal,hashlib,io,json,math,os,posixpath,re,struct,time,zipfile
from xml.parsers import expat
CHUNK=64*1024
MAX_SST_ITEMS=200000
MAX_SST_CHARS=16*1024*1024
MAX_CELL_CHARS=1024*1024

def parse_part(z,name,start=None,end=None,chars=None):
    parser=expat.ParserCreate(namespace_separator='}')
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    parser.ExternalEntityRefHandler=lambda *args:0
    def doctype(*args):raise ValueError('DTD is not supported')
    parser.StartDoctypeDeclHandler=doctype
    if start:parser.StartElementHandler=lambda tag,attrs:start(tag.rsplit('}',1)[-1],attrs)
    if end:parser.EndElementHandler=lambda tag:end(tag.rsplit('}',1)[-1])
    if chars:parser.CharacterDataHandler=chars
    with z.open(name) as stream:
        while chunk:=stream.read(CHUNK):parser.Parse(chunk,False)
    parser.Parse(b'',True)

def unescape_excel(text):
    text=re.sub(r'_x([0-9a-fA-F]{4})_',lambda m:chr(int(m[1],16)),text)
    return text.encode('utf-16-le','surrogatepass').decode('utf-16-le','surrogatepass')

def numeric(text):
    value=float(text)
    if not math.isfinite(value):raise ValueError('Non-finite stored scalar')
    if value==0:value=0.0
    return struct.pack('>d',value)

def iso_serial(text,date1904):
    stamp=datetime.datetime.fromisoformat(text.replace('Z','+00:00'))
    if stamp.tzinfo:stamp=stamp.astimezone(datetime.timezone.utc).replace(tzinfo=None)
    base=datetime.datetime(1904,1,1) if date1904 else datetime.datetime(1899,12,30)
    value=(stamp-base).total_seconds()/86400
    if not date1904 and stamp<datetime.datetime(1900,3,1):value-=1
    return numeric(str(value))

def scalar(kind,text,strings,date1904):
    if kind=='s':
        index=int(text)
        if not 0<=index<len(strings):raise ValueError('Shared-string index outside table')
        return 'string',strings[index].encode('utf-8','surrogatepass')
    if kind in ('str','inlineStr'):return 'string',unescape_excel(text).encode('utf-8','surrogatepass')
    if text=='':return 'blank',b''
    if kind=='b':
        if text not in ('0','1','true','false'):raise ValueError('Invalid Boolean cache')
        return 'boolean',b'1' if text in ('1','true') else b'0'
    if kind=='e':return 'error',text.encode('utf-8','surrogatepass')
    if kind=='d':return 'number',iso_serial(text,date1904)
    if kind not in ('n',''):raise ValueError('Unknown scalar type')
    return 'number',numeric(text)

def shared_strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist():return []
    strings=[];stack=[];chunks=[];length=0;total=0
    def start(tag,attrs):
        nonlocal chunks,length
        stack.append(tag)
        if tag=='si':chunks=[];length=0
    def chars(text):
        nonlocal length
        if stack and stack[-1]=='t' and 'rPh' not in stack:
            length+=len(text)
            if length>MAX_CELL_CHARS:raise ValueError('Shared string exceeds bounded cell limit')
            chunks.append(text)
    def end(tag):
        nonlocal total
        if tag=='si':
            value=unescape_excel(''.join(chunks));total+=len(value)
            if total>MAX_SST_CHARS or len(strings)>=MAX_SST_ITEMS:raise ValueError('Shared-string table exceeds explicit audit bound')
            strings.append(value)
        stack.pop()
    parse_part(z,'xl/sharedStrings.xml',start,end,chars)
    return strings

def workbook_parts(z):
    sheets=[];rels={};date1904=False
    def rel(tag,attrs):
        if tag=='Relationship' and attrs.get('TargetMode')!='External':rels[attrs['Id']]=posixpath.normpath(posixpath.join('xl',attrs['Target'])).lstrip('/')
    parse_part(z,'xl/_rels/workbook.xml.rels',rel)
    def book(tag,attrs):
        nonlocal date1904
        if tag=='workbookPr':date1904=attrs.get('date1904') in ('1','true')
        elif tag=='sheet':
            rid=next((v for k,v in attrs.items() if k.rsplit('}',1)[-1]=='id'),None)
            sheets.append({'part':rels[rid],'nameDigest':hashlib.sha256(attrs['name'].encode('utf-8')).hexdigest()})
    parse_part(z,'xl/workbook.xml',book)
    return sheets,date1904

def sheet_digest(z,part,strings,date1904):
    stack=[];cell=None;row=0;last_col=0;last_coord=(-1,-1)
    full=hashlib.sha256(b'WIXEL-CACHE-VALUES-v1\0');nonblank=hashlib.sha256(b'WIXEL-CACHE-VALUES-v1\0')
    out={'cells':0,'storedScalars':0,'explicitBlanks':0,'cellsWithoutValue':0,'formulas':0,'formulasWithoutCache':0,'types':{},'ordered':True}
    def start(tag,attrs):
        nonlocal cell,row,last_col
        stack.append(tag)
        if tag=='row':row=int(attrs.get('r',row+1));last_col=0
        elif tag=='c':
            address=attrs.get('r');column=last_col+1;cell_row=row
            if address:
                match=re.fullmatch(r'([A-Za-z]+)([1-9][0-9]*)',address)
                if not match:raise ValueError('Invalid cell coordinate')
                column=0
                for letter in match[1].upper():column=26*column+ord(letter)-64
                cell_row=int(match[2])
            last_col=column
            cell={'row':cell_row,'col':column,'type':attrs.get('t','n'),'formula':False,'valuePresent':False,'inlinePresent':False,'value':[],'inline':[],'length':0}
        elif cell is not None:
            if tag=='f':cell['formula']=True
            elif tag=='v':cell['valuePresent']=True
            elif tag=='is':cell['inlinePresent']=True
    def chars(text):
        if cell is None:return
        if stack[-1]=='v':target=cell['value']
        elif stack[-1]=='t' and 'is' in stack and 'rPh' not in stack:target=cell['inline']
        else:return
        cell['length']+=len(text)
        if cell['length']>MAX_CELL_CHARS:raise ValueError('Cell cache exceeds bounded audit limit')
        target.append(text)
    def end(tag):
        nonlocal cell,last_coord
        if tag=='c':
            out['cells']+=1;out['formulas']+=int(cell['formula'])
            present=cell['valuePresent'] or cell['inlinePresent']
            if not present:
                out['cellsWithoutValue']+=1;out['formulasWithoutCache']+=int(cell['formula'])
            else:
                kind='inlineStr' if cell['inlinePresent'] else cell['type']
                text=''.join(cell['inline'] if cell['inlinePresent'] else cell['value'])
                kind,value=scalar(kind,text,strings,date1904)
                coord=(cell['row'],cell['col'])
                if coord<=last_coord:out['ordered']=False
                last_coord=coord
                token=struct.pack('>II',*coord)+kind.encode('ascii')+b'\0'+struct.pack('>Q',len(value))+value
                full.update(token)
                if kind!='blank':nonblank.update(token)
                out['storedScalars']+=1;out['explicitBlanks']+=int(kind=='blank');out['types'][kind]=out['types'].get(kind,0)+1
            cell=None
        stack.pop()
    parse_part(z,part,start,end,chars)
    out['valueDigest']=full.hexdigest();out['nonblankValueDigest']=nonblank.hexdigest();return out

def digest(source):
    started=time.monotonic()
    with zipfile.ZipFile(source) as z:
        strings=shared_strings(z);parts,date1904=workbook_parts(z)
        sheets=[]
        for index,info in enumerate(parts):sheets.append({'index':index,**info,**sheet_digest(z,info['part'],strings,date1904)})
    return {'scope':'Every stored scalar cache; formulas counted but not expanded or independently evaluated; style-only cells excluded from value hashes and explicit blank caches counted separately','chunkBytes':CHUNK,'sharedStringCount':len(strings),'sheets':sheets,'elapsedMs':round((time.monotonic()-started)*1000)}

def compare(left,right):
    fields=['nameDigest','storedScalars','explicitBlanks','types','valueDigest','nonblankValueDigest']
    differences=[];formula_metadata=[]
    if len(left['sheets'])!=len(right['sheets']):differences.append({'kind':'sheet-count'})
    for a,b in zip(left['sheets'],right['sheets']):
        for key in fields:
            if a[key]!=b[key]:differences.append({'sheet':a['index'],'field':key})
        if not a['ordered'] or not b['ordered']:differences.append({'sheet':a['index'],'field':'ordered'})
        for key in ('formulas','formulasWithoutCache'):
            if a[key]!=b[key]:formula_metadata.append({'sheet':a['index'],'field':key})
    return {'ok':not differences,'differences':differences,'formulaMetadataDifferences':formula_metadata,'left':left,'right':right}

def fixture(sheet,sst='',date1904=False):
    data=io.BytesIO()
    with zipfile.ZipFile(data,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('xl/workbook.xml','<workbook xmlns:r="urn:rel"><workbookPr date1904="'+str(int(date1904))+'"/><sheets><sheet name="synthetic" r:id="rId1"/></sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels','<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr('xl/worksheets/sheet1.xml','<worksheet><sheetData>'+sheet+'</sheetData></worksheet>')
        if sst:z.writestr('xl/sharedStrings.xml','<sst>'+sst+'</sst>')
    data.seek(0);return data

def self_test():
    tests=0
    def same(a,b,expected=True):
        nonlocal tests
        result=compare(digest(a),digest(b));assert result['ok']==expected,result['differences'];tests+=1
    same(fixture('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1E2</v></c></row>','<si><r><t>A</t></r><r><t>한글</t></r><rPh><t>ignore</t></rPh></si>'),fixture('<row r="1"><c r="A1" t="inlineStr"><is><t>A한글</t></is></c><c r="B1"><v>100.0</v></c></row>'))
    same(fixture('<row r="1"><c r="A1" t="s"><v>0</v></c></row>','<si><t>_xD83D__xDE00__x000D__x005F_x0041_</t></si>'),fixture('<row r="1"><c r="A1" t="inlineStr"><is><t>😀_x000D__x005F_x0041_</t></is></c></row>'))
    same(fixture('<row r="1"><c r="A1" t="b"><v>true</v></c><c r="B1" t="e"><v>#N/A</v></c><c r="C1"><v>-0.0</v></c></row>'),fixture('<row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="e"><v>#N/A</v></c><c r="C1"><v>0</v></c></row>'))
    same(fixture('<row r="1"><c r="A1" s="3"/></row>'),fixture(''))
    same(fixture('<row r="1"><c r="A1"><v/></c></row>'),fixture(''),False)
    same(fixture('<row r="1"><c r="A1" t="str"><v/></c></row>'),fixture(''),False)
    same(fixture('<row r="1"><c r="A1"><v>1</v></c></row>'),fixture('<row r="1"><c r="B1"><v>1</v></c></row>'),False)
    same(fixture('<row r="1"><c r="A1"><f t="shared" si="0">A2</f><v>3</v></c></row>'),fixture('<row r="1"><c r="A1"><f>A2</f><v>3.0</v></c></row>'))
    same(fixture('<row r="1"><c r="A1" t="d"><v>1900-03-01T00:00:00</v></c></row>'),fixture('<row r="1"><c r="A1"><v>61</v></c></row>'))
    same(fixture('<row r="1"><c r="A1" t="d"><v>1904-01-01T00:00:00</v></c></row>',date1904=True),fixture('<row r="1"><c r="A1"><v>0</v></c></row>',date1904=True))
    same(fixture('<row r="1"><c r="A1"><f>A2</f></c></row>'),fixture(''))
    report=compare(digest(fixture('<row r="1"><c r="A1"><f>A2</f></c></row>')),digest(fixture('')))
    assert report['formulaMetadataDifferences'];tests+=1
    print(json.dumps({'selfTests':tests,'ok':True}))

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--left');parser.add_argument('--right');parser.add_argument('--out');parser.add_argument('--self-test',action='store_true');args=parser.parse_args()
    if args.self_test:self_test()
    else:
        if not all((args.left,args.right,args.out)):parser.error('--left --right --out are required')
        target=os.path.normcase(os.path.realpath(args.out))
        if any(os.path.normcase(os.path.realpath(p))==target for p in (args.left,args.right)):parser.error('입력과 결과 경로가 같을 수 없습니다.')
        before=[os.stat(p) for p in (args.left,args.right)]
        result=compare(digest(args.left),digest(args.right))
        after=[os.stat(p) for p in (args.left,args.right)]
        result['sourcesUnchanged']=all(a.st_size==b.st_size and a.st_mtime_ns==b.st_mtime_ns for a,b in zip(before,after))
        with open(args.out,'w',encoding='utf-8') as f:json.dump(result,f,ensure_ascii=False,indent=2)
        print(json.dumps({'ok':result['ok'],'differences':result['differences'],'formulaMetadataDifferences':result['formulaMetadataDifferences'],'sourcesUnchanged':result['sourcesUnchanged'],'leftMs':result['left']['elapsedMs'],'rightMs':result['right']['elapsedMs']},ensure_ascii=False))
        raise SystemExit(0 if result['ok'] and result['sourcesUnchanged'] else 1)
