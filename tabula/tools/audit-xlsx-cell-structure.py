#!/usr/bin/env python3
"""Bounded-memory OOXML worksheet ordering and scalar-index checks, without content logging."""
import argparse, json, math, os, re, time, zipfile
from xml.parsers import expat
CHUNK = 64 * 1024

def parse_entry(archive, name, start=None, end=None, chars=None):
    parser=expat.ParserCreate(namespace_separator='}')
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    parser.ExternalEntityRefHandler=lambda *args: 0
    if start: parser.StartElementHandler=lambda tag,attrs: start(parser,tag.rsplit('}',1)[-1],attrs)
    if end: parser.EndElementHandler=lambda tag: end(parser,tag.rsplit('}',1)[-1])
    if chars: parser.CharacterDataHandler=lambda value: chars(parser,value)
    with archive.open(name) as stream:
        while chunk:=stream.read(CHUNK): parser.Parse(chunk,False)
    parser.Parse(b'',True)

def inspect(source):
    t=time.monotonic();out={'scope':'Worksheet ordering, cell style/shared-string indices, scalar values and finite row/column dimensions only; not complete OOXML schema validation','chunkBytes':CHUNK,'errors':[],'sheets':[]}
    with zipfile.ZipFile(source) as z:
        paths=set(z.namelist());counts={'styles':0,'strings':0};stack=[]
        def styles_start(p,tag,attrs):
            if tag=='xf' and stack and stack[-1]=='cellXfs': counts['styles']+=1
            stack.append(tag)
        def styles_end(p,tag): stack.pop()
        parse_entry(z,'xl/styles.xml',styles_start,styles_end)
        if 'xl/sharedStrings.xml' in paths:
            parse_entry(z,'xl/sharedStrings.xml',lambda p,tag,attrs:counts.__setitem__('strings',counts['strings']+(tag=='si')))
        out['indexCounts']=counts
        for name in sorted(paths):
            if not re.fullmatch(r'xl/worksheets/sheet\d+\.xml',name): continue
            state={'row':0,'lastRow':0,'lastCol':0,'cell':None,'value':None,'rows':0,'cells':0,'maxRow':0,'maxCol':0,'stylesChecked':0,'stringsChecked':0,'numbersChecked':0}
            def issue(parser,kind):
                if len(out['errors'])<100:out['errors'].append({'entry':name,'kind':kind,'byteOffset':parser.CurrentByteIndex})
            def start(parser,tag,attrs):
                if tag in ('row','col'):
                    for key in ('ht','width'):
                        if key in attrs:
                            try:
                                if not math.isfinite(float(attrs[key])):issue(parser,'nonfinite-dimension')
                            except ValueError:issue(parser,'invalid-dimension')
                if tag=='row':
                    row=int(attrs.get('r',state['lastRow']+1));state['rows']+=1
                    if not 1<=row<=1048576:issue(parser,'row-outside-excel-limit')
                    if row<=state['lastRow']:issue(parser,'row-duplicate-or-unsorted')
                    state['row']=state['lastRow']=row;state['lastCol']=0;state['maxRow']=max(state['maxRow'],row)
                elif tag=='c':
                    state['cells']+=1;state['cell']={'type':attrs.get('t','n')};match=re.fullmatch(r'([A-Z]+)([1-9][0-9]*)',attrs.get('r',''))
                    if not match:issue(parser,'invalid-cell-address')
                    else:
                        col=0
                        for letter in match[1]:col=26*col+ord(letter)-64
                        if col<=state['lastCol']:issue(parser,'cell-duplicate-or-unsorted')
                        if int(match[2])!=state['row']:issue(parser,'cell-row-mismatch')
                        if col>16384:issue(parser,'column-outside-excel-limit')
                        state['lastCol']=col;state['maxCol']=max(state['maxCol'],col)
                    if 's' in attrs:
                        state['stylesChecked']+=1
                        try:
                            if not 0<=int(attrs['s'])<counts['styles']:issue(parser,'cell-style-index-range')
                        except ValueError:issue(parser,'invalid-cell-style-index')
                elif tag=='v' and state['cell'] is not None:state['value']=''
            def chars(parser,value):
                if state['value'] is not None:
                    if len(state['value'])<256:state['value']+=value[:256-len(state['value'])]
            def end(parser,tag):
                if tag=='v' and state['cell'] is not None:
                    kind=state['cell']['type'];value=state['value'];state['value']=None
                    if value and kind=='s':
                        state['stringsChecked']+=1
                        try:
                            if not 0<=int(value)<counts['strings']:issue(parser,'shared-string-index-range')
                        except ValueError:issue(parser,'invalid-shared-string-index')
                    elif value and kind=='n':
                        state['numbersChecked']+=1
                        try:
                            if not math.isfinite(float(value)):issue(parser,'nonfinite-cell-number')
                        except ValueError:issue(parser,'invalid-cell-number')
                    elif value and kind=='b' and value not in ('0','1'):issue(parser,'invalid-cell-boolean')
                elif tag=='c':state['cell']=None;state['value']=None
            parse_entry(z,name,start,end,chars)
            out['sheets'].append({'entry':name,**{key:state[key] for key in ('rows','cells','maxRow','maxCol','stylesChecked','stringsChecked','numbersChecked')}})
    out['elapsedMs']=round((time.monotonic()-t)*1000);out['ok']=not out['errors'];return out

if __name__=='__main__':
    args=argparse.ArgumentParser(description=__doc__);args.add_argument('source');args.add_argument('--out',required=True);args=args.parse_args()
    if os.path.normcase(os.path.realpath(args.source))==os.path.normcase(os.path.realpath(args.out)):
        raise SystemExit('결과 경로는 입력 파일과 달라야 합니다.')
    before=os.stat(args.source)
    result=inspect(args.source)
    after=os.stat(args.source)
    result['sourceUnchanged']=before.st_size==after.st_size and before.st_mtime_ns==after.st_mtime_ns
    if not result['sourceUnchanged']:
        result['errors'].append({'kind':'source-changed-during-audit'});result['ok']=False
    with open(args.out,'w',encoding='utf-8') as f:json.dump(result,f,ensure_ascii=False,indent=2)
    print(json.dumps({key:result[key] for key in ('ok','indexCounts','errors','elapsedMs')},ensure_ascii=False))
    raise SystemExit(0 if result['ok'] else 1)
