#!/usr/bin/env python3
"""Read only small pivot/slicer XML definitions; never read worksheet or cacheRecords bodies."""
import argparse,json,os,posixpath,re,zipfile
from xml.etree import ElementTree as ET
MAX_PART=8*1024*1024
local=lambda tag:tag.rsplit('}',1)[-1]
def kids(el,name):return [] if el is None else [x for x in el if local(x.tag)==name]
def child(el,name):return next(iter(kids(el,name)),None)
def attr(el,name,default=None):return next((v for k,v in el.attrib.items() if local(k)==name),default)
def audit(path):
 out={'scope':'Small pivot/cache/slicer definition count and index bounds; subtotal/grand row/column item semantics excluded','checks':0,'errors':[],'reviews':[],'caches':[],'pivots':0,'slicers':0}
 with zipfile.ZipFile(path) as z:
  def xml(name):
   info=z.getinfo(name)
   if info.file_size>MAX_PART:raise ValueError('Definition XML exceeds fixed small-part limit: '+name)
   with z.open(name) as f:return ET.fromstring(f.read(MAX_PART+1))
  def check(ok,part,kind,**meta):
   out['checks']+=1
   if not ok:out['errors'].append({'part':part,'kind':kind,**meta})
  def count(el,part,field=None):
   if el is not None and 'count' in el.attrib:check(int(el.attrib['count'])==len(el),part,'declared-count',container=local(el.tag),field=field,declared=int(el.attrib['count']),actual=len(el))
  wb=xml('xl/workbook.xml');rels={e.attrib['Id']:posixpath.normpath(posixpath.join('xl',e.attrib['Target'])) for e in xml('xl/_rels/workbook.xml.rels')}
  cache_ids={};ext_ids={}
  for link in kids(child(wb,'pivotCaches'),'pivotCache'):
   cid=int(link.attrib['cacheId']);part=rels[attr(link,'id')];root=xml(part);fs=kids(child(root,'cacheFields'),'cacheField');count(child(root,'cacheFields'),part)
   item_counts=[];field_names=[]
   for i,field in enumerate(fs):
    shared=child(field,'sharedItems');group=child(child(field,'fieldGroup'),'groupItems');count(shared,part,i);count(group,part,i)
    item_counts.append(len(group) if group is not None else len(shared) if shared is not None else 0);field_names.append(field.attrib.get('name',''))
    g=child(field,'fieldGroup')
    if g is not None:
     for k in ('base','par'):
      if k in g.attrib:check(0<=int(g.attrib[k])<len(fs),part,'group-field-index',field=i,attribute=k,index=int(g.attrib[k]))
     disc=child(g,'discretePr');count(disc,part,i)
     for x in kids(disc,'x'):check(0<=int(x.attrib.get('v',0))<item_counts[-1],part,'discrete-group-index',field=i,index=int(x.attrib.get('v',0)))
   meta={'part':part,'fieldCount':len(fs),'itemCounts':item_counts,'names':field_names};cache_ids[cid]=meta
   for e in root.iter():
    if local(e.tag)=='pivotCacheDefinition' and 'pivotCacheId' in e.attrib:ext_ids[int(e.attrib['pivotCacheId'])]=meta
   out['caches'].append({'part':part,'fieldCount':len(fs),'itemCounts':item_counts})
  for name in z.namelist():
   if not re.fullmatch(r'xl/pivotTables/pivotTable\d+\.xml',name):continue
   out['pivots']+=1;root=xml(name);cid=int(root.attrib['cacheId']);cache=cache_ids.get(cid);check(cache is not None,name,'cache-reference',index=cid)
   if not cache:continue
   fields=kids(child(root,'pivotFields'),'pivotField');count(child(root,'pivotFields'),name);check(len(fields)==cache['fieldCount'],name,'pivot-cache-field-count',actual=len(fields),expected=cache['fieldCount']);pf_counts=[]
   for i,pf in enumerate(fields):
    items=child(pf,'items');count(items,name,i);pf_counts.append(len(items) if items is not None else 0)
    for it in kids(items,'item'):
     if 'x' in it.attrib:check(i<len(cache['itemCounts']) and 0<=int(it.attrib['x'])<cache['itemCounts'][i],name,'pivot-item-cache-index',field=i,index=int(it.attrib['x']))
   data=kids(child(root,'dataFields'),'dataField');count(child(root,'dataFields'),name)
   for i,d in enumerate(data):check(0<=int(d.attrib['fld'])<len(fields),name,'value-field-index',field=i,index=int(d.attrib['fld']))
   for tag in ('rowFields','colFields'):
    count(child(root,tag),name)
    for f in kids(child(root,tag),'field'):check(int(f.attrib['x'])==-2 or 0<=int(f.attrib['x'])<len(fields),name,'axis-field-index',axis=tag,index=int(f.attrib['x']))
   for d in data:
    if 'baseField' in d.attrib:
     field=int(d.attrib['baseField']);check(0<=field<len(fields),name,'base-field-index',field=field)
     if 'baseItem' in d.attrib:
      value=int(d.attrib['baseItem'])
      check(value in (1048828,1048829) or 0<=field<len(pf_counts) and 0<=value<pf_counts[field],name,'base-item-index',field=field,index=value)
   filters=child(root,'filters');count(filters,name)
   for flt in kids(filters,'filter'):
    if 'fld' in flt.attrib:check(0<=int(flt.attrib['fld'])<len(fields),name,'filter-field-index',field=int(flt.attrib['fld']))
   pages=child(root,'pageFields');count(pages,name)
   for i,p in enumerate(kids(pages,'pageField')):
    field=int(p.attrib['fld']);check(0<=field<len(fields),name,'page-field-index',field=i,index=field)
    if 'item' in p.attrib and 0<=field<len(fields):check(0<=int(p.attrib['item'])<pf_counts[field],name,'page-item-index',field=field,index=int(p.attrib['item']))
   for axis in ('row','col'):
    fs=[int(e.attrib['x']) for e in kids(child(root,axis+'Fields'),'field')];items=child(root,axis+'Items');count(items,name)
    for row in kids(items,'i'):
     if 'i' in row.attrib:check(0<=int(row.attrib['i'])<len(data),name,'axis-value-index',axis=axis,index=int(row.attrib['i']))
     if row.attrib.get('t','data')!='data':continue
     depth=int(row.attrib.get('r',0));xs=kids(row,'x')
     if not fs and xs:continue
     for offset,x in enumerate(xs):
      if depth+offset>=len(fs):out['reviews'].append({'part':name,'kind':'axis-value-position-outside-declared-fields','axis':axis,'depth':depth,'offset':offset});continue
      field=fs[depth+offset];limit=len(data) if field==-2 else pf_counts[field] if 0<=field<len(pf_counts) else 0
      check(0<=int(x.attrib.get('v',0))<limit,name,'axis-item-index',axis=axis,field=field,index=int(x.attrib.get('v',0)))
  for name in z.namelist():
   if not re.fullmatch(r'xl/slicerCaches/slicerCache\d+\.xml',name):continue
   out['slicers']+=1;root=xml(name);tab=next((e for e in root.iter() if local(e.tag)=='tabular'),None)
   if tab is None:continue
   cid=int(tab.attrib['pivotCacheId']);cache=ext_ids.get(cid) or cache_ids.get(cid);check(cache is not None,name,'slicer-cache-reference',index=cid)
   if not cache:continue
   matches=[i for i,n in enumerate(cache['names']) if n.casefold()==root.attrib.get('sourceName','').casefold()];check(len(matches)==1,name,'slicer-source-field-count',actual=len(matches))
   if len(matches)!=1:continue
   field=matches[0];items=child(tab,'items');count(items,name,field)
   for item in kids(items,'i'):check(0<=int(item.attrib['x'])<cache['itemCounts'][field],name,'slicer-item-index',field=field,index=int(item.attrib['x']))
 out['ok']=not out['errors'];return out
if __name__=='__main__':
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('source');p.add_argument('--out',required=True);a=p.parse_args()
 if os.path.normcase(os.path.realpath(a.source))==os.path.normcase(os.path.realpath(a.out)):p.error('입력과 결과 경로가 같을 수 없습니다.')
 before=os.stat(a.source);r=audit(a.source);after=os.stat(a.source);r['sourceUnchanged']=before.st_size==after.st_size and before.st_mtime_ns==after.st_mtime_ns
 with open(a.out,'w',encoding='utf-8') as f:json.dump(r,f,ensure_ascii=False,indent=2)
 print(json.dumps({k:r[k] for k in ['ok','checks','pivots','slicers','errors','reviews']},ensure_ascii=False))
 raise SystemExit(0 if r['ok'] else 1)
