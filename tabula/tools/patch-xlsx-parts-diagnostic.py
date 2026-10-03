"""Replace diagnostic ZIP32 parts in a new copy, preserving all other compressed bytes."""
import io,json,struct,zipfile
from pathlib import Path

def central(raw,count):
 out=[];at=0
 for _ in range(count):
  if raw[at:at+4]!=b'PK\x01\x02':raise ValueError('Invalid central record')
  n,x,c=struct.unpack_from('<HHH',raw,at+28);end=at+46+n+x+c;out.append(bytearray(raw[at:end]));at=end
 return out

def patch_parts(source,output,replacements):
 source,output=Path(source).resolve(),Path(output).resolve()
 if source==output or output.exists():raise ValueError('Output must be new')
 before=source.stat()
 with zipfile.ZipFile(source) as z:
  infos=z.infolist();order=sorted(infos,key=lambda x:x.header_offset)
  if len(infos)>=65535 or order[0].header_offset!=0 or len({i.filename for i in infos})!=len(infos):raise ValueError('Unique ZIP32 entries required')
  if not set(replacements).issubset({i.filename for i in infos}):raise ValueError('Unknown replacement')
  with source.open('rb') as f:f.seek(z.start_dir);records=central(f.read(),len(infos))
  original=dict(zip((i.filename for i in infos),records));offsets={};changed={}
  with source.open('rb') as src,output.open('xb') as dst:
   for index,info in enumerate(order):
    offsets[info.filename]=dst.tell()
    if info.filename in replacements:
     stream=io.BytesIO()
     with zipfile.ZipFile(stream,'w',compression=zipfile.ZIP_DEFLATED) as one:one.writestr(info.filename,replacements[info.filename])
     raw=stream.getvalue()
     with zipfile.ZipFile(io.BytesIO(raw)) as one:dst.write(raw[:one.start_dir]);changed[info.filename]=central(raw[one.start_dir:],1)[0]
    else:
     end=order[index+1].header_offset if index+1<len(order) else z.start_dir;left=end-info.header_offset;src.seek(info.header_offset)
     while left:
      block=src.read(min(left,1<<20))
      if not block:raise ValueError('Truncated local entry')
      dst.write(block);left-=len(block)
   start=dst.tell()
   if start>=0xffffffff:raise ValueError('ZIP64 unsupported')
   for info in infos:
    rec=changed.get(info.filename,original[info.filename]);struct.pack_into('<I',rec,42,offsets[info.filename]);dst.write(rec)
   size=dst.tell()-start;dst.write(struct.pack('<4s4H2IH',b'PK\x05\x06',0,0,len(infos),len(infos),size,start,len(z.comment)));dst.write(z.comment)
 with zipfile.ZipFile(source) as a,zipfile.ZipFile(output) as b:
  differences=[i.filename for i in a.infolist() if (i.CRC,i.file_size)!=(b.getinfo(i.filename).CRC,b.getinfo(i.filename).file_size)]
  for n,v in replacements.items():
   if b.read(n)!=v:raise ValueError('Replacement verification failed')
  if set(differences)-set(replacements):raise ValueError('Unexpected changed part')
 after=source.stat()
 if (before.st_size,before.st_mtime_ns)!=(after.st_size,after.st_mtime_ns):raise ValueError('Source changed')
 return {'diagnosticOnly':True,'sourceUnchanged':True,'changedParts':differences,'outputBytes':output.stat().st_size}

if __name__=='__main__':
 import argparse
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('source');p.add_argument('output');p.add_argument('part');p.add_argument('replacement');a=p.parse_args()
 print(json.dumps(patch_parts(a.source,a.output,{a.part:Path(a.replacement).read_bytes()})))
