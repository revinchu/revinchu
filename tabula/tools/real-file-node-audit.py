"""Run each private workbook in its own bounded Node process (Windows, no dependencies).
Example: python tools/real-file-node-audit.py --manifest D:/private/manifest.json --out D:/private/results --mode audit F20
Manifest: [{"id":"F20","path":"D:/private/source.xlsx"}]. Originals are read-only.
Only anonymous IDs and numeric status appear on stdout; detailed private diagnostics stay under --out.
"""
from pathlib import Path
import argparse, ctypes, datetime, json, os, subprocess, sys, time
from ctypes import wintypes

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--manifest',default=os.environ.get('WIXEL_AUDIT_MANIFEST'))
    parser.add_argument('--out',default=os.environ.get('WIXEL_NODE_AUDIT_OUT'))
    parser.add_argument('--mode',choices=['inventory','pivots','formulas','audit'],default='audit')
    parser.add_argument('--reader',choices=['sync','async'],default=os.environ.get('WIXEL_AUDIT_READER','async'))
    parser.add_argument('--heap-mib',type=int,default=2048)
    parser.add_argument('--rss-mib',type=int,default=3072)
    parser.add_argument('--timeout',type=float,default=600)
    parser.add_argument('--min-free-mib',type=int,default=1024,help='Stop the child if system available physical memory falls below this value.')
    parser.add_argument('--node',default=os.environ.get('WIXEL_NODE_BINARY','node'))
    parser.add_argument('ids',nargs='*')
    args=parser.parse_args()
    if not args.manifest or not args.out:parser.error('--manifest와 --out을 지정하세요. 업무 파일/결과는 저장소 밖 D: 경로를 사용하세요.')
    if sys.platform!='win32':parser.error('이 실행기는 Windows 작업 집합 제한을 사용합니다. 다른 운영체제에서는 .mjs 단일 파일 도구를 이용하세요.')
    if args.heap_mib<64 or args.rss_mib<64 or args.timeout<=0 or args.min_free_mib<0:parser.error('메모리/시간 제한이 올바르지 않습니다.')
    root=Path(__file__).resolve().parent.parent;out=Path(args.out).resolve();out.mkdir(parents=True,exist_ok=True)
    manifest=Path(args.manifest).resolve();entries=json.loads(manifest.read_text(encoding='utf-8-sig'))
    entries=entries if isinstance(entries,list) else entries['files']
    import re
    if any(not re.fullmatch(r'[A-Za-z0-9_-]{1,64}',str(e.get('id',''))) for e in entries):parser.error('ID는 경로/이름 대신 영문·숫자·밑줄·하이픈만 사용하세요.')
    known={str(e['id']) for e in entries}
    if len(known)!=len(entries) or set(args.ids)-known:parser.error('중복되거나 찾을 수 없는 ID가 있습니다.')
    class Memory(ctypes.Structure):
        _fields_=[('cb',wintypes.DWORD),('PageFaultCount',wintypes.DWORD),('PeakWorkingSetSize',ctypes.c_size_t),('WorkingSetSize',ctypes.c_size_t),('QuotaPeakPagedPoolUsage',ctypes.c_size_t),('QuotaPagedPoolUsage',ctypes.c_size_t),('QuotaPeakNonPagedPoolUsage',ctypes.c_size_t),('QuotaNonPagedPoolUsage',ctypes.c_size_t),('PagefileUsage',ctypes.c_size_t),('PeakPagefileUsage',ctypes.c_size_t),('PrivateUsage',ctypes.c_size_t)]
    class SystemMemory(ctypes.Structure):
        _fields_=[('dwLength',wintypes.DWORD),('dwMemoryLoad',wintypes.DWORD),('ullTotalPhys',ctypes.c_ulonglong),('ullAvailPhys',ctypes.c_ulonglong),('ullTotalPageFile',ctypes.c_ulonglong),('ullAvailPageFile',ctypes.c_ulonglong),('ullTotalVirtual',ctypes.c_ulonglong),('ullAvailVirtual',ctypes.c_ulonglong),('ullAvailExtendedVirtual',ctypes.c_ulonglong)]
    ctypes.windll.kernel32.GlobalMemoryStatusEx.argtypes=[ctypes.POINTER(SystemMemory)]
    ctypes.windll.kernel32.OpenProcess.restype=wintypes.HANDLE
    ctypes.windll.psapi.GetProcessMemoryInfo.argtypes=[wintypes.HANDLE,ctypes.POINTER(Memory),wintypes.DWORD]
    ctypes.windll.kernel32.CloseHandle.argtypes=[wintypes.HANDLE]
    results=[];wanted=set(args.ids)
    for entry in entries:
        ident=str(entry['id'])
        if wanted and ident not in wanted:continue
        base=out/(ident+'-'+args.mode);started=time.monotonic()
        command=[args.node,'--max-old-space-size='+str(args.heap_mib),str(root/'tools/real-file-node-audit.mjs'),str(manifest),ident,args.mode,str(base)+'.json',args.reader]
        status={'id':ident,'mode':args.mode,'reader':args.reader,'started':datetime.datetime.now(datetime.timezone.utc).isoformat(),'heapLimitMiB':args.heap_mib,'workingSetLimitMiB':args.rss_mib,'minimumSystemFreeMiB':args.min_free_mib,'peakWorkingSetBytes':0,'peakPrivateBytes':0}
        proc=None;handle=None;cancelled=False;pause_seconds=0;last_poll=started
        with Path(str(base)+'.log').open('w',encoding='utf-8') as log:
            try:
                proc=subprocess.Popen(command,cwd=root,stdout=log,stderr=subprocess.STDOUT)
                handle=ctypes.windll.kernel32.OpenProcess(0x0400|0x0010,False,proc.pid)
                if not handle:raise RuntimeError('자식 프로세스 메모리 모니터를 열지 못했습니다.')
                while proc.poll() is None:
                    now=time.monotonic()
                    waiting=False
                    try:
                        checkpoint=Path(str(base)+'.json')
                        if checkpoint.stat().st_size<512:waiting=json.loads(checkpoint.read_text(encoding='utf-8-sig')).get('phase')=='waiting-slot'
                    except (OSError,ValueError):pass
                    if waiting:pause_seconds+=now-last_poll
                    last_poll=now;status['waitingSlotSeconds']=round(pause_seconds,3)
                    system=SystemMemory();system.dwLength=ctypes.sizeof(system)
                    if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(system)):
                        status['minimumSystemFreeBytes']=min(status.get('minimumSystemFreeBytes',system.ullAvailPhys),system.ullAvailPhys)
                        if system.ullAvailPhys<args.min_free_mib*1024**2:status['state']='system_memory_limit';proc.kill();break
                    mem=Memory();mem.cb=ctypes.sizeof(mem)
                    if ctypes.windll.psapi.GetProcessMemoryInfo(handle,ctypes.byref(mem),mem.cb):
                        status['peakWorkingSetBytes']=max(status['peakWorkingSetBytes'],mem.PeakWorkingSetSize)
                        status['peakPrivateBytes']=max(status['peakPrivateBytes'],mem.PrivateUsage)
                        if mem.WorkingSetSize>args.rss_mib*1024**2:status['state']='memory_limit';proc.kill();break
                    if time.monotonic()-started-pause_seconds>args.timeout:status['state']='timeout';proc.kill();break
                    (out/(ident+'-'+args.mode+'-progress.json')).write_text(json.dumps({**status,'seconds':round(time.monotonic()-started,2)}),encoding='utf-8')
                    time.sleep(.25)
                status['exitCode']=proc.wait();status.setdefault('state','completed' if status['exitCode']==0 else 'failed')
            except KeyboardInterrupt:
                status['state']='cancelled';cancelled=True
            except Exception as error:
                status['state']='error';status['error']=type(error).__name__
            finally:
                if proc and proc.poll() is None:proc.kill();proc.wait()
                if handle:ctypes.windll.kernel32.CloseHandle(handle)
        status['seconds']=round(time.monotonic()-started,3);status['activeSeconds']=round(status['seconds']-pause_seconds,3);results.append(status)
        (out/('runs-'+args.mode+'.json')).write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps(status,ensure_ascii=False),flush=True)
        if cancelled:break
    return 1 if any(x['state']!='completed' for x in results) else 0
if __name__=='__main__':sys.exit(main())
