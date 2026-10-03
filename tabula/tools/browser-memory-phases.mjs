// 검수용 관찰만 추가한다. 제품 API·데이터·쓰기 권한은 바꾸지 않는다.
export async function installBrowserMemoryPhases(page, events) {
 page.on('console',message=>{const text=message.text();if(!text.startsWith('__WIXEL_MEMORY__'))return;try{const event=JSON.parse(text.slice('__WIXEL_MEMORY__'.length));events.push(event);console.log(JSON.stringify({stage:'memory-phase',...event}));}catch{}});
 try {
  const cdp=await page.context().newCDPSession(page);await cdp.send('Performance.enable');
  let pending=false,last=0;
  const timer=setInterval(async()=>{if(pending)return;pending=true;try{const metrics=Object.fromEntries((await cdp.send('Performance.getMetrics')).metrics.map(m=>[m.name,m.value]));const heap=metrics.JSHeapUsedSize;if(Math.abs(heap-last)>=64*1024*1024){last=heap;const event={name:'cdp',state:'sample',at:Date.now(),heap,documents:metrics.Documents,nodes:metrics.Nodes};events.push(event);console.log(JSON.stringify({stage:'memory-phase',...event}));}}catch{}finally{pending=false;}},1000);
  page.once('close',()=>clearInterval(timer));
 }catch{} // WebKit은 CDP가 없으며 단계 시각·OS 메모리 보호는 그대로 수집한다.
 await page.evaluate(()=>{
  const phase=(name,state,extra={})=>console.log('__WIXEL_MEMORY__'+JSON.stringify({name,state,at:Math.round(performance.now()),heap:performance.memory?.usedJSHeapSize??null,...extra}));
  const proto=Object.getPrototypeOf(tabula.wb());
  for(const name of ['prepareGraph','buildGraphSteps','loadAsync','serializeBlob']){
   const original=proto[name];if(typeof original!=='function')continue;
   proto[name]=function(...args){phase(name,'start');try{const result=original.apply(this,args);if(result&&typeof result.then==='function')return result.then(value=>{phase(name,'end');return value;},error=>{phase(name,'error');throw error;});phase(name,'end');return result;}catch(error){phase(name,'error');throw error;}};
  }
  for(const name of ['restoreSteps','cellRunChunks']){
   const original=proto[name];if(typeof original!=='function')continue;
   proto[name]=function*(...args){phase(name,'start',name==='cellRunChunks'?{sheet:args[0]}:{});try{return yield* original.apply(this,args);}finally{phase(name,'end',name==='cellRunChunks'?{sheet:args[0]}:{});}};
  }
  const view=tabula.gv(),viewProto=Object.getPrototypeOf(view),model=view.host.slicerModel;
  const slicers=new WeakMap();let nextSlicer=0;
  const slicerInfo=sl=>{let info=slicers.get(sl);if(!info){info={number:++nextSlicer,count:null};slicers.set(sl,info);}return info;};
  if(typeof model==='function')view.host.slicerModel=function(sl,...args){const started=performance.now(),result=model.call(this,sl,...args),info=slicerInfo(sl),count=result?.items?.length??0;if(info.count!==count){info.count=count;phase('slicerModel','end',{slicer:info.number,items:count,ms:Math.round(performance.now()-started)});}return result;};
  const originalSlicerHtml=viewProto.slicerHtml;
  if(typeof originalSlicerHtml==='function')viewProto.slicerHtml=function(sl,...args){const info=slicerInfo(sl),started=performance.now();phase('slicerHtml','start',{slicer:info.number,items:info.count});try{const result=originalSlicerHtml.call(this,sl,...args);phase('slicerHtml','end',{slicer:info.number,items:info.count,ms:Math.round(performance.now()-started),htmlLength:result.length});return result;}catch(error){phase('slicerHtml','error',{slicer:info.number});throw error;}};
  const originalPut=IDBObjectStore.prototype.put;
  IDBObjectStore.prototype.put=function(value,key){
   const bytes=ArrayBuffer.isView(value)?value.byteLength:0;
   const label=bytes>1048576?'idb.array':Array.isArray(value?.chunks)?'idb.sheet':typeof key==='string'&&key.includes('#pivot-cache')?'idb.pivot-cache':null;
   if(label)phase(label,'start',{bytes});
   const request=originalPut.apply(this,arguments);
   if(label){request.addEventListener('success',()=>phase(label,'end',{bytes}),{once:true});request.addEventListener('error',()=>phase(label,'error',{bytes}),{once:true});}
   return request;
  };
  let previous=performance.memory?.usedJSHeapSize??0,count=0;
  setInterval(()=>{const current=performance.memory?.usedJSHeapSize??0;if(Math.abs(current-previous)>=64*1024*1024&&count++<200){previous=current;phase('timer','sample');}},1000);
  phase('probe','installed');
 });
}
