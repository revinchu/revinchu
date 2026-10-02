// Persist remote video metadata, never executable markup or arbitrary URL schemes.
export const MEDIA_OBJECT_URI = '{46DC1CA8-1F4B-47D8-8EF7-574958454C56}';
export const VIDEO_POSTER = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
export function normalizeVideo(value) {
  if(value?.kind !== 'video' || typeof value.src !== 'string' || value.src.length>8192) return null;
  let url;try{url=new URL(value.src);}catch{return null;}
  if(!['https:','http:'].includes(url.protocol)||url.username||url.password)return null;
  const out={kind:'video',src:url.href};
  for(const key of ['title','source','page','credit','license','mime'])if(typeof value[key]==='string')out[key]=value[key].slice(0,key==='page'?8192:1000);
  return out;
}
