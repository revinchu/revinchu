// One pass over the inflated worksheet: retain its metadata and current row,
// never a buffer sized to the expanded ZIP entry. Prefixes and UTF-8 boundaries
// are preserved, including drawings/merges stored after sheetData.
function tagEnd(text) {
  if(text.startsWith('<!--')){const p=text.indexOf('-->');return p<0?-1:p+3;}
  if(text.startsWith('<![CDATA[')){const p=text.indexOf(']]>');return p<0?-1:p+3;}
  if(text.startsWith('<?')){const p=text.indexOf('?>');return p<0?-1:p+2;}
  let quote='';for(let p=1;p<text.length;p++){const c=text[p];if(quote){if(c===quote)quote='';}else if(c==='"'||c==="'")quote=c;else if(c==='>')return p+1;}return -1;
}
function xmlCursor(chunks) {
  const iterator=chunks[Symbol.iterator](),decoder=new TextDecoder();
  return {pending:'',done:false,
    more(){const s=iterator.next();this.done=!!s.done;this.pending+=s.done?decoder.decode():decoder.decode(s.value,{stream:true});},
    token(){
      while(!this.pending&&!this.done)this.more();if(!this.pending)return null;
      const open=this.pending.indexOf('<');if(open!==0){const n=open<0?this.pending.length:open,text=this.pending.slice(0,n);this.pending=this.pending.slice(n);return {text};}
      let end=tagEnd(this.pending);while(end<0&&!this.done){this.more();end=tagEnd(this.pending);}
      if(end<0)throw new Error('시트 XML 태그가 끝까지 저장되지 않았습니다.');
      const text=this.pending.slice(0,end);this.pending=this.pending.slice(end);
      const m=/^<(\/?)([\w.-]+:)?([\w.-]+)(?=[\s/>])/.exec(text);
      return m?{text,name:m[3],qualified:(m[2]??'')+m[3],closing:!!m[1],self:/\/\s*>$/.test(text)}:{text};
    },
    rowBody(name){
      const close='</'+name,boundary=Math.max(close.length+1,9);let search=0;
      for(;;){
        const end=this.pending.indexOf(close,search),declaration=this.pending.indexOf('<!',search),instruction=this.pending.indexOf('<?',search);
        const special=declaration<0?instruction:instruction<0?declaration:Math.min(declaration,instruction);
        if(special>=0&&(end<0||special<end)){
          const tail=this.pending.slice(special);
          if(!this.done&&['<![CDATA[','<!--','<?'].some(marker=>marker.startsWith(tail))){this.more();continue;}
          const cdata=this.pending.startsWith('<![CDATA[',special),comment=this.pending.startsWith('<!--',special),pi=this.pending.startsWith('<?',special);
          if(cdata||comment||pi){
            const marker=cdata?']]>':comment?'-->':'?>',p=this.pending.indexOf(marker,special+(cdata?9:comment?4:2));
            if(p>=0){search=p+marker.length;continue;}
            if(this.done)throw new Error('시트 행이 끝까지 저장되지 않았습니다.');
            this.more();continue;
          }
          search=special+2;continue;
        }
        if(end>=0){
          let p=end+close.length;
          while([32,9,10,13].includes(this.pending.charCodeAt(p)))p++;
          if(p<this.pending.length){
            if(this.pending[p]==='>'){const n=p+1,out=this.pending.slice(0,n);this.pending=this.pending.slice(n);return out;}
            search=end+close.length;continue;
          }
          if(this.done)throw new Error('시트 행이 끝까지 저장되지 않았습니다.');
          this.more();continue;
        }
        if(this.done)throw new Error('시트 행이 끝까지 저장되지 않았습니다.');
        // Recheck only an incomplete close or special delimiter across chunks.
        search=Math.max(search,this.pending.length-boundary);this.more();
      }
    },
    tail(){while(!this.done)this.more();const out=this.pending;this.pending='';return out;},
    close(){iterator.return?.();}
  };
}
export function sheetXmlStream(chunks) {
  const cursor=xmlCursor(chunks);let head='',worksheet='',container;
  try {
    for(;;){const t=cursor.token();if(!t)throw new Error('시트의 셀 데이터가 없습니다.');if(t.name==='worksheet'&&!t.closing)worksheet=t.qualified;
      if(t.name==='sheetData'&&!t.closing){container=t;break;}head+=t.text;}
    if(!worksheet)throw new Error('시트 XML 형식이 올바르지 않습니다.');
  } catch(e){cursor.close();throw e;}
  const metadataHead=head+'<'+container.qualified+'/>';let consumed=false;
  const result={head:metadataHead+'</'+worksheet+'>',prefix:container.qualified.includes(':'),
    *rows(){
      if(consumed)throw new Error('시트 데이터는 한 번만 읽을 수 있습니다.');consumed=true;
      try {
        if(!container.self)for(;;){const t=cursor.token();if(!t)throw new Error('시트 셀 데이터가 끝까지 저장되지 않았습니다.');
          if(t.name===container.name&&t.closing)break;
          if(t.name==='row'&&!t.closing)yield t.self?t.text:t.text+cursor.rowBody(t.qualified);
          else if(t.name||t.text.trim()&&!/^<(?:!|\?)/.test(t.text))throw new Error('시트 셀 데이터 형식이 올바르지 않습니다.');
        }
        result.rest=metadataHead+cursor.tail();
      } finally {cursor.close();}
    }};return result;
}
