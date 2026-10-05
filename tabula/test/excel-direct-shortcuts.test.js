import { test } from 'node:test';
import assert from 'node:assert/strict';
import { excelDirectCommand } from '../src/keyboard-shortcuts.js';
const keys = [
  [{key:'F1',altKey:true,shiftKey:true},'addSheet'],
  [{key:'F1',ctrlKey:true,shiftKey:true},'fullScreen'],
  [{key:'ㅎ',code:'KeyG',ctrlKey:true,shiftKey:true},'workbookStats'],
  [{key:'S',code:'KeyS',ctrlKey:true,shiftKey:true},'saveAs'],
  [{key:'F12',ctrlKey:true},'open'],
  [{key:'F12',shiftKey:true},'save'],
  [{key:'F12',ctrlKey:true,shiftKey:true},'print'],
  [{key:'=',code:'Equal',ctrlKey:true,altKey:true},'zoomIn'],
  [{key:'-',code:'Minus',ctrlKey:true,altKey:true},'zoomOut'],
  [{key:'+',code:'NumpadAdd',ctrlKey:true,altKey:true},'zoomIn'],
  [{key:'-',code:'NumpadSubtract',ctrlKey:true,altKey:true},'zoomOut'],
];
test('exact Excel shortcut modifiers and physical Korean keys route to their own commands',()=>{
  for(const [event,cmd] of keys) {
    assert.equal(excelDirectCommand(event),cmd);
    if(event.ctrlKey)assert.equal(excelDirectCommand({...event,ctrlKey:false,metaKey:true}),cmd);
  }
});
test('composition, AltGraph and unrelated modifier combinations do not borrow commands',()=>{
  for(const [event] of keys)for(const extra of [{isComposing:true},{keyCode:229},{getModifierState:k=>k==='AltGraph'}])assert.equal(excelDirectCommand({...event,...extra}),'');
  for(const event of [{key:'F1',ctrlKey:true},{key:'F1',altKey:true},{key:'G',code:'KeyG',ctrlKey:true},{key:'F2',ctrlKey:true,shiftKey:true},{key:'S',code:'KeyS',ctrlKey:true,shiftKey:true,altKey:true},{key:'=',code:'Equal',ctrlKey:true}])assert.equal(excelDirectCommand(event),'');
});
