import { appendFileSync } from 'node:fs';
const origin = new URL(process.env.WIXEL_URL).origin;
const log = (item) => { if (process.env.WIXEL_AUDIT_GUARD_LOG) appendFileSync(process.env.WIXEL_AUDIT_GUARD_LOG, JSON.stringify({ time:Date.now(), ...item })+'\n'); };
const allowed = (address, method = 'GET') => { const u = new URL(String(address)); return ['GET','HEAD','OPTIONS'].includes(method.toUpperCase()) && u.origin===origin && !/^\/api(?:\/|$)/.test(u.pathname); };
const fetch = globalThis.fetch;
globalThis.fetch = (input, init={}) => { const address=typeof input==='object'&&input.url?input.url:String(input),method=init.method??input?.method??'GET'; if(!allowed(address,method)){log({kind:'node-fetch-blocked',url:address,method});return Promise.reject(new Error('감사 격리: 실제 외부/API/쓰기 요청 차단'));}return fetch(input,init); };
