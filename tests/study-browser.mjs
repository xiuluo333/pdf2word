// Run with an isolated Chrome profile listening on localhost:9224.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
const pages = await (await fetch('http://127.0.0.1:9224/json')).json();
const ws = new WebSocket(pages.find(p => p.type === 'page').webSocketDebuggerUrl);
await new Promise(resolve => ws.addEventListener('open', resolve, { once: true }));
let next = 0;
const pending = new Map();
ws.addEventListener('message', e => {
  const message = JSON.parse(e.data);
  if (pending.has(message.id)) {
    const [resolve, reject] = pending.get(message.id); pending.delete(message.id);
    if (message.error) reject(message.error); else resolve(message.result);
  }
});
function send(method, params = {}) {
  return new Promise((resolve, reject) => { const id = ++next; pending.set(id, [resolve, reject]); ws.send(JSON.stringify({id, method, params})); });
}
async function evaluate(expression) {
  const result = await send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
  if (result.exceptionDetails) throw new Error(JSON.stringify(result.exceptionDetails));
  return result.result.value;
}
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vocab-study-test-'));
const file = path.join(dir, 'reader.html');
const template = fs.readFileSync(new URL('../generate_html.py', import.meta.url), 'utf8').split("HTML_TEMPLATE = r'''")[1].split("'''\n")[0];
const rows = Array.from({ length: 8000 }, (_, i) => ({ id: `W${i}`, word: `word${i}`, translate: '中文释义', english_definition: 'an English definition', example_sentence: 'We enjoyed the voyage. (a note)' }));
fs.writeFileSync(file, template.replace('__VOCAB_JSON__', JSON.stringify(rows)));
await send('Page.enable');
async function navigate() {
  const loaded = new Promise((resolve, reject) => {
    const timer = setTimeout(() => { ws.removeEventListener('message', handler); reject(new Error('navigation timeout')); }, 15000);
    function handler(event) {
      if (JSON.parse(event.data).method === 'Page.loadEventFired') { clearTimeout(timer); ws.removeEventListener('message', handler); resolve(); }
    }
    ws.addEventListener('message', handler);
  });
  await send('Page.navigate', { url: pathToFileURL(file).href });
  await loaded;
  assert.equal(await evaluate("!!document.querySelector('[data-review]')"),true);
}
try {
  await navigate();
  await evaluate("localStorage.clear()");
  await navigate();
  assert.equal(await evaluate('document.querySelectorAll(".card").length'), 36);
  assert.equal(await evaluate('state.matches.length'), 8000);
  await evaluate(`document.querySelector('[data-review="W0"]').click()`);
  assert.equal(await evaluate('records.get("word0").review'), true);
  await evaluate(`document.querySelector('[data-flip="W0"]').click()`);
  assert.deepEqual(await evaluate(`(() => { const c = document.getElementById('word-card-W0'); return [c.querySelector('.card-front').hidden, c.querySelector('.card-back').hidden, c.querySelector('.card-back').textContent]; })()`), [true, false, 'word0We enjoyed the voyage. (a note)']);
  await evaluate(`document.getElementById('review-only').click()`);
  assert.equal(await evaluate('state.matches.length'), 1);
  assert.equal(await evaluate('document.querySelector(".card-front").hidden'), true);
  await evaluate(`document.querySelector('[data-review="W0"]').click()`);
  assert.equal(await evaluate('state.matches.length'), 0);
  assert.equal(await evaluate('document.getElementById("empty").classList.contains("visible")'), true);
  await evaluate(`document.getElementById('review-only').click(); locate('W7999')`);
  assert.equal(await evaluate('document.querySelectorAll(".card").length'), 8);
  await evaluate(`document.querySelector('[data-review="W7999"]').click()`);
  await navigate();
  assert.equal(await evaluate('isReview(byId.get("W7999"))'), true);
  assert.equal(await evaluate('state.lastId'), 'W7999');
  assert.equal(await evaluate('state.page'), 222);
  // Reordering IDs must not move the mark or position to a different word.
  rows.reverse().forEach((row, i) => { row.id = `N${i}`; });
  fs.writeFileSync(file, template.replace('__VOCAB_JSON__', JSON.stringify(rows)));
  await navigate();
  assert.equal(await evaluate('state.lastId'), 'N0');
  assert.equal(await evaluate('isReview(byId.get("N0"))'), true);
  await evaluate(`localStorage.setItem(recordKey('word7999'), '{broken')`);
  await navigate();
  assert.equal(await evaluate('isReview(byId.get("N0"))'), true);
  assert.equal(await evaluate(`document.getElementById('storage-status').textContent.includes('备用')`), true);
  assert.equal(await evaluate(`(() => { const before=JSON.stringify(backupData().records); try { importData({format:'vocabulary-study',version:2, records:[{word:'x',review:true,at:1},{word:'y',review:'bad',at:1}],position:null}); } catch {} return before===JSON.stringify(backupData().records); })()`), true);
  await evaluate(`importData({format:'vocabulary-study',version:2,records:[{word:'word7999',review:false,at:Date.now()+10}],position:null})`);
  assert.equal(await evaluate('isReview(byId.get("N0"))'), false);
  await evaluate(`importData({format:'vocabulary-study',version:2,records:[{word:'word7999',review:true,at:1}],position:null})`);
  assert.equal(await evaluate('isReview(byId.get("N0"))'), false);
  // Export is a real download; import restores data after browser storage is cleared.
  await send('Browser.setDownloadBehavior', {behavior:'allow', downloadPath:dir});
  await evaluate(`document.getElementById('export-study').click()`);
  let backupFile;
  for(let i=0;i<100;i++) {
    const candidate=fs.readdirSync(dir).find(name=>name.endsWith('.json'));
    if(candidate) { backupFile=path.join(dir,candidate); break; }
    await new Promise(r=>setTimeout(r,30));
  }
  assert.ok(backupFile, 'backup download completed');
  const backup=JSON.parse(fs.readFileSync(backupFile,'utf8'));
  assert.equal(backup.format,'vocabulary-study');
  await evaluate('localStorage.clear()');
  await navigate();
  const documentNode=await send('DOM.getDocument');
  const fileInput=await send('DOM.querySelector',{nodeId:documentNode.root.nodeId,selector:'#import-file'});
  await send('DOM.setFileInputFiles',{nodeId:fileInput.nodeId,files:[backupFile]});
  for(let i=0;i<100;i++) {
    if(await evaluate('records.has("word7999")')) break;
    await new Promise(r=>setTimeout(r,30));
  }
  assert.equal(await evaluate('records.get("word7999").review'),false);
  // Persist all 8000 marks and reload, measuring the storage volume, not only DOM size.
  await evaluate(`importData({format:'vocabulary-study',version:2,records:VOCAB.map(item=>({word:item.word,review:true,at:Date.now()+100})),position:null})`);
  await navigate();
  assert.equal(await evaluate('records.size'),8000);
  assert.equal(await evaluate('VOCAB.filter(isReview).length'),8000);
  console.log('Stored 8000 marks with backups; localStorage entries:',await evaluate('localStorage.length'));
  // Set a known unmarked state before simulating a failed save.
  await evaluate(`records.set('word7999',{word:'word7999',review:false,at:Date.now()+200}); renderPage()`);
  await evaluate(`localStorage.setItem = () => { throw new Error('quota'); }; document.querySelector('[data-review="N0"]').click()`);
  assert.equal(await evaluate('isReview(byId.get("N0"))'), true);
  assert.equal(await evaluate('storageFailed'), true);
  assert.equal(await evaluate('backupData().records.find(r => r.word === "word7999").review'), true);
  // Live storage updates (another tab) merge one word, not the entire map.
  await evaluate(`window.dispatchEvent(new StorageEvent('storage', { key:recordKey('word123'),newValue:JSON.stringify({word:'word123',review:true,at:Date.now()}) }))`);
  assert.equal(await evaluate('records.get("word123").review'), true);
  await send('Emulation.setDeviceMetricsOverride', {width:375,height:812,deviceScaleFactor:1,mobile:true});
  assert.equal(await evaluate('document.documentElement.scrollWidth <= innerWidth'), true);
  console.log('PASS: 8000 words, bounded DOM, flip, review filter, reload, reordered IDs, backup recovery, import validation/merge, save failure, cross-tab updates, mobile layout');
} finally { ws.close(); }
