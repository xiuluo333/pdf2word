import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

for (const file of ['考研生词.html', 'generate_html.py']) {
  const html = fs.readFileSync(new URL('../' + file, import.meta.url), 'utf8');
  const script = html.match(/<script>([\s\S]*?)<\/script>/)[1].replace('__VOCAB_JSON__', '[]');
  new vm.Script(script);
  const spoken = [];
  const handlers = [];
  const voice = { lang: 'en-GB' };
  const context = vm.createContext({
    document: {
      getElementById: () => ({ value: 'en-GB', classList: { remove() {} } }),
      addEventListener: (name, handler) => handlers.push(handler),
    },
    window: { speechSynthesis: { cancel() {}, getVoices: () => [voice], speak: utterance => spoken.push(utterance) } },
    SpeechSynthesisUtterance: function (text) { this.text = text; },
  });
  vm.runInContext(script.split('// Shared by both card faces:')[0], context);
  const markup = vm.runInContext(`cardMarkup({ id: 'test', word: 'voyage', english_definition: 'a journey by sea', example_sentence: 'We enjoyed the voyage.' })`, context);
  assert.match(markup, /aria-label="朗读英译英"/);
  assert.match(markup, /aria-label="朗读日常例句"/);
  assert.equal((markup.match(/data-speak-example=/g) || []).length, 3);
  assert.doesNotMatch(vm.runInContext(`cardMarkup({ id: 'test', word: 'voyage' })`, context), /data-speak-example=/);
  assert.match(vm.runInContext(`cardMarkup({ id: 'test', word: 'voyage', english_definition: '<script>' })`, context), /&lt;script&gt;/);
  vm.runInContext('refreshVoices()', context);
  vm.runInContext(script.slice(script.indexOf("document.addEventListener('click'"), script.indexOf("document.addEventListener('dblclick'")), context);
  for (const text of ['a journey by sea', 'We enjoyed the voyage.']) {
    handlers[0]({ target: { closest: selector => selector === '[data-speak-example]' ? { dataset: { speakExample: text } } : null } });
    assert.equal(spoken.at(-1).text, text);
    assert.equal(spoken.at(-1).lang, 'en-GB');
    assert.equal(spoken.at(-1).voice, voice);
  }
  for (const [shown, expected] of [
    ['a journey by sea (a trip on water)', 'a journey by sea'],
    ['We enjoyed the voyage. (我们享受了这次航行。)', 'We enjoyed the voyage.'],
    ['a journey by sea（a trip on water）.', 'a journey by sea.'],
    ['a journey by sea (note)（another note）  ', 'a journey by sea'],
    ['We (all of us) enjoyed the voyage.', 'We (all of us) enjoyed the voyage.'],
  ]) {
    handlers[0]({ target: { closest: selector => selector === '[data-speak-example]' ? { dataset: { speakExample: shown } } : null } });
    assert.equal(spoken.at(-1).text, expected);
  }
  console.log(`${file}: definition/example rendering and speech dispatch passed`);
}
