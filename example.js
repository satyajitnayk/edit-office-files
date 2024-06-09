import {SearchAndReplace} from './search-and-replace.js';

async function main() {
  const searchTexts = ['Hello World', 'are You', 'coloured'];
  const replacementTexts = ['REPLACEMENT1', 'REPLACEMENT2', 'REPLACEMENT3'];
  const reader = new SearchAndReplace('assets/document.docx', searchTexts, replacementTexts, 'updated.docx');
  await reader.process();
}

main();
