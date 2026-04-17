import fs from 'node:fs/promises';
import path from 'node:path';
import { parsePptxProfileInput } from './src/lib/pptProfileParser.js';

const dir = '/home/user/sample-ppts';
const names = (await fs.readdir(dir)).filter((name) => name.endsWith('.pptx')).sort();
const results = [];

for (const name of names) {
  const filePath = path.join(dir, name);
  const buf = await fs.readFile(filePath);
  const parsed = await parsePptxProfileInput(buf, name);
  results.push(parsed);
}

console.log(JSON.stringify(results, null, 2));
