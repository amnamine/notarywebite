const fs = require('fs');
const path = require('path');
const baseDir = path.resolve(__dirname, '..');
const JSZip = require(path.join(baseDir, 'static', 'js', 'jszip.min.js'));
global.window = global;
require(path.join(baseDir, 'static', 'js', 'core.js'));
require(path.join(baseDir, 'static', 'js', 'doc2docx.js'));

const shells = JSON.parse(fs.readFileSync(path.join(baseDir, 'static', 'shells.json'), 'utf8'));

const [,, inPath, outPath, paramsFile] = process.argv;
const params = paramsFile ? JSON.parse(fs.readFileSync(paramsFile, 'utf8')) : {};
let buf = fs.readFileSync(inPath);

async function run() {
  if (inPath.toLowerCase().endsWith('.doc')) {
    buf = (await DOC2DOCX.docToDocx(JSZip, buf)).buffer;
  }
  const analysed = await EDD.analyse(JSZip, buf);
  const result = await EDD.convert(JSZip, analysed, { tpl: shells.tpl }, {
    ...params,
    pageLayout: true
  });
  fs.writeFileSync(outPath, Buffer.from(result.template));
  console.log(JSON.stringify({ success: true, stats: result.stats, box: analysed.box }));
}

run().catch(e => {
  console.error(JSON.stringify({ success: false, error: e.message || String(e) }));
  process.exit(1);
});
