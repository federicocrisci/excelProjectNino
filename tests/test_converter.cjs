/* Independent regression/functional tests. Run:
 * node tests/test_converter.cjs [original.xlsx] [reference.xls]
 * The private workbooks are not distributed with the application.
 */
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const os = require('node:os');
const XLSX = require('../vendor/xlsx.full.min.js');
const fflate = require('../vendor/fflate.js');
const root = path.resolve(__dirname, '..');

function filesUnder(dir) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir, {withFileTypes:true}).flatMap(ent => {
    const p = path.join(dir, ent.name);
    return ent.isDirectory() ? filesUnder(p) : [p];
  });
}
const attachments = filesUnder('/workspace/attachments');
const originalPath = process.argv[2] || attachments.find(p => path.basename(p) === '44,1 26-30ott.xlsx');
const referencePath = process.argv[3] || attachments.find(p => path.basename(p) === 'New Grieco N°44.xls');
if (!originalPath || !referencePath || !fs.existsSync(originalPath) || !fs.existsSync(referencePath)) {
  console.error('Fixture mancanti. Specificare il Canvas originale e il riferimento come argomenti. Nessun test eseguito.');
  process.exit(2);
}
for (const name of ['template.js','converter.js']) {
  if (!fs.existsSync(path.join(root,name))) {
    console.error(name + ' non ancora disponibile. Nessun test eseguito.');
    process.exit(2);
  }
}
const ctx = {
  XLSX, fflate, console, Uint8Array, ArrayBuffer, TextEncoder, TextDecoder,
  atob: s => Buffer.from(s,'base64').toString('binary'),
  btoa: s => Buffer.from(s,'binary').toString('base64'),
  setTimeout, clearTimeout,
};
ctx.window = ctx;
ctx.globalThis = ctx;
vm.createContext(ctx);
for (const name of ['template.js','converter.js']) vm.runInContext(fs.readFileSync(path.join(root,name),'utf8'), ctx, {filename:name});
const converter = ctx.CanvasConverter;
assert.ok(converter && typeof converter.analyze === 'function' && typeof converter.generate === 'function');
const readWorkbook = bytes => XLSX.read(bytes,{type:'array',cellFormula:true,cellStyles:true});
const originalBytes = fs.readFileSync(originalPath);
const referenceWorkbook = readWorkbook(fs.readFileSync(referencePath));
const reference = referenceWorkbook.Sheets[referenceWorkbook.SheetNames[0]];
const ab = bytes => bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength);
const value = cell => !cell || cell.v === null || cell.v === undefined ? '' : cell.v;
const str = cell => String(value(cell));
const cleanFormula = f => String(f || '').replace(/^=/,'').replace(/\$/g,'').replace(/\s+/g,'').toUpperCase();
const sheetOf = wb => wb.Sheets[wb.SheetNames[0]];
const productRows = sheet => {
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const rows=[];
  for(let r=1;r<=range.e.r+1;r++) if(value(sheet['C'+r]) !== '') rows.push(r);
  return rows;
};
const rowForCode = (sheet,code) => productRows(sheet).find(r=>str(sheet['C'+r])===String(code));
const output = async (analysis,stockMode='reference') => {
  const result=await converter.generate(analysis,{stockMode});
  const bytes=result && result.bytes ? result.bytes : result;
  assert.ok(bytes && bytes.byteLength>1000,'output must be a real workbook');
  return {bytes,sheet:sheetOf(readWorkbook(bytes))};
};
const mutated = edit => {
  const wb=readWorkbook(originalBytes);
  const s=sheetOf(wb);
  edit(s);
  return XLSX.write(wb,{bookType:'xlsx',type:'buffer'});
};
const put = (s,address,v) => {s[address]={t:typeof v==='number'?'n':'s',v};};
const analyzeBytes = bytes => converter.analyze(ab(bytes),'test.xlsx');
let passed=0;
async function test(name,run) {
  await run();
  passed++;
  console.log('PASS '+name);
}

(async()=>{
  const baseline=await analyzeBytes(originalBytes);
  const baselineOutput=await output(baseline,'reference');
  const generated=baselineOutput.sheet;

  await test('Baseline: all products and duplicate occurrences survive the intended selection',async()=>{
    assert.equal(baseline.stats.sourceProducts,6237);
    assert.equal(baseline.stats.outputProducts,6235);
    assert.equal(productRows(generated).length,6235);
    assert.equal(XLSX.utils.decode_range(generated['!ref']).e.r+1,6619);
    const frequencies=s=>{const m=new Map();for(const r of productRows(s)){const c=str(s['C'+r]);m.set(c,(m.get(c)||0)+1);}return m;};
    assert.deepEqual(frequencies(generated),frequencies(reference));
  });

  await test('Baseline: complete cell comparison to client reference, with documented corrections only',async()=>{
    let cells=0;
    for(let r=1;r<=6619;r++)for(let c=0;c<13;c++){
      const address=XLSX.utils.encode_cell({r:r-1,c});
      const expected=reference[address],actual=generated[address];
      if(c===12){
        if(r===2)assert.equal(cleanFormula(actual?.f),'SUM(M3:M6619)',address+' total includes final row');
        else if(expected?.f)assert.equal(cleanFormula(actual?.f),cleanFormula(expected.f),address+' line formula');
        else assert.equal(value(actual),value(expected),address+' blank/nonformula total cell');
      }else if(str(reference['C'+r])==='26794'&&(c===7||c===8)){
        assert.equal(value(actual),c===7?22.45:27.39,address+' uses source price instead of 224.45 typo');
      }else assert.equal(value(actual),value(expected),address);
      cells++;
    }
    console.log('  Compared '+cells+' cells, including section and detail rows.');
  });

  await test('All original EANs and article codes remain text',async()=>{
    for(const r of productRows(generated)){
      assert.equal(generated['C'+r].t,'s','article code at row '+r);
      if(value(generated['B'+r])!=='')assert.equal(generated['B'+r].t,'s','EAN at row '+r);
    }
  });

  await test('Line formulas and total recalculate in Excel',async()=>{
    for(const r of productRows(generated))assert.equal(cleanFormula(generated['M'+r]?.f),`D${r}*E${r}*H${r}`);
    const zip=fflate.unzipSync(baselineOutput.bytes);
    const workbook=fflate.strFromU8(zip['xl/workbook.xml']);
    assert.match(workbook,/<calcPr\b[^>]*(?:fullCalcOnLoad="(?:1|true)"|forceFullCalc="(?:1|true)")/);
  });

  await test('Blank-stock mode never invents the missing CxS inventory',async()=>{
    const blank=await output(baseline,'blank');
    assert.equal(value(blank.sheet.K1),'CxS');
    for(const r of productRows(blank.sheet))assert.equal(value(blank.sheet['K'+r]),'','K'+r);
    assert.ok(productRows(reference).some(r=>Number(value(reference['K'+r]))>0),'fixture contains actual historical stock');
  });

  await test('Workbook preserves reference style parts and removes broken external workbook links',async()=>{
    const model=ctx.CANVAS_MODEL;
    assert.ok(model?.templateZip,'embedded template available');
    const template=fflate.unzipSync(new Uint8Array(Buffer.from(model.templateZip,'base64')));
    const zip=fflate.unzipSync(baselineOutput.bytes);
    assert.equal(fflate.strFromU8(zip['xl/styles.xml']),fflate.strFromU8(template['xl/styles.xml']),'font/border/fill/number-format XML unchanged');
    const xml=fflate.strFromU8(zip['xl/worksheets/sheet1.xml']);
    const templ=fflate.strFromU8(template['xl/worksheets/sheet1.xml']);
    assert.equal(xml.match(/<cols>.*?<\/cols>/s)?.[0],templ.match(/<cols>.*?<\/cols>/s)?.[0],'reference column widths retained');
    assert.ok(!Object.keys(zip).some(k=>k.startsWith('xl/externalLinks/')),'no external link parts');
    assert.ok(!fflate.strFromU8(zip['xl/workbook.xml']).includes('<externalReferences>'),'no external refs');
    assert.ok(!fflate.strFromU8(zip['xl/workbook.xml']).includes('#REF!'),'no broken defined names');
  });

  await test('Updated input prices, discount and VAT replace old reference prices',async()=>{
    const input=mutated(s=>{put(s,'O2',0.99);put(s,'P2',0.1);put(s,'Q2',10);});
    const result=await output(await analyzeBytes(input));
    const r=rowForCode(result.sheet,29065);
    assert.ok(r,'changed product retained');
    assert.equal(value(result.sheet['H'+r]),0.89);
    assert.equal(value(result.sheet['I'+r]),0.98);
  });

  await test('A new article is retained, classified visibly and preserves leading-zero EAN text',async()=>{
    const input=mutated(s=>{
      const r=XLSX.utils.decode_range(s['!ref']).e.r+2;
      put(s,'A'+r,'CATEGORIA TEST NUOVA');put(s,'B'+r,'I');put(s,'C'+r,99999999);
      put(s,'E'+r,12);put(s,'F'+r,24);put(s,'G'+r,'ARTICOLO TEST NUOVO');
      put(s,'O'+r,2.5);put(s,'P'+r,'TP');put(s,'Q'+r,22);put(s,'R'+r,'0012345678901');
      s['!ref']='A1:R'+r;
    });
    const analysis=await analyzeBytes(input);
    const result=await output(analysis);
    assert.equal(productRows(result.sheet).length,6236);
    const r=rowForCode(result.sheet,99999999);assert.ok(r);
    assert.equal(value(result.sheet['B'+r]),'0012345678901');
    assert.equal(result.sheet['B'+r].t,'s');
    assert.equal(value(result.sheet['H'+r]),2.5);assert.equal(value(result.sheet['I'+r]),3.05);
    assert.ok(Object.values(result.sheet).some(c=>c&&c.v&&String(c.v).includes('CATEGORIA TEST NUOVA')),'new category heading retained');
    assert.ok(analysis.warnings?.length,'new products need visible review');
  });

  await test('An unknown SM gift without a source price never inherits an invented price',async()=>{
    const input=mutated(s=>{
      const r=XLSX.utils.decode_range(s['!ref']).e.r+2;
      put(s,'A'+r,'OMAGGIO TEST NUOVO');put(s,'B'+r,'I');put(s,'C'+r,99999998);
      put(s,'E'+r,1);put(s,'F'+r,1);put(s,'G'+r,'OMAGGIO TEST SENZA PREZZO');
      put(s,'P'+r,'SM');put(s,'Q'+r,22);s['!ref']='A1:R'+r;
    });
    const analysis=await analyzeBytes(input);
    const result=await output(analysis);
    const r=rowForCode(result.sheet,99999998);assert.ok(r,'unknown gift retained');
    assert.equal(value(result.sheet['H'+r]),'');assert.equal(value(result.sheet['I'+r]),'');
    assert.ok(analysis.audit.some(a=>String(a.code)==='99999998'&&a.type==='prezzo mancante'),'missing price explicitly reported');
  });

  await test('A changed gift cannot inherit an old externally supplied price',async()=>{
    const input=mutated(s=>{put(s,'J557','PIASTRA NUOVO MODELLO TEST');});
    const analysis=await analyzeBytes(input);
    const result=await output(analysis);
    const r=rowForCode(result.sheet,8534);assert.ok(r,'changed gift retained');
    assert.equal(value(result.sheet['H'+r]),'','old manual net 10.00 must not be reused');
    assert.equal(value(result.sheet['I'+r]),'','old manual gross must not be reused');
    assert.ok(analysis.audit.some(a=>String(a.code)==='8534'&&a.type==='prezzo mancante'));
  });

  await test('Unknown purchasing conditions are explicitly reported for review',async()=>{
    const text='ACQUISTANDO 7 CT CONDIZIONE TEST NUOVA';
    const input=mutated(s=>{
      const r=XLSX.utils.decode_range(s['!ref']).e.r+2;
      put(s,'A'+r,'CATEGORIA NOTA TEST');put(s,'G'+r,text);s['!ref']='A1:R'+r;
    });
    const analysis=await analyzeBytes(input);
    assert.ok(analysis.audit.some(a=>String(a.detail).includes(text)),'unmapped condition present in report');
    assert.ok(analysis.warnings.some(w=>String(w).toLowerCase().includes('note commerciali')),'unmapped condition visibly warned');
  });

  await test('Deleted products are not reintroduced from the old template',async()=>{
    const input=mutated(s=>{for(let c=0;c<18;c++)delete s[XLSX.utils.encode_cell({r:1,c})];});
    const result=await output(await analyzeBytes(input));
    assert.equal(rowForCode(result.sheet,29065),undefined);
    assert.equal(productRows(result.sheet).length,6234);
    const last=XLSX.utils.decode_range(result.sheet['!ref']).e.r+1;
    assert.equal(cleanFormula(result.sheet.M2?.f),`SUM(M3:M${last})`,'total adapts to changed row count');
  });

  await test('Deleted display products never leave stale composition lines behind',async()=>{
    const input=mutated(s=>{for(let c=0;c<18;c++)delete s[XLSX.utils.encode_cell({r:820,c})];});
    const result=await output(await analyzeBytes(input));
    assert.equal(rowForCode(result.sheet,23475),undefined);
    assert.ok(!Object.values(result.sheet).some(c=>c&&String(c.v).includes('48 DEO BOROTALCO SPRAY ORIGINALE')),'composition belongs to deleted display parent');
  });

  await test('Changed display compositions use the new input quantities, text and EAN',async()=>{
    const input=mutated(s=>{put(s,'I822',60);put(s,'J822','DEO BOROTALCO TEST AGGIORNATO');put(s,'R822','0012345678901');});
    const result=await output(await analyzeBytes(input));
    const range=XLSX.utils.decode_range(result.sheet['!ref']);
    let found;
    for(let r=1;r<=range.e.r+1;r++)if(str(result.sheet['F'+r]).includes('DEO BOROTALCO TEST AGGIORNATO'))found=r;
    assert.ok(found,'new composition description must be exported');
    assert.match(str(result.sheet['F'+found]),/60/,'new composition quantity');
    assert.equal(value(result.sheet['B'+found]),'0012345678901');
    assert.equal(result.sheet['B'+found].t,'s');
  });

  await test('Changed source description suffixes never retain their old wording',async()=>{
    const input=mutated(s=>{put(s,'J1002','palette+zucchero NUOVO');});
    const result=await output(await analyzeBytes(input));
    const r=rowForCode(result.sheet,28187);assert.ok(r);
    assert.ok(!str(result.sheet['F'+r]).includes('paletta+bicchiere+zucchero'),'old source suffix removed');
    assert.ok(Object.values(result.sheet).some(c=>c&&String(c.v).includes('palette+zucchero NUOVO')),'new source suffix retained');
  });

  await test('Descriptions are written as literal strings, never executable Excel formulas',async()=>{
    const text='=HYPERLINK("https://example.com","test")';
    const input=mutated(s=>{put(s,'G2',text);});
    const result=await output(await analyzeBytes(input));
    const r=rowForCode(result.sheet,29065);assert.ok(r);
    assert.equal(result.sheet['F'+r].f,undefined);
    assert.ok(str(result.sheet['F'+r]).startsWith(text));
  });

  await test('Malformed and already-processed workbooks fail clearly',async()=>{
    await assert.rejects(()=>analyzeBytes(Buffer.from('not an Excel file')));
    await assert.rejects(()=>analyzeBytes(fs.readFileSync(referencePath)));
  });

  const tempDir=fs.mkdtempSync(path.join(os.tmpdir(),'canvas-verified-'));
  fs.writeFileSync(path.join(tempDir,'New.xlsx'),baselineOutput.bytes);
  console.log(`\n${passed} functional tests passed. Verified workbook: ${path.join(tempDir,'New.xlsx')}`);
})().catch(error=>{console.error('\nFAIL after '+passed+' completed tests.');console.error(error);process.exitCode=1;});
