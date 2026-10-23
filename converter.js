/* New Canvas — trasformazione locale. Nessuna richiesta di rete. */
(function (root) {
  'use strict';
  const LIMIT = 30 * 1024 * 1024;
  const MAX_ROWS = 100000;
  const COLS = 'ABCDEFGHIJKLM'.split('');
  const empty = value => value === null || value === undefined || value === '';
  const text = value => empty(value) ? '' : String(value).trim();
  const code = value => text(value);
  const normal = value => empty(value) ? null : typeof value === 'string' ? value.trim() : value;
  const signature = values => JSON.stringify(values.map(normal));
  const same = (a,b) => signature(a) === signature(b);
  const xml = value => String(value).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
  function number(value) {
    if (empty(value)) return null;
    if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
    let v = String(value).trim().replace(/\s/g,'');
    if (v.includes(',')) v = v.replace(/\./g,'').replace(',','.');
    return /^[-+]?\d+(?:\.\d+)?$/.test(v) ? Number(v) : NaN;
  }
  function rounded(value) { return Math.round((value + 1e-10) * 100) / 100; }
  function discount(value) {
    const v = text(value).toUpperCase();
    if (v === 'TP' || v === 'SM' || !v) return 0;
    return v.endsWith('%') ? number(v.slice(0,-1))/100 : number(value);
  }
  function readCell(sheet,row,column) {
    const c = sheet[root.XLSX.utils.encode_cell({r:row,c:column})];
    if (!c || c.t === 'z' || empty(c.v)) return null;
    if (c.t === 'e') throw new Error('Il foglio contiene un errore Excel alla riga ' + (row+1) + '. Correggilo prima della conversione.');
    return c.v;
  }
  function sourceRow(sheet,row) { return Array.from({length:18},(_,c) => readCell(sheet,row,c)); }
  function isItem(s) {
    return !empty(s[2]) && !empty(s[6]) && (text(s[1]).toUpperCase() === 'I' || (!empty(s[4]) && !empty(s[5])));
  }
  function validate(s,row) {
    const fields = [[4,'pezzi per cartone'],[5,'CxB'],[16,'IVA']];
    for (const [idx,name] of fields) {
      const n = number(s[idx]);
      if (n === null || !Number.isFinite(n) || n < 0 || (idx===16 && n>100)) throw new Error('Riga '+row+': valore '+name+' mancante o non valido.');
    }
    const d = discount(s[15]);
    if (!Number.isFinite(d) || d<0 || d>1) throw new Error('Riga '+row+': sconto non riconosciuto. Sono ammessi TP, SM o una percentuale.');
    const p = number(s[14]);
    if ((p === null && text(s[15]).toUpperCase() !== 'SM') || (p !== null && (!Number.isFinite(p) || p<0))) throw new Error('Riga '+row+': prezzo origine mancante o non valido.');
    if (typeof s[17] === 'number' && (!Number.isSafeInteger(s[17]) || s[17]<0)) throw new Error('Riga '+row+': il codice EAN numerico non è integro. Salvalo come testo nel file originale.');
  }
  function chooseReference(s,options,used) {
    let best=null,score=-Infinity;
    for (const r of options || []) {
      if (used.has(r.r)) continue;
      let n = 0;
      if (same(s,r.source)) n+=1000;
      if (normal(s[0])===normal(r.source[0])) n+=40;
      if (normal(s[6])===normal(r.source[6])) n+=30;
      if (normal(s[17])===normal(r.source[17])) n+=20;
      if (normal(s[15])===normal(r.source[15])) n+=8;
      if (normal(s[4])===normal(r.source[4])) n+=2;
      if (normal(s[5])===normal(r.source[5])) n+=2;
      if (n>score) { score=n;best=r; }
    }
    // Un codice può apparire sia in vendita sia in omaggio: non confondere le due righe.
    if (best && (text(s[15]).toUpperCase()==='SM') !== (text(best.source[15]).toUpperCase()==='SM')) return null;
    return best;
  }
  function makeProduct(s,row,ref,netMap,grossMap,audit) {
    const disc = discount(s[15]);
    const rawPrice = number(s[14]);
    const vat = number(s[16]);
    const originalPrice = ref && same([s[14],s[15],s[16]],[ref.source[14],ref.source[15],ref.source[16]]);
    const originalDescription = ref && same(s.slice(6,11),ref.source.slice(6,11));
    const metadataSame = ref && same(s,ref.source);
    let net=null,gross=null,pricing='standard';
    if (originalPrice && !empty(ref.values[7]) && (rawPrice!==null || metadataSame)) {
      net=number(ref.values[7]);gross=number(ref.values[8]);pricing='campione';
    } else if (rawPrice !== null) {
      const key=JSON.stringify([rawPrice,disc]);
      net=netMap.has(key) ? netMap.get(key) : rounded(rawPrice*(1-disc));
      const gkey=JSON.stringify([net,vat]);
      gross=grossMap.has(gkey) ? grossMap.get(gkey) : rounded(net*(1+vat/100));
      pricing=netMap.has(key) && grossMap.has(gkey) ? 'arrotondamento campione' : 'standard';
    }
    // Nel campione questo articolo contiene un evidente errore di battitura (224,45).
    if (code(s[2])==='26794' && rawPrice!==null && net!==null && Math.abs(net-rawPrice*(1-disc))>0.02) {
      net=rounded(rawPrice*(1-disc));gross=rounded(net*(1+vat/100));pricing='prezzo origine';
      audit.push({type:'prezzo corretto',sourceRow:row,code:code(s[2]),description:text(s[6]),detail:'Il campione riportava 224,45; usato il prezzo origine '+net+'.'});
    }
    const suffix=extraDescription(s)+(empty(s[10])?'':(extraDescription(s)?' ':'')+text(s[10]));
    let description=originalDescription ? ref.values[5] : text(s[6])+(suffix?'    '+suffix:'');
    if (!ref && text(s[15]).toUpperCase()==='SM' && !/OMAGGIO/i.test(description)) description='IN OMAGGIO '+description;
    const sourceFlag=/novit/i.test(text(s[12]))?'NOVITA':/promo/i.test(text(s[12]))?'PROMO':'';
    const flag=metadataSame ? ref.values[6] : sourceFlag;
    const variance=metadataSame ? ref.values[9] : '';
    const values=['',empty(s[17])?'':String(s[17]),code(s[2]),'',number(s[4]),description,flag||'',net,gross,variance||'','',number(s[5]),0];
    const p={code:code(s[2]),ean:values[1],description,units:values[4],net,gross,vat,discount:disc,sourceRow:row,source:s,reference:ref,category:ref?ref.category:text(s[0]),values,pricing};
    if (!ref) audit.push({type:'articolo nuovo',sourceRow:row,code:p.code,description,detail:'Non presente nel campione; dati origine e stile standard. Categoria: '+p.category});
    if (ref && !metadataSame && (!empty(ref.values[6]) || !empty(ref.values[9]))) audit.push({type:'annotazioni da rivedere',sourceRow:row,code:p.code,description,detail:'Dati cambiati: etichette/Varianza manuali del campione non riprese.'});
    if (net===null) audit.push({type:'prezzo mancante',sourceRow:row,code:p.code,description,detail:'Omaggio senza prezzo: H e I lasciate vuote; il suo totale è zero finché non viene inserito un prezzo.'});
    if (text(s[15]).toUpperCase()==='SM' && originalPrice && net!==null) audit.push({type:'prezzo omaggio dal campione',sourceRow:row,code:p.code,description,detail:'Il prezzo non è nel file origine; è il valore storico del modello da verificare.'});
    if (pricing==='standard' && net!==null) audit.push({type:'arrotondamento standard',sourceRow:row,code:p.code,description,detail:'Nuova combinazione di prezzo/sconto/IVA: arrotondamento commerciale a 2 decimali.'});
    return p;
  }
  function extraDescription(s) {
    return [s[7],s[8],s[9]].filter(v=>!empty(v)).map(text).join(' ');
  }
  async function analyze(buffer,filename) {
    if (!root.XLSX || !root.fflate || !root.CANVAS_MODEL) throw new Error('File del programma mancanti: estrai tutto lo ZIP nella stessa cartella.');
    const bytes=buffer instanceof Uint8Array?buffer:new Uint8Array(buffer);
    if (!bytes.length || bytes.length>LIMIT) throw new Error('Scegli un Excel non vuoto, fino a 30 MB.');
    const zip=bytes[0]===0x50 && bytes[1]===0x4b;
    const ole=bytes[0]===0xd0 && bytes[1]===0xcf && bytes[2]===0x11 && bytes[3]===0xe0;
    if (!zip && !ole) throw new Error('Il file non è un Excel .xlsx o .xls valido.');
    let book;
    try { book=root.XLSX.read(bytes,{type:'array',cellFormula:true,cellStyles:false,cellDates:false}); }
    catch (_) { throw new Error('Impossibile leggere l’Excel. Controlla che non sia danneggiato o protetto da password.'); }
    const sheets=[];
    for (const name of book.SheetNames) {
      const sheet=book.Sheets[name];
      if (!sheet['!ref']) continue;
      const range=root.XLSX.utils.decode_range(sheet['!ref']);
      if (range.e.r>=MAX_ROWS) throw new Error('Il foglio '+name+' supera le 100.000 righe. Riduci l’area utilizzata prima di riprovare.');
      if (range.e.c<17) continue;
      const items=[],notes=[],compositions=[];
      let footer=0,parentItem=null;
      for (let r=range.s.r;r<=range.e.r;r++) {
        const s=sourceRow(sheet,r);
        if (isItem(s)) { validate(s,r+1);parentItem={source:s,sourceRow:r+1};items.push(parentItem); }
        else if (!empty(s[2]) && !empty(s[6])) footer++;
        else if (!empty(s[6]) && empty(s[2])) notes.push({source:s,sourceRow:r+1});
        else if (empty(s[2]) && (!empty(s[8]) || !empty(s[9])) && parentItem && text(s[0])===text(parentItem.source[0])) compositions.push({source:s,sourceRow:r+1,parentSourceRow:parentItem.sourceRow});
      }
      if (items.length) sheets.push({name,sheet,items,notes,compositions,footer});
    }
    if (!sheets.length) throw new Error('Non riconosco il Canvas originale. Servono le colonne A–R del formato ricevuto: codice in C, pezzi in E, CxB in F, descrizione in G, prezzo in O, sconto in P, IVA in Q ed EAN in R. Carica l’originale, non un file New già lavorato.');
    if (sheets.length>1) throw new Error('Ci sono più fogli con articoli compatibili. Salva il Canvas da lavorare in un file con un solo foglio dati.');
    const input=sheets[0],m=root.CANVAS_MODEL,audit=[],warnings=[];
    const refs=m.rows.filter(r=>r.kind==='product');
    const byCode=new Map(),categoryTargets=new Map(),categoryStyles=new Map();
    for (const r of refs) {
      const key=code(r.source[2]);
      if (!byCode.has(key)) byCode.set(key,[]);
      byCode.get(key).push(r);
      const a=text(r.source[0]);
      if (!categoryTargets.has(a)) categoryTargets.set(a,new Set());
      categoryTargets.get(a).add(r.category);
    }
    for (const r of m.rows) if(r.kind==='category') categoryStyles.set(r.category || text(r.values[0]),r);
    const netMap=new Map((m.netLookup||[]).map(([p,d,n])=>[JSON.stringify([p,d]),n]));
    const grossMap=new Map((m.grossLookup||[]).map(([p,v,n])=>[JSON.stringify([p,v]),n]));
    const used=new Set(),byReference=new Map(),unknown=[],excluded=[];
    for (const item of input.items) {
      const s=item.source;
      const omit=(m.excluded||[]).find(e=>same(e.source,s));
      if (omit) {
        excluded.push(item);
        audit.push({type:'riga accorpata/esclusa',sourceRow:item.sourceRow,code:code(s[2]),description:text(s[6]),detail:omit.reason||'Esclusione identica al modello campione.'});
        continue;
      }
      const ref=chooseReference(s,byCode.get(code(s[2])),used);
      if (ref) used.add(ref.r);
      const p=makeProduct(s,item.sourceRow,ref,netMap,grossMap,audit);
      if (ref) byReference.set(ref.r,p);
      else {
        const target=categoryTargets.get(text(s[0]));
        if (target && target.size===1) p.category=Array.from(target)[0];
        else p.category=text(s[0])||'ARTICOLI DA CLASSIFICARE';
        unknown.push(p);
      }
    }
    if (!byReference.size && !unknown.length) throw new Error('Il file non contiene articoli da esportare.');
    const generic=m.rows.find(r=>r.r===4)||refs[0];
    const categoryGeneric=m.rows.find(r=>r.kind==='category'&&r.r!==2)||m.rows.find(r=>r.kind==='category');
    const blocks=[];
    let block=null;
    for (const r of m.rows) {
      if (r.kind==='header') continue;
      if (r.kind==='category') {
        block={category:r.category||text(r.values[0]),header:r,rows:[],products:[]};blocks.push(block);
      } else if (block) block.rows.push(r);
    }
    const extraByCategory=new Map();
    for (const p of unknown) {
      if (!extraByCategory.has(p.category)) extraByCategory.set(p.category,[]);
      extraByCategory.get(p.category).push(p);
    }
    const rows=[{...m.rows.find(r=>r.kind==='header'),values:m.rows.find(r=>r.kind==='header').values.slice()}];
    const products=[],categories=[];
    function pushProduct(p,style) {
      rows.push({kind:'product',styles:style.styles,attrs:style.attrs,values:p.values,product:p});products.push(p);
    }
    const usedNotes=new Set();
    for (const b of blocks) {
      const selected=b.rows.filter(r=>byReference.has(r.r));
      const extra=extraByCategory.get(b.category)||[];
      if (!selected.length && !extra.length) continue;
      categories.push(b.category);
      rows.push({...b.header,values:b.header.values.slice()});
      let detailParent=null,detailIndex=0;
      for (const r of b.rows) {
        if (r.kind==='product') {
          detailParent=r;detailIndex=0;
          const p=byReference.get(r.r);if(p)pushProduct(p,r);
        } else if (r.kind==='detail') {
          // Riprende le note campione solo se nel gruppo rimangono i dati origine corrispondenti.
          const body=text(r.values[5]);
          const isCondition=/^ACQUISTANDO/i.test(body);
          const parent=detailParent&&byReference.get(detailParent.r);
          const referenceDetails=(m.sourceDetails||[]).filter(d=>d.kind==='composition'&&detailParent&&d.parentSourceRow===detailParent.sourceRow);
          const currentDetails=parent?input.compositions.filter(d=>d.parentSourceRow===parent.sourceRow):[];
          const baseline=referenceDetails[detailIndex];
          const current=currentDetails[detailIndex];
          const valid=isCondition ? input.notes.some(n=>text(n.source[6])===body && selected.some(x=>text(x.source[0])===text(n.source[0]))) : Boolean(parent && baseline && current && same(baseline.source,current.source));
          if (valid) {
            rows.push({...r,values:r.values.slice()});
            if(isCondition)for(const n of input.notes)if(text(n.source[6])===body&&selected.some(x=>text(x.source[0])===text(n.source[0])))usedNotes.add(n.sourceRow);
          }
          if(!isCondition)detailIndex++;
        } else if (r.kind==='blank') rows.push({...r,values:r.values.slice()});
      }
      for(const p of extra)pushProduct(p,generic);
      extraByCategory.delete(b.category);
    }
    for (const [category,ps] of extraByCategory) {
      categories.push(category);
      const header=categoryStyles.get(category)||categoryGeneric;
      rows.push({kind:'category',styles:header.styles,attrs:header.attrs,values:[' '+category,'','','','','','','','','','','','']});
      for(const p of ps)pushProduct(p,generic);
    }
    // Composizioni nuove o modificate: usa le righe caricate, mai quelle storiche.
    for(const p of products) {
      const current=input.compositions.filter(d=>d.parentSourceRow===p.sourceRow);
      const baseline=p.reference?(m.sourceDetails||[]).filter(d=>d.kind==='composition'&&d.parentSourceRow===p.reference.sourceRow):[];
      if(!current.length || (baseline.length===current.length && current.every((d,i)=>same(d.source,baseline[i].source))))continue;
      let insertion=rows.findIndex(r=>r.product===p)+1;
      // Le righe base uguali sono già state inserite; sostituire l'intera composizione evita doppioni.
      const rowTemplates=m.rows.filter(r=>r.kind==='detail' && r.parentR===p.reference?.r);
      while(rows[insertion]&&rows[insertion].kind==='detail'&&!/^ACQUISTANDO/i.test(text(rows[insertion].values[5])))rows.splice(insertion,1);
      for(let i=0;i<current.length;i++) {
        const d=current[i],style=rowTemplates[i]||m.rows.find(r=>r.kind==='detail')||generic;
        const values=['',empty(d.source[17])?'':String(d.source[17]),'','','','  '+extraDescription(d.source),'','','','','','',''];
        rows.splice(insertion++,0,{kind:'detail',styles:style.styles,attrs:style.attrs,values});
        audit.push({type:'composizione aggiornata',sourceRow:d.sourceRow,code:p.code,description:p.description,detail:values[5]});
      }
    }
    for (const n of input.notes) {
      if (usedNotes.has(n.sourceRow)) continue;
      // La nota ELVIVE non esiste nel risultato di riferimento; le altre note nuove sono esposte nel controllo.
      audit.push({type:'nota origine non esportata',sourceRow:n.sourceRow,code:'',description:text(n.source[0]),detail:text(n.source[6])});
    }
    if (rows.length<3 || !products.length) throw new Error('Nessun articolo esportabile.');
    const missed=refs.filter(r=>!used.has(r.r)).length;
    if(unknown.length)warnings.push(unknown.length+' articoli non presenti nel campione: inclusi senza eliminarli, con stile standard. Controlla le categorie nel dettaglio delle trasformazioni.');
    const changedMetadata=audit.filter(a=>a.type==='annotazioni da rivedere').length;
    if(changedMetadata)warnings.push(changedMetadata+' articoli hanno dati cambiati: PROMO/NOVITA e Varianza manuali del campione non sono state riprese.');
    const missing=audit.filter(a=>a.type==='prezzo mancante');
    if(missing.length)warnings.push(missing.length+' omaggi non hanno un prezzo: H e I restano vuote (codici '+missing.map(p=>p.code).slice(0,8).join(', ')+').');
    const gifts=audit.filter(a=>a.type==='prezzo omaggio dal campione').length;
    if(gifts)warnings.push(gifts+' prezzi degli omaggi provengono dal campione e richiedono verifica: non sono presenti nell’originale.');
    if(audit.some(a=>a.type==='prezzo corretto'))warnings.push('Codice 26794: usato il prezzo origine 22,45 al posto di 224,45 presente nel campione.');
    const fresh=audit.filter(a=>a.type==='arrotondamento standard').length;
    if(fresh)warnings.push(fresh+' combinazioni nuove di prezzo/sconto/IVA: usato arrotondamento commerciale a 2 decimali. Gli arrotondamenti già osservati nel campione vengono mantenuti.');
    if(excluded.length)warnings.push(excluded.length+' righe di omaggio accorpate/escluse come nel campione. Il dettaglio indica quali.');
    const notes=audit.filter(a=>a.type==='nota origine non esportata').length;
    if(notes)warnings.push('Note commerciali non presenti nel modello: '+notes+'. Sono riportate nel dettaglio delle trasformazioni.');
    const compositions=audit.filter(a=>a.type==='composizione aggiornata').length;
    if(compositions)warnings.push('Righe di composizione nuove o modificate: '+compositions+'. Usati i dati caricati; le annotazioni manuali del campione non sono state copiate.');
    if(products.some(p=>!empty(p.values[6])||!empty(p.values[9])))warnings.push('PROMO/NOVITA e Varianza sono annotazioni manuali del campione: conservate solo per righe con dati invariati, da verificare per il nuovo periodo.');
    audit.unshift({type:'riepilogo',sourceRow:'',code:'',description:filename||'Canvas',detail:input.items.length+' articoli origine; '+products.length+' esportati; '+excluded.length+' righe accorpate/escluse; '+missed+' righe campione assenti nel nuovo file. CxS vuote. Totale generale esteso a tutti gli articoli.'});
    return {filename:filename||'Canvas',sheetName:input.name,products,rows,categories,warnings,audit,stats:{sourceProducts:input.items.length,outputProducts:products.length,categoryCount:categories.length,excludedProducts:excluded.length,unknownProducts:unknown.length,absentReferenceProducts:missed,footerRows:input.footer},preview:products.slice(0,8)};
  }
  function rowXML(row,index,totalStyle,lastRow) {
    const attrs={...(row.attrs||{}),r:String(index)};
    delete attrs.spans;
    let out='<row'+Object.entries(attrs).map(([k,v])=>' '+k+'="'+xml(v)+'"').join('')+'>';
    for(let col=0;col<13;col++) {
      const address=COLS[col]+index;
      const style=index===2&&col===12?totalStyle:row.styles[col];
      const start='<c r="'+address+'" s="'+(style||0)+'"';
      if(col===12&&(row.kind==='product'||index===2)) {
        out+=start+' t="n"><f>'+(index===2?'SUM(M3:M'+lastRow+')':'D'+index+'*E'+index+'*H'+index)+'</f><v>0</v></c>';continue;
      }
      const v=row.values[col];
      if(empty(v)) {out+=start+'/>';continue;}
      if(typeof v==='number' && Number.isFinite(v))out+=start+' t="n"><v>'+v+'</v></c>';
      else out+=start+' t="inlineStr"><is><t xml:space="preserve">'+xml(v)+'</t></is></c>';
    }
    return out+'</row>';
  }
  async function generate(analysis,options) {
    if(!analysis || !Array.isArray(analysis.rows) || !analysis.products.length)throw new Error('Carica e analizza prima il Canvas originale.');
    const model=root.CANVAS_MODEL;
    const decoded=Uint8Array.from(root.atob(model.templateZip),c=>c.charCodeAt(0));
    const files=root.fflate.unzipSync(decoded);
    const decoder=new TextDecoder(),encoder=new TextEncoder();
    let sheet=decoder.decode(files['xl/worksheets/sheet1.xml']);
    const totalStyle=model.rows.find(r=>r.r===2).styles[12];
    const stock=options&&options.stockMode==='reference';
    const body=analysis.rows.map((r,i)=> {
      if(r.kind==='product') {
        const values=r.values.slice();
        values[10]=stock&&r.product.reference?r.product.reference.values[10]:'';
        return rowXML({...r,values},i+1,totalStyle,analysis.rows.length);
      }
      return rowXML(r,i+1,totalStyle,analysis.rows.length);
    }).join('');
    sheet=sheet.replace(/<sheetData(?:\s[^>]*)?\s*\/>|<sheetData(?:\s[^>]*)?>[\s\S]*?<\/sheetData>/,'<sheetData>'+body+'</sheetData>').replace(/<dimension\b[^>]*\/>/,'<dimension ref="A1:M'+analysis.rows.length+'"/>');
    files['xl/worksheets/sheet1.xml']=encoder.encode(sheet);
    let wb=decoder.decode(files['xl/workbook.xml']);
    wb=wb.replace(/<calcPr\b[^>]*\/>/,'<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>');
    wb=wb.replace(/<definedNames>[\s\S]*?<\/definedNames>/,'<definedNames><definedName localSheetId="0" name="_xlnm.Print_Area">CANVAS!$A$3:$M$'+analysis.rows.length+'</definedName><definedName localSheetId="0" name="_xlnm.Print_Titles">CANVAS!$1:$1</definedName></definedNames>');
    files['xl/workbook.xml']=encoder.encode(wb);
    return root.fflate.zipSync(files,{level:6});
  }
  function reportCSV(analysis) {
    const quote=v=>'"'+String(v??'').replace(/"/g,'""')+'"';
    const lines=[['Tipo','Riga origine','Codice','Descrizione','Dettaglio'],...analysis.audit.map(a=>[a.type,a.sourceRow,a.code,a.description,a.detail])];
    // Apostrofo iniziale nei campi testuali per evitare formule quando il CSV viene aperto in Excel.
    return '\ufeff'+lines.map(row=>row.map(v=>quote(typeof v==='string'&&/^[=+\-@\t\r]/.test(v)?"'"+v:v)).join(';')).join('\r\n');
  }
  root.CanvasConverter=Object.freeze({analyze,generate,reportCSV});
})(typeof window!=='undefined'?window:globalThis);
