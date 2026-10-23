/* New: interfaccia locale. Non effettua richieste di rete. */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const el = {};
  ['fileInput','dropzone','selectedFile','fileName','fileSize','replaceFile','progress','progressText','errorBox','errorText','reviewSection','reviewBadge','sourceCount','outputCount','categoryCount','warningsBox','warningsList','ackWarnings','categoriesBox','categoriesList','previewBody','previewCaption','downloadButton','downloadSuccess','downloadTitle','downloadDescription','liveStatus','stepUpload','stepReview','stepDownload','reportButton'].forEach(id => { el[id] = $(id); });
  let analysis = null;
  let processing = false;
  let warningsNeedAcknowledgement = false;
  let selectionToken = 0;
  const countFormat = new Intl.NumberFormat('it-IT');
  const moneyFormat = new Intl.NumberFormat('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  function setStatus(message) { el.liveStatus.textContent = message; }
  function setStep(step) {
    [el.stepUpload,el.stepReview,el.stepDownload].forEach((item,index) => {
      item.classList.toggle('active', index + 1 === step);
      item.classList.toggle('complete', index + 1 < step);
      if (index + 1 === step) item.setAttribute('aria-current','step');
      else item.removeAttribute('aria-current');
    });
  }
  function updateDownload() {
    el.downloadButton.disabled = processing || !analysis || (warningsNeedAcknowledgement && !el.ackWarnings.checked);
    el.reportButton.disabled = processing || !analysis;
  }
  function setBusy(busy,message) {
    processing = busy;
    el.progress.hidden = !busy;
    el.progressText.textContent = message || 'Lettura del Canvas…';
    el.dropzone.classList.toggle('busy',busy);
    el.dropzone.setAttribute('aria-disabled',String(busy));
    el.fileInput.disabled = busy;
    el.replaceFile.disabled = busy;
    el.downloadButton.querySelector('span').textContent = busy && analysis ? 'Generazione in corso…' : 'Genera e scarica New.xlsx';
    updateDownload();
    if (message) setStatus(message);
  }
  function showError(error) {
    el.errorText.textContent = error && error.message ? error.message : String(error || 'Si è verificato un errore inatteso. Riprova con il file originale.');
    el.errorBox.hidden = false;
    setStatus('Elaborazione non riuscita. ' + el.errorText.textContent);
  }
  function pickFile() { if (!processing) el.fileInput.click(); }
  function readableCount(value) { return typeof value === 'number' && Number.isFinite(value) ? countFormat.format(value) : '—'; }
  function cellValue(value) { return value === null || value === undefined ? '' : String(value); }
  function currency(value) { return typeof value === 'number' && Number.isFinite(value) ? moneyFormat.format(value) : cellValue(value); }
  function extractWarning(warning) { return typeof warning === 'string' ? warning : warning.message || warning.text || String(warning); }

  function renderAnalysis(result) {
    analysis = result;
    const stats = result.stats || {};
    const products = result.products || [];
    const categories = result.categories || [];
    const warnings = (result.warnings || []).map(extractWarning).filter(Boolean);
    const preview = result.preview || products.slice(0,8);
    el.sourceCount.textContent = readableCount(stats.sourceProducts ?? stats.inputProducts ?? products.length);
    el.outputCount.textContent = readableCount(stats.outputProducts ?? stats.products ?? products.length);
    el.categoryCount.textContent = readableCount(stats.categoryCount ?? categories.length);
    el.warningsList.replaceChildren();
    warnings.forEach(message => { const li = document.createElement('li'); li.textContent = message; el.warningsList.append(li); });
    warningsNeedAcknowledgement = warnings.length > 0;
    el.warningsBox.hidden = !warningsNeedAcknowledgement;
    el.ackWarnings.checked = false;
    el.reviewBadge.textContent = warningsNeedAcknowledgement ? 'Da controllare' : 'Analisi completata';
    el.reviewBadge.classList.toggle('needs-review',warningsNeedAcknowledgement);
    el.categoriesList.replaceChildren();
    categories.slice(0,8).forEach(category => {
      const chip = document.createElement('span');
      chip.className = 'category-chip';
      chip.textContent = typeof category === 'string' ? category : category.name || category.title || category.label || '';
      if (chip.textContent) el.categoriesList.append(chip);
    });
    if (categories.length > 8) { const more = document.createElement('span'); more.className = 'category-more'; more.textContent = '+ ' + countFormat.format(categories.length - 8) + ' altre'; el.categoriesList.append(more); }
    el.categoriesBox.hidden = !el.categoriesList.childElementCount;
    el.previewBody.replaceChildren();
    preview.slice(0,8).forEach(product => {
      const row = document.createElement('tr');
      const values = [cellValue(product.code ?? product.articleCode),cellValue(product.description ?? product.name),cellValue(product.units ?? product.pack ?? product.pieces),currency(product.net ?? product.netPrice),currency(product.gross ?? product.grossPrice)];
      values.forEach((value,index) => { const cell = document.createElement('td'); cell.textContent = value; if (index >= 2) cell.className = 'numeric'; row.append(cell); });
      el.previewBody.append(row);
    });
    if (!preview.length) { const row = document.createElement('tr'); const cell = document.createElement('td'); cell.colSpan=5; cell.textContent='L’anteprima non è disponibile per questo file.'; row.append(cell); el.previewBody.append(row); }
    el.previewCaption.textContent = 'PRIMI ' + Math.min(preview.length,8) + ' ARTICOLI';
    el.downloadSuccess.hidden = true;
    el.downloadTitle.textContent = warningsNeedAcknowledgement ? 'Controlla le segnalazioni, poi scarica.' : 'Pronto per il prossimo passaggio.';
    el.downloadDescription.textContent = warningsNeedAcknowledgement ? 'Conferma di aver letto i punti qui sopra per generare New.xlsx.' : 'Genera il file e aprilo con Excel per lavorare sul tuo Canvas.';
    el.reviewSection.hidden = false;
    setStep(2);
    updateDownload();
    setStatus('Analisi completata. ' + warnings.length + ' segnalazioni da controllare.');
  }

  async function readFile(file) {
    if (!file || processing) return;
    const token = ++selectionToken;
    analysis = null;
    updateDownload();
    el.reviewSection.hidden = true;
    el.selectedFile.hidden = true;
    el.dropzone.hidden = false;
    el.errorBox.hidden = true;
    el.downloadSuccess.hidden = true;
    setStep(1);
    if (!/\.(xlsx|xls)$/i.test(file.name)) { showError(new Error('Scegli un file Excel originale con estensione .xlsx o .xls.')); return; }
    if (file.size > 30 * 1024 * 1024) { showError(new Error('Il file supera il limite di 30 MB. Scegli un Canvas più piccolo.')); return; }
    el.fileName.textContent = file.name;
    el.fileSize.textContent = file.size >= 1024*1024 ? (file.size / (1024*1024)).toLocaleString('it-IT',{maximumFractionDigits:1}) + ' MB' : Math.max(1,Math.round(file.size/1024)).toLocaleString('it-IT') + ' KB';
    el.selectedFile.hidden = false;
    el.dropzone.hidden = true;
    setBusy(true,'Lettura del file e confronto con il modello…');
    try {
      if (!window.CanvasConverter || typeof window.CanvasConverter.analyze !== 'function') throw new Error('Mancano alcuni file del programma. Estrai prima tutto lo ZIP in una cartella e apri index.html da lì.');
      await new Promise(resolve => setTimeout(resolve,30));
      const bytes = await file.arrayBuffer();
      const result = await window.CanvasConverter.analyze(bytes,file.name);
      if (token !== selectionToken) return;
      renderAnalysis(result);
    } catch (error) { if (token === selectionToken) showError(error); }
    finally { if (token === selectionToken) setBusy(false); }
  }

  function saveBlob(blob,name) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url; link.download = name;
    document.body.append(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url),60000);
  }

  function downloadReport() {
    if (!analysis || processing) return;
    try {
      const csv = window.CanvasConverter.reportCSV(analysis);
      saveBlob(new Blob([csv],{type:'text/csv;charset=utf-8'}),'Controlli.csv');
      setStatus('Dettaglio delle trasformazioni generato. Download di Controlli.csv avviato.');
    } catch (error) { showError(error); }
  }

  async function download() {
    if (!analysis || processing || (warningsNeedAcknowledgement && !el.ackWarnings.checked)) return;
    el.errorBox.hidden = true;
    el.downloadSuccess.hidden = true;
    setBusy(true,'Creazione di New.xlsx con formati e formule…');
    try {
      await new Promise(resolve => setTimeout(resolve,30));
      const generated = await window.CanvasConverter.generate(analysis,{stockMode:'blank'});
      const bytes = generated && generated.bytes ? generated.bytes : generated;
      saveBlob(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),'New.xlsx');
      el.downloadSuccess.hidden = false;
      el.downloadTitle.textContent = 'Il tuo New.xlsx è pronto.';
      el.downloadDescription.textContent = 'Puoi scaricarlo di nuovo oppure caricare un altro Canvas.';
      setStep(3);
      setStatus('New.xlsx generato. Download avviato.');
    } catch (error) { showError(error); }
    finally { setBusy(false); }
  }

  el.dropzone.addEventListener('click',pickFile);
  el.dropzone.addEventListener('keydown',event => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); pickFile(); } });
  el.replaceFile.addEventListener('click',pickFile);
  el.fileInput.addEventListener('change',event => { const file = event.target.files[0]; if (file) readFile(file); el.fileInput.value=''; });
  ['dragenter','dragover'].forEach(type => el.dropzone.addEventListener(type,event => { event.preventDefault(); if (!processing) el.dropzone.classList.add('dragging'); }));
  ['dragleave','drop'].forEach(type => el.dropzone.addEventListener(type,event => { event.preventDefault(); el.dropzone.classList.remove('dragging'); }));
  el.dropzone.addEventListener('drop',event => { if (event.dataTransfer.files.length > 1) { showError(new Error('Carica un solo Canvas alla volta.')); return; } readFile(event.dataTransfer.files[0]); });
  document.addEventListener('dragover',event => { if (event.dataTransfer && Array.from(event.dataTransfer.types).includes('Files')) event.preventDefault(); });
  document.addEventListener('drop',event => { if (event.dataTransfer && event.dataTransfer.files.length) event.preventDefault(); });
  el.downloadButton.addEventListener('click',download);
  el.reportButton.addEventListener('click',downloadReport);
  el.ackWarnings.addEventListener('change',updateDownload);
  window.CanvasApp = Object.freeze({ setAnalysis:renderAnalysis, readFile });
  updateDownload();
})();
