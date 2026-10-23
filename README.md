# New — Canvas offline

Convertitore locale del formato Canvas A–R nel layout `New Grieco N°44`, con download di `New.xlsx`. Il file di distribuzione `New.html` è autonomo: nessun server, installazione o collegamento Internet richiesto. Le CxS rimangono vuote, come richiesto.

## Scarica il programma

**[Scarica lo ZIP pronto all'uso](https://github.com/federicocrisci/excelProjectNino/raw/refs/heads/main/download/New-Canvas-offline.zip)**, estrailo e apri `New.html` con un doppio clic. Carica il Canvas originale, controlla le segnalazioni e scarica `New.xlsx`.

La cartella [download](download/) contiene anche l'HTML autonomo. Per scaricare tutto il progetto, usa **Code → Download ZIP** su GitHub e apri `index.html` nella cartella estratta.

Per l'uso quotidiano leggere [LEGGIMI.txt](LEGGIMI.txt). Le regole inferite, le informazioni manuali e le anomalie del campione sono descritte in [REPORT.md](REPORT.md).

## Sviluppo

Aprire `index.html` per usare la versione a file separati, oppure generare l'applicazione autonoma:

```sh
python tools/package_app.py
```

Il comando produce `dist/New.html` e `dist/New-Canvas-offline.zip`. Il pacchetto contiene l'HTML, la guida e le licenze. Le librerie JavaScript sono incluse in `vendor`; non serve `npm install`. `template.js` include il catalogo e gli stili ricavati dagli esempi del cliente, senza dover distribuire gli allegati originali.

Per ricostruire il modello servono Python, openpyxl, xlrd e la conversione XLS→XLSX del campione mediante LibreOffice. Il comando completo è nel report. Non è necessario ricostruire il modello per usare il programma o eseguire i test.

## Validazione

```sh
node tests/test_converter.cjs /percorso/originale.xlsx /percorso/riferimento.xls
```

Le fixture originali sono dati del cliente e non sono incluse nello ZIP. Nella macchina cloud i test trovano automaticamente i due allegati in `/workspace/attachments`.

Risultato sul campione: 17 test funzionali superati; 86.047 celle confrontate; 6.235 articoli e 22 dettagli riprodotti. Il confronto con CxS storiche serve esclusivamente al test; l'interfaccia esporta sempre CxS vuote. Differenze deliberate: totale generale esteso fino all'ultima riga e prezzo dell'articolo 26794 corretto dai dati origine. Verificati ricalcolo effettivo con LibreOffice, font, dimensioni, allineamenti, bordi, altezze e larghezze esatte rispetto al riferimento XLS.

La versione autonoma è stata verificata nel browser con rete disabilitata e CSP attiva: caricamento, analisi e download Excel/CSV funzionanti, nessuna richiesta di rete. La prova su un altro Canvas reale resta da eseguire quando sarà fornito.

## Struttura

- `converter.js`: riconoscimento input, associazione articoli, trasformazione, scrittura OOXML e rapporto CSV.
- `template.js`: modello locale, catalogo di associazione e componenti OOXML per gli stili.
- `app.js`, `index.html`, `styles.css`: interfaccia italiana accessibile e responsive.
- `tools/`: ricostruzione modello e pacchetto autonomo.
- `tests/`: confronti e casi di regressione sui dati forniti.

I testi dei fogli sono trattati come dati, non come codice o istruzioni. I valori testuali vengono scritti in celle stringa; non vengono eseguite macro o formule provenienti dai nomi/descrizioni caricati.
