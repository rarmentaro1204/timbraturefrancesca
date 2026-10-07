/**
 * Timbrature di Francesca - Google Apps Script
 * Ogni timbratura viene salvata come riga nel foglio "Timbrature"
 * del Google Sheet a cui è collegato lo script.
 * L'orario è preso dal server (non dal telefono), quindi non è modificabile.
 */
var TZ = 'Europe/Rome';
var SHEET = 'Timbrature';
var SHEET_ID = '1I2dV7k-UZ4lOqqUVh7oJ0RncrT0yMMgDCoJ41aI6W1Q'; // Google Sheet "Timbrature Francesca"
var ENTRATA = 'ENTRATA';
var USCITA = 'USCITA';

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try { return json_(stato()); } catch (err) { return json_({ errore: String(err) }); }
}

function doPost(e) {
  try { return json_(timbra()); } catch (err) { return json_({ errore: String(err) }); }
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['Data e ora', 'Data', 'Ora', 'Tipo']);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('C:C').setNumberFormat('@');
  }
  return sh;
}

function leggi_() {
  var sh = getSheet_();
  var n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 4).getValues().map(function (r) {
    var ts = new Date(r[0]);
    return {
      ts: ts,
      giorno: Utilities.formatDate(ts, TZ, 'yyyy-MM-dd'),
      ora: Utilities.formatDate(ts, TZ, 'HH:mm:ss'),
      tipo: r[3]
    };
  });
}

/** Ore lavorate (in minuti) per giorno: accoppia ENTRATA->USCITA dello stesso giorno. */
function minutiPerGiorno_(righe, adesso) {
  var perGiorno = {};
  var aperta = null;
  righe.forEach(function (r) {
    if (r.tipo === ENTRATA) {
      aperta = r;
    } else if (r.tipo === USCITA && aperta && aperta.giorno === r.giorno) {
      perGiorno[r.giorno] = (perGiorno[r.giorno] || 0) + (r.ts - aperta.ts) / 60000;
      aperta = null;
    } else {
      aperta = null;
    }
  });
  var oggi = Utilities.formatDate(adesso, TZ, 'yyyy-MM-dd');
  if (aperta && aperta.giorno === oggi) {
    perGiorno[oggi] = (perGiorno[oggi] || 0) + (adesso - aperta.ts) / 60000;
  }
  return perGiorno;
}

function stato() {
  var adesso = new Date();
  var oggi = Utilities.formatDate(adesso, TZ, 'yyyy-MM-dd');
  var mese = oggi.substring(0, 7);
  var righe = leggi_();
  var ultima = righe.length ? righe[righe.length - 1] : null;

  var prossimo = ENTRATA;
  var dimenticata = null;
  if (ultima && ultima.tipo === ENTRATA) {
    if (ultima.giorno === oggi) prossimo = USCITA;
    else dimenticata = ultima.giorno; // uscita non timbrata in un giorno precedente
  }

  var perGiorno = minutiPerGiorno_(righe, adesso);
  var minMese = 0;
  Object.keys(perGiorno).forEach(function (g) {
    if (g.indexOf(mese) === 0) minMese += perGiorno[g];
  });

  return {
    adesso: adesso.getTime(),
    prossimo: prossimo,
    dimenticata: dimenticata,
    oggi: righe.filter(function (r) { return r.giorno === oggi; })
      .map(function (r) { return { ora: r.ora, tipo: r.tipo }; }),
    minutiOggi: Math.round(perGiorno[oggi] || 0),
    minutiMese: Math.round(minMese)
  };
}

function timbra() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var s = stato();
    var adesso = new Date();
    getSheet_().appendRow([
      adesso,
      Utilities.formatDate(adesso, TZ, 'yyyy-MM-dd'),
      Utilities.formatDate(adesso, TZ, 'HH:mm:ss'),
      s.prossimo
    ]);
  } finally {
    lock.releaseLock();
  }
  try { aggiornaReport_(); } catch (err) { /* il report non deve bloccare la timbratura */ }
  return stato();
}

// ---------------------------------------------------------------------------
// REPORT: schede "Report mensile" e "Riepilogo mesi" (si aggiornano da sole)
// ---------------------------------------------------------------------------
var MESI = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio',
            'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
var GIORNI = ['domenica', 'lunedì', 'martedì', 'mercoledì', 'giovedì', 'venerdì', 'sabato'];

/** Da lanciare una volta dall'editor (o quando vuoi ricostruire i report). */
function creaReport() {
  aggiornaReport_();
}

/** Raggruppa le righe per giorno: fasce ENTRATA->USCITA e minuti lavorati. */
function calcolaGiorni_(righe) {
  var giorni = {};
  var aperta = null;
  righe.forEach(function (r) {
    var g = giorni[r.giorno] || (giorni[r.giorno] = { fasce: [], minuti: 0, incompleto: false });
    if (r.tipo === ENTRATA) {
      if (aperta) aperta.giorno.incompleto = true;
      var f = { da: r.ora.substring(0, 5), a: null };
      g.fasce.push(f);
      aperta = { fascia: f, ts: r.ts, giorno: g, key: r.giorno };
    } else if (r.tipo === USCITA && aperta && aperta.key === r.giorno) {
      aperta.fascia.a = r.ora.substring(0, 5);
      g.minuti += (r.ts - aperta.ts) / 60000;
      aperta = null;
    } else {
      g.fasce.push({ da: null, a: r.ora.substring(0, 5) });
      g.incompleto = true;
      aperta = null;
    }
  });
  if (aperta) aperta.giorno.incompleto = true;
  return giorni;
}

function descrizioneFasce_(g) {
  return g.fasce.map(function (f) {
    return (f.da || '?') + ' - ' + (f.a || '?');
  }).join('   |   ');
}

function sheetReport_(nome) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  sh.clear();
  return sh;
}

function aggiornaReport_() {
  var giorni = calcolaGiorni_(leggi_());
  var oggi = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var chiavi = Object.keys(giorni).sort().reverse(); // dal più recente
  var mesi = {};
  var ordineMesi = [];
  chiavi.forEach(function (k) {
    var m = k.substring(0, 7);
    if (!mesi[m]) { mesi[m] = []; ordineMesi.push(m); }
    mesi[m].push(k);
  });

  function nomeMese(m) {
    return MESI[parseInt(m.substring(5, 7), 10) - 1] + ' ' + m.substring(0, 4);
  }

  // --- Report mensile: un blocco per mese, giorno per giorno ---
  var sh = sheetReport_('Report mensile');
  var righe = [];
  var titoli = [], intestazioni = [], totali = [], anomalie = [];
  ordineMesi.forEach(function (m) {
    var riga = righe.length + 1;
    righe.push([nomeMese(m), '', '', '', '']);
    titoli.push('A' + riga + ':E' + riga);
    riga++;
    righe.push(['Data', 'Giorno', 'Timbrature', 'Ore lavorate', 'Note']);
    intestazioni.push('A' + riga + ':E' + riga);
    var minMese = 0, nGiorni = 0;
    mesi[m].forEach(function (k) {
      var g = giorni[k];
      var d = new Date(k + 'T12:00:00Z');
      var p = k.split('-');
      riga++;
      var inCorso = g.incompleto && k === oggi;
      righe.push([p[2] + '/' + p[1] + '/' + p[0], GIORNI[d.getUTCDay()], descrizioneFasce_(g),
                  g.minuti / 1440,
                  inCorso ? 'In corso' : (g.incompleto ? 'Timbratura mancante: controllare' : '')]);
      if (g.incompleto && !inCorso) anomalie.push('A' + riga + ':E' + riga);
      minMese += g.minuti;
      if (g.minuti > 0) nGiorni++;
    });
    riga++;
    righe.push(['Totale ' + nomeMese(m), '', nGiorni + ' giorni lavorati', minMese / 1440, '']);
    totali.push('A' + riga + ':E' + riga);
    righe.push(['', '', '', '', '']);
  });
  if (!righe.length) righe.push(['Nessuna timbratura registrata', '', '', '', '']);
  sh.getRange(1, 1, righe.length, 5).setValues(righe);
  sh.getRange(1, 4, righe.length, 1).setNumberFormat('[h]:mm');
  if (titoli.length) {
    sh.getRangeList(titoli).setFontWeight('bold').setFontSize(13).setBackground('#1e9e5a').setFontColor('#ffffff');
    sh.getRangeList(intestazioni).setFontWeight('bold').setBackground('#e8f0ec');
    sh.getRangeList(totali).setFontWeight('bold').setBackground('#f1f3f4');
  }
  if (anomalie.length) sh.getRangeList(anomalie).setBackground('#fde8e6');
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 280);
  sh.setColumnWidth(4, 110); sh.setColumnWidth(5, 240);

  // --- Riepilogo mesi: tabella mese per mese ---
  var rs = sheetReport_('Riepilogo mesi');
  var tab = [['Mese', 'Giorni lavorati', 'Ore totali (h:mm)', 'Ore totali (decimali)', 'Media ore al giorno']];
  ordineMesi.forEach(function (m, i) {
    var minMese = 0, nGiorni = 0;
    mesi[m].forEach(function (k) {
      minMese += giorni[k].minuti;
      if (giorni[k].minuti > 0) nGiorni++;
    });
    tab.push([nomeMese(m), nGiorni, minMese / 1440, Math.round(minMese / 60 * 100) / 100,
              nGiorni ? minMese / nGiorni / 1440 : 0]);
  });
  if (ordineMesi.length) {
    var tot = ordineMesi.length + 1;
    tab.push(['TOTALE', '=SUM(B2:B' + tot + ')', '=SUM(C2:C' + tot + ')', '=SUM(D2:D' + tot + ')', '']);
  }
  rs.getRange(1, 1, tab.length, 5).setValues(tab);
  rs.getRange(2, 3, Math.max(tab.length - 1, 1), 1).setNumberFormat('[h]:mm');
  rs.getRange(2, 5, Math.max(tab.length - 1, 1), 1).setNumberFormat('[h]:mm');
  rs.getRange(2, 4, Math.max(tab.length - 1, 1), 1).setNumberFormat('0.00');
  rs.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1e9e5a').setFontColor('#ffffff');
  if (ordineMesi.length) rs.getRange(tab.length, 1, 1, 5).setFontWeight('bold').setBackground('#f1f3f4');
  rs.setFrozenRows(1);
  rs.setColumnWidth(1, 150);
  for (var c = 2; c <= 5; c++) rs.setColumnWidth(c, 150);
}
