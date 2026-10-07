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
    return { ts: new Date(r[0]), giorno: String(r[1]), ora: String(r[2]), tipo: r[3] };
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
  return stato();
}
