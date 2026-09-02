// ─────────────────────────────────────────────
// TARJETAS — Google Apps Script Backend
// Paste this entire file into your Apps Script editor
// ─────────────────────────────────────────────

const SHEET_DECKS   = 'decks';
const SHEET_CARDS   = 'cards';
const SHEET_SCRIPTS = 'scripts';

// ── Entry point ──────────────────────────────
function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    const body = e.postData ? JSON.parse(e.postData.contents) : {};
    const action = (e.parameter && e.parameter.action) || body.action;
    let result;

    if      (action === 'read')      result = readAll();
    else if (action === 'write')     result = writeAll(body.data);
    else if (action === 'translate') result = translateText(e.parameter.text || body.text, e.parameter.from || body.from, e.parameter.to || body.to);
    else                             result = { error: 'Unknown action' };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Translate ─────────────────────────────────
function translateText(text, from, to) {
  if (!text) return { error: 'No text provided' };
  try {
    const srcLang = from || 'en';
    const tgtLang = to || 'es';
    const translated = LanguageApp.translate(text, srcLang, tgtLang);
    return { ok: true, result: translated, spanish: translated };
  } catch (err) {
    return { error: err.toString() };
  }
}

// ── Read all ──────────────────────────────────
function readAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheets(ss);

  // Decks
  const deckRows = ss.getSheetByName(SHEET_DECKS).getDataRange().getValues();
  const cardRows = ss.getSheetByName(SHEET_CARDS).getDataRange().getValues();

  const decks = deckRows.slice(1).map(r => ({
    id: r[0], name: r[1], order: Number(r[2]),
    lastStudied: r[3] ? new Date(r[3]).getTime() : null,
  })).sort((a, b) => a.order - b.order);

  const cards = cardRows.slice(1).map(r => ({
    id: r[0], deckId: r[1], en: r[2], es: r[3], ctx: r[4] || '',
  }));

  decks.forEach(d => {
    d.cards = cards.filter(c => c.deckId === d.id).map(c => ({
      id: c.id, en: c.en, es: c.es, ctx: c.ctx
    }));
  });

  // Scripts
  const scriptRows = ss.getSheetByName(SHEET_SCRIPTS).getDataRange().getValues();
  const scripts = scriptRows.slice(1).map((r, i) => ({
    id:          r[0],
    title:       r[1],
    text:        r[2],
    listenCount: Number(r[3]) || 0,
    order:       Number(r[4]) || i,
  })).sort((a, b) => a.order - b.order);

  return { ok: true, decks, scripts };
}

// ── Write all ─────────────────────────────────
function writeAll(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheets(ss);

  const decks   = data.decks   || [];
  const scripts = data.scripts || [];

  // Write decks
  const deckSheet = ss.getSheetByName(SHEET_DECKS);
  const cardSheet = ss.getSheetByName(SHEET_CARDS);
  clearBelowHeader(deckSheet);
  clearBelowHeader(cardSheet);

  if (decks.length) {
    deckSheet.getRange(2, 1, decks.length, 4).setValues(
      decks.map((d, i) => [d.id, d.name, i, d.lastStudied ? new Date(d.lastStudied).toISOString() : ''])
    );
  }
  const cardRows = [];
  decks.forEach(d => (d.cards || []).forEach(c => cardRows.push([c.id, d.id, c.en, c.es, c.ctx || ''])));
  if (cardRows.length) cardSheet.getRange(2, 1, cardRows.length, 5).setValues(cardRows);

  // Write scripts
  const scriptSheet = ss.getSheetByName(SHEET_SCRIPTS);
  clearBelowHeader(scriptSheet);
  if (scripts.length) {
    scriptSheet.getRange(2, 1, scripts.length, 5).setValues(
      scripts.map((s, i) => [s.id, s.title, (s.text || '').replace(/\n/g, '\\n'), s.listenCount || 0, i])
    );
  }

  return { ok: true };
}

// ── Helpers ───────────────────────────────────
function ensureSheets(ss) {
  if (!ss.getSheetByName(SHEET_DECKS)) {
    const s = ss.insertSheet(SHEET_DECKS);
    s.appendRow(['id','name','order','lastStudied']);
    s.getRange(1,1,1,4).setFontWeight('bold');
  }
  if (!ss.getSheetByName(SHEET_CARDS)) {
    const s = ss.insertSheet(SHEET_CARDS);
    s.appendRow(['id','deckId','english','spanish','context']);
    s.getRange(1,1,1,5).setFontWeight('bold');
  }
  if (!ss.getSheetByName(SHEET_SCRIPTS)) {
    const s = ss.insertSheet(SHEET_SCRIPTS);
    s.appendRow(['id','title','text','listenCount','order']);
    s.getRange(1,1,1,5).setFontWeight('bold');
  }
}

function clearBelowHeader(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow-1, sheet.getLastColumn()).clearContent();
}
