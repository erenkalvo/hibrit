// Hibrit kayıt betiği. Google Apps Script'e yapıştırılır, Web uygulaması olarak yayınlanır.
// Kayıtlar "Hibrit Kayıtlar" tablosuna yazılır.
const SHEET_ID = '1cGFZJHlWBME7gRG3EE1bSxgcdv0RxkduOZ7RFc9_PGY';
const H_SEANS = ['seans_id', 'tarih', 'seans', 'json', 'silindi'];
const H_SET = ['tarih', 'seans', 'hareket', 'set', 'kg', 'tekrar', 'bant', 'seans_id'];

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function sheet_(name, head) {
  const ss = ss_();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); s.setFrozenRows(1); }
  return s;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() {
  const rows = sheet_('Seanslar', H_SEANS).getDataRange().getValues().slice(1);
  const sessions = rows.filter(r => r[0] && !r[4]).map(r => JSON.parse(r[3]));
  const meta = {};
  sheet_('Ayarlar', ['anahtar', 'json']).getDataRange().getValues().slice(1)
    .forEach(r => { if (r[0]) meta[r[0]] = JSON.parse(r[1] || 'null'); });
  return out_({ ok: true, sessions: sessions, tests: meta.tests || {}, combos: meta.combos || {} });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    (b.sessions || []).forEach(addSession_);
    (b.del || []).forEach(deleteSession_);
    if (b.tests) setMeta_('tests', b.tests);
    if (b.combos) setMeta_('combos', b.combos);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function addSession_(rec) {
  const s = sheet_('Seanslar', H_SEANS);
  const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0]));
  if (ids.indexOf(String(rec.ts)) >= 0) return;
  s.appendRow([String(rec.ts), rec.date, rec.label || rec.type, JSON.stringify(rec), '']);
  const sets = [];
  Object.keys(rec.ex || {}).forEach(name => {
    const e = rec.ex[name];
    (e.s || []).forEach((x, i) => sets.push([rec.date, rec.label || rec.type, name, i + 1,
      x.kg == null ? '' : x.kg, x.r == null ? '' : x.r, e.band || '', String(rec.ts)]));
    if (!(e.s || []).length && e.band) sets.push([rec.date, rec.label || rec.type, name, '', '', '', e.band, String(rec.ts)]);
  });
  if (rec.cardio && (rec.cardio.dk || rec.cardio.km || rec.cardio.hr)) {
    sets.push([rec.date, rec.label || rec.type, 'Kardiyo', '', rec.cardio.km || '', rec.cardio.dk || '', rec.cardio.hr ? 'nabız ' + rec.cardio.hr : '', String(rec.ts)]);
  }
  if (sets.length) {
    const t = sheet_('Setler', H_SET);
    t.getRange(t.getLastRow() + 1, 1, sets.length, H_SET.length).setValues(sets);
  }
}

function deleteSession_(ts) {
  ts = String(ts);
  const s = sheet_('Seanslar', H_SEANS);
  const v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === ts) s.getRange(i + 1, 5).setValue('evet');
  const t = sheet_('Setler', H_SET);
  const w = t.getDataRange().getValues();
  for (let i = w.length - 1; i >= 1; i--) if (String(w[i][7]) === ts) t.deleteRow(i + 1);
}

function setMeta_(key, val) {
  const s = sheet_('Ayarlar', ['anahtar', 'json']);
  const v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === key) { s.getRange(i + 1, 2).setValue(JSON.stringify(val)); return; }
  s.appendRow([key, JSON.stringify(val)]);
}

// Kurulumda bir kez çalıştır: sayfaları oluşturur ve izin ister.
function kurulum() {
  sheet_('Seanslar', H_SEANS); sheet_('Setler', H_SET); sheet_('Ayarlar', ['anahtar', 'json']);
}
