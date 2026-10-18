/**
 * Ma liste de courses : stockage dans Google Sheets.
 * À coller dans Extensions > Apps Script (voir README.md).
 * Le site lit et écrit tes recettes ici ; tu peux aussi les modifier à la main dans les onglets.
 */
const HEAD_R = ['Recette', 'Ingrédient', 'Quantité (pour 1 fois)', 'Unité', 'Rayon', 'id (ne pas toucher)'];
const HEAD_P = ['N°', 'Recette', 'Nb de fois'];
const HEAD_C = ['Articles cochés (ne pas toucher)'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sh_(name, head) {
  let s = ss_().getSheetByName(name);
  if (!s) {
    s = ss_().insertSheet(name);
    if (head) {
      s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
      s.setFrozenRows(1);
    }
  }
  return s;
}

// À lancer une fois (bouton ▶ Exécuter) pour créer les onglets.
function installer() {
  sh_('Recettes', HEAD_R); sh_('Semaine', HEAD_P); sh_('Cochés', HEAD_C); sh_('Meta');
}

function uid_() { return Utilities.getUuid().slice(0, 8); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ts_() { return Number(sh_('Meta').getRange('A1').getValue()) || 0; }
function room_(s, n) { const m = s.getMaxRows(); if (m < n) s.insertRowsAfter(m, n - m); }

// Une modification à la main dans le Sheet met à jour l'horodatage : le site la récupère.
function onEdit(e) {
  if (e && e.range && e.range.getSheet().getName() !== 'Meta') sh_('Meta').getRange('A1').setValue(Date.now());
}

function read_() {
  const rs = sh_('Recettes', HEAD_R);
  const rows = rs.getDataRange().getValues().slice(1);
  const ids = {};
  rows.forEach(r => { const n = String(r[0]).trim(), id = String(r[5]).trim(); if (n && id && !ids[n]) ids[n] = id; });
  const recipes = [], by = {}, fix = [];
  rows.forEach((r, i) => {
    const n = String(r[0]).trim(); if (!n) return;
    if (!ids[n]) ids[n] = uid_();
    if (String(r[5]).trim() !== ids[n]) fix.push([i + 2, ids[n]]);
    let rec = by[n];
    if (!rec) { rec = by[n] = { id: ids[n], name: n, ings: [] }; recipes.push(rec); }
    const g = String(r[1]).trim();
    if (g) {
      const q = r[2] === '' ? '' : Number(r[2]);
      rec.ings.push({ n: g, q: isNaN(q) ? '' : q, u: String(r[3]).trim(), r: String(r[4]).trim() });
    }
  });
  fix.forEach(f => rs.getRange(f[0], 6).setValue(f[1]));

  const plan = [];
  sh_('Semaine', HEAD_P).getDataRange().getValues().slice(1).forEach(r => {
    const name = String(r[1]).trim();
    if (r[0] === '' && name === '' && r[2] === '') return;
    plan.push({ rid: by[name] ? by[name].id : '', times: Math.max(1, Math.min(9, Number(r[2]) || 1)) });
  });
  if (!plan.length) plan.push({ rid: '', times: 1 });

  const checked = {};
  sh_('Cochés', HEAD_C).getDataRange().getValues().slice(1).forEach(r => { if (r[0] !== '') checked[String(r[0])] = true; });
  return { recipes: recipes, plan: plan, checked: checked, ts: ts_() };
}

function write_(d) {
  const rs = sh_('Recettes', HEAD_R), out = [];
  (d.recipes || []).forEach(r => {
    if (!r.ings || !r.ings.length) out.push([r.name, '', '', '', '', r.id]);
    else r.ings.forEach(g => out.push([r.name, g.n || '', (g.q === '' || g.q == null) ? '' : g.q, g.u || '', g.r || '', r.id]));
  });
  room_(rs, out.length + 1);
  rs.getRange(2, 1, Math.max(rs.getMaxRows() - 1, 1), 6).clearContent();
  if (out.length) rs.getRange(2, 1, out.length, 6).setValues(out);

  const names = {}; (d.recipes || []).forEach(r => names[r.id] = r.name);
  const ps = sh_('Semaine', HEAD_P), plan = (d.plan || []).map((p, i) => [i + 1, names[p.rid] || '', p.times || 1]);
  room_(ps, plan.length + 1);
  ps.getRange(2, 1, Math.max(ps.getMaxRows() - 1, 1), 3).clearContent();
  if (plan.length) ps.getRange(2, 1, plan.length, 3).setValues(plan);

  const cs = sh_('Cochés', HEAD_C), keys = Object.keys(d.checked || {}).filter(k => d.checked[k]).map(k => [k]);
  room_(cs, keys.length + 1);
  cs.getRange(2, 1, Math.max(cs.getMaxRows() - 1, 1), 1).clearContent();
  if (keys.length) cs.getRange(2, 1, keys.length, 1).setValues(keys);

  sh_('Meta').getRange('A1').setValue(d.ts || Date.now());
}

function doGet() { return json_(read_()); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if ((d.ts || 0) < ts_()) return json_({ ok: true, stale: true, data: read_() });
    write_(d);
    return json_({ ok: true, ts: d.ts });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
