// =====================================================================
//  COMPÉTENCES — niveaux d'équipe, suivi de progression, seuils & règles
//  Fichier à AJOUTER au projet Apps Script (à côté de DashboardCode.gs).
//  Crée 3 onglets : « Niveaux Équipe », « Suivi Progression », « Seuils & Règles »
//  et alimente les 3 vues « Compétences » de la web app.
// =====================================================================

const CP_SHEETS = { team: 'Niveaux Équipe', prog: 'Suivi Progression', rules: 'Seuils & Règles' };
const CP_MIN_EVAL = 5;      // nb minimum de livrables pour évaluer un risque de descente
const CP_CC_WEEKS = 26;     // historique utilisé pour les cartes de contrôle
const CP_DEMO_SEED = 2;
const CP_ANOVA_ALPHA = 0.05;  // seuil de significativité de l'ANOVA
const CP_CC_SHOW = 150;     // nb max de points affichés par carte (les anomalies sont comptées sur ces points)

// Positions fixes dans « Seuils & Règles »
const CP_R = { lvlRow: 5, thrRow: 12, parRow: 18, nPar: 8 };

const CP_DEFAULT_LEVELS = [
  [1, 'Débutant', 'Travail supervisé : chaque livrable est contrôlé (crosscheck) avant envoi',
   'Gammes simples et standard, saisie des données, préparation des dossiers', 'Gamme Simple, Gamme Standard'],
  [2, 'Autonome', 'Livre seul sur les gammes courantes, crosscheck par un niveau 3-4 en échantillonnage',
   'Production gammes simples/standard/révisées, mise à jour du suivi', 'Gamme Simple, Gamme Standard, Gamme Révisée'],
  [3, 'Confirmé', 'Autonomie complète, peut faire le crosscheck des niveaux 1-2',
   'Toutes les gammes, crosscheck N1-N2, appui au reporting hebdo (TRM)', 'Gamme Simple à Gamme Complexe'],
  [4, 'Référent', 'Référent technique : valide les méthodes, forme l\'équipe, traite les cas complexes',
   'Toutes les gammes + Analyse Technique, formation, revue des NRFT, support FO/BO', 'Toutes les gammes']
];
// [niveau, volume min, RFT min %, OTD min %, efficacité min %, rework moy. max (h), RFT plancher %, OTD plancher %]
const CP_DEFAULT_THR = [
  [1, 0, 0, 0, 0, 99, 0, 0],
  [2, 10, 90, 90, 70, 0.40, 85, 85],
  [3, 20, 95, 95, 80, 0.20, 90, 90],
  [4, 30, 98, 97, 90, 0.10, 95, 95]
];
const CP_DEFAULT_PARAMS = [
  ['Fenêtre d\'évaluation (semaines)', 8, 'Période glissante sur laquelle les indicateurs sont calculés'],
  ['Responsable validation', 'Mouad', 'Personne qui valide les changements de niveau'],
  ['Email du responsable', 'mouad@example.com', 'Destinataire des alertes (valeur fictive à remplacer)'],
  ['Notifications email actives', 'NON', 'OUI = les alertes sont réellement envoyées par email'],
  ['Cartes de contrôle : sigma', 3, 'Largeur des limites de contrôle (moyenne ± k × écart-type)'],
  ['Cartes de contrôle : points minimum', 8, 'En dessous, la gamme est « données insuffisantes »'],
  ['Cartes de contrôle : longueur de série', 10, 'N points consécutifs du même côté de la moyenne = anomalie'],
  ['Mode démo (données fictives)', 'OUI', 'OUI = indicateurs simulés ; NON = calculés depuis l\'onglet DashBoard']
];
const CP_DUMMY_NAMES = ['Camille Roy', 'Lucas Bernard', 'Sofia Tremblay', 'Karim Haddad', 'Emma Gagnon', 'Noah Lefebvre', 'Inès Moreau', 'Thomas Pelletier'];
const CP_DUMMY_LEVELS = [4, 3, 3, 2, 2, 2, 1, 1];
// profils de démo : rft, otd, ratio (temps réel / target), livrables par semaine
const CP_DEMO_PROFILES = [
  { rft: 0.99, otd: 0.985, ratio: 0.95, perWeek: 5 },
  { rft: 0.965, otd: 0.965, ratio: 1.08, perWeek: 4.5 },
  { rft: 0.86, otd: 0.87, ratio: 1.3, perWeek: 4 },
  { rft: 0.985, otd: 0.99, ratio: 1.05, perWeek: 4.5 },
  { rft: 0.92, otd: 0.93, ratio: 1.25, perWeek: 3.5 },
  { rft: 0.91, otd: 0.94, ratio: 1.2, perWeek: 3 },
  { rft: 0.94, otd: 0.93, ratio: 1.2, perWeek: 3.5 },
  { rft: 0.84, otd: 0.88, ratio: 1.5, perWeek: 2.5 }
];
const CP_DEMO_GAMMES = [
  { name: 'Gamme Simple', target: 2, w: 3, minLevel: 1 },
  { name: 'Gamme Standard', target: 4, w: 3, minLevel: 1 },
  { name: 'Gamme Révisée', target: 5, w: 2, minLevel: 2 },
  { name: 'Gamme Complexe', target: 9, w: 1.5, minLevel: 3 },
  { name: 'Analyse Technique', target: 6, w: 0.12, minLevel: 4 }
];

// ---------------------------------------------------------------------
//  Menu
// ---------------------------------------------------------------------
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Compétences')
      .addItem('Initialiser les 3 onglets (données fictives)', 'menuSetupCompetency')
      .addItem('Actualiser « Suivi Progression »', 'menuRefreshCompetency')
      .addItem('Notifier le responsable (propositions en attente)', 'menuNotifyCompetency')
      .addToUi();
  } catch (e) {}
}
function menuSetupCompetency() {
  const r = setupCompetencyTabs();
  SpreadsheetApp.getUi().alert(r);
}
function menuRefreshCompetency() {
  const r = refreshCompetencyTabs();
  SpreadsheetApp.getUi().alert('Suivi Progression actualisé : ' + r.members + ' membre(s), ' + r.pending + ' proposition(s) en attente.');
}
function menuNotifyCompetency() {
  const r = notifyResponsible();
  SpreadsheetApp.getUi().alert(r.message);
}

// ---------------------------------------------------------------------
//  Utilitaires
// ---------------------------------------------------------------------
function cpIso_(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return '';
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}
function cpParseDate_(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v)) { const d = new Date(v.slice(0, 10) + 'T12:00:00'); return isNaN(d.getTime()) ? null : d; }
  return null;
}
function cpNum_(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function cpRound_(v, p) { const k = Math.pow(10, p || 1); return Math.round(v * k) / k; }
function cpRng_(seed) {            // générateur pseudo-aléatoire déterministe (mulberry32)
  let a = seed >>> 0;
  return function () {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
function cpGauss_(rnd) {
  let u = 0, v = 0;
  while (u === 0) u = rnd();
  while (v === 0) v = rnd();
  return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
}

// ---------------------------------------------------------------------
//  Lecture de la configuration et de l'équipe
// ---------------------------------------------------------------------
function cpReadRules_(ss) {
  const rules = { levels: [], thr: {}, params: {}, exists: false };
  const sh = ss.getSheetByName(CP_SHEETS.rules);
  const lv = CP_DEFAULT_LEVELS.map(r => r.slice());
  const th = CP_DEFAULT_THR.map(r => r.slice());
  const pa = CP_DEFAULT_PARAMS.map(r => r.slice());
  if (sh) {
    rules.exists = true;
    const a = sh.getRange(CP_R.lvlRow, 1, 4, 5).getValues();
    a.forEach((r, i) => { if (r[0] !== '') lv[i] = r; });
    const b = sh.getRange(CP_R.thrRow, 1, 4, 8).getValues();
    b.forEach((r, i) => { if (r[0] !== '') th[i] = r; });
    const c = sh.getRange(CP_R.parRow, 1, CP_R.nPar, 3).getValues();
    c.forEach((r, i) => { if (r[1] !== '') pa[i] = r; });
  }
  rules.levels = lv.map(r => ({ level: cpNum_(r[0]), name: String(r[1]), rights: String(r[2]), tasks: String(r[3]), gammes: String(r[4]) }));
  th.forEach(r => {
    rules.thr[cpNum_(r[0])] = { vol: cpNum_(r[1]), rft: cpNum_(r[2]), otd: cpNum_(r[3]), eff: cpNum_(r[4]), rework: cpNum_(r[5]), floorRft: cpNum_(r[6]), floorOtd: cpNum_(r[7]) };
  });
  rules.params = {
    window: Math.max(1, cpNum_(pa[0][1]) || 8), resp: String(pa[1][1]), email: String(pa[2][1]),
    notify: String(pa[3][1]).toUpperCase() === 'OUI', sigma: cpNum_(pa[4][1]) || 3,
    minPts: Math.max(3, cpNum_(pa[5][1]) || 8), run: Math.max(3, cpNum_(pa[6][1]) || 10),
    demo: String(pa[7][1]).toUpperCase() === 'OUI'
  };
  return rules;
}

function cpDashboardNames_(ss) {
  const sh = ss.getSheetByName('DashBoard');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (!data.length) return [];
  const col = data[0].findIndex(h => h && h.toString().trim() === 'Execution');
  if (col < 0) return [];
  const count = {};
  for (let i = 1; i < data.length; i++) {
    const v = data[i][col];
    if (v && typeof v === 'string' && v.trim() !== '') count[v.trim()] = (count[v.trim()] || 0) + 1;
  }
  return Object.keys(count).sort((a, b) => count[b] - count[a]);
}

function cpDefaultTeam_(ss) {
  const real = cpDashboardNames_(ss).slice(0, 8);
  const names = real.length >= 3 ? real : CP_DUMMY_NAMES;
  return names.map((n, i) => ({
    name: n, team: i % 3 === 2 ? 'BO' : 'FO', level: CP_DUMMY_LEVELS[i] || 1,
    since: '2026-0' + (1 + (i % 6)) + '-01', proposal: '—', decision: '—', comment: '', lastEval: '', row: 0
  }));
}

function cpReadTeam_(ss) {
  const sh = ss.getSheetByName(CP_SHEETS.team);
  if (!sh) return { team: cpDefaultTeam_(ss), exists: false };
  const last = sh.getLastRow();
  const out = [];
  if (last >= 4) {
    const v = sh.getRange(4, 1, last - 3, 10).getValues();
    v.forEach((r, i) => {
      if (String(r[0]).trim() === '') return;
      out.push({
        name: String(r[0]).trim(), team: String(r[1] || '').trim(), level: Math.min(4, Math.max(1, cpNum_(r[2]) || 1)),
        since: cpIso_(r[5] instanceof Date ? r[5] : null) || String(r[5] || ''), proposal: String(r[6] || '—') || '—',
        decision: String(r[7] || '—') || '—', comment: String(r[8] || ''), lastEval: cpIso_(r[9] instanceof Date ? r[9] : null) || String(r[9] || ''), row: i + 4
      });
    });
  }
  return { team: out, exists: true };
}

// ---------------------------------------------------------------------
//  Livrables : lecture réelle (onglet DashBoard) ou génération fictive
// ---------------------------------------------------------------------
function cpLoadItems_(ss) {
  const sh = ss.getSheetByName('DashBoard');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const col = n => h.findIndex(x => x && x.toString().trim() === n);
  const c = {
    status: col('Status'), date: col('Delivery Date'), otd: col('OTD'), rft: col('RFT'),
    pT: col('Production Time Target (Hours)'), cT: col('Crosscheck Time Target (Hours)'),
    pA: col('Production Time (Hours)'), cA: col('Crosscheck Time (Hours)'),
    type: col('Type of deliverable'), exec: col('Execution'), rework: col('Rework Time (Hours)')
  };
  const items = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (r[c.status] !== 'Delivered') continue;
    const d = r[c.date];
    if (!(d instanceof Date)) continue;
    const ex = r[c.exec];
    if (!ex || typeof ex !== 'string' || ex.trim() === '') continue;
    const t = r[c.type];
    items.push({
      date: d, exec: ex.trim(), type: (t && t.toString().trim() !== '') ? t.toString().trim() : 'Non défini',
      rft: r[c.rft] === true ? true : (r[c.rft] === false ? false : null),
      otd: r[c.otd] === true ? true : (r[c.otd] === false ? false : null),
      target: (c.pT > -1 ? cpNum_(r[c.pT]) : 0) + (c.cT > -1 ? cpNum_(r[c.cT]) : 0),
      actual: (c.pA > -1 ? cpNum_(r[c.pA]) : 0) + (c.cA > -1 ? cpNum_(r[c.cA]) : 0),
      rework: c.rework > -1 ? cpNum_(r[c.rework]) : 0
    });
  }
  return items;
}

function cpDemoItems_(team, today) {
  const rnd = cpRng_(CP_DEMO_SEED);
  const items = [];
  team.forEach((m, idx) => {
    const p = CP_DEMO_PROFILES[idx % CP_DEMO_PROFILES.length];
    let seq = 0;
    const phR = rnd(), phO = rnd();
    const hit = (p, ph, k) => Math.floor((k + 1) * (1 - p) + ph) === Math.floor(k * (1 - p) + ph);   // taux exact, échecs régulièrement espacés
    const allowed = CP_DEMO_GAMMES.filter(g => g.minLevel <= Math.min(4, m.level + 1));
    const totalW = allowed.reduce((s, g) => s + g.w, 0);
    for (let w = 0; w < CP_CC_WEEKS; w++) {
      const n = Math.max(0, Math.round(p.perWeek + (rnd() - 0.5) * 3));
      for (let k = 0; k < n; k++) {
        let x = rnd() * totalW, g = allowed[0];
        for (let j = 0; j < allowed.length; j++) { x -= allowed[j].w; if (x <= 0) { g = allowed[j]; break; } }
        const d = new Date(today.getTime() - (w * 7 + Math.floor(rnd() * 5)) * 86400000);
        const target = Math.max(0.5, g.target * (0.85 + rnd() * 0.3));
        let ratio = Math.max(0.5, p.ratio + cpGauss_(rnd) * 0.12);
        if (g.name === 'Gamme Complexe' && w <= 2 && rnd() < 0.35) ratio *= 2.1;      // anomalie injectée (démo)
        const rft = hit(p.rft, phR, seq), otd = hit(p.otd, phO, seq);
        seq++;
        items.push({
          date: d, exec: m.name, type: g.name, rft: rft, otd: otd,
          target: target, actual: target * ratio, rework: rft ? 0 : (1 + rnd() * 2.4)
        });
      }
    }
  });
  return items;
}

// ---------------------------------------------------------------------
//  Évaluation
// ---------------------------------------------------------------------
function cpCriteria_(stats, req) {
  const mk = (key, label, value, need, dir, unit) => ({ key: key, label: label, value: value, need: need, dir: dir, unit: unit, ok: dir === '>=' ? value >= need : value <= need });
  return [
    mk('vol', 'Volume', stats.n, req.vol, '>=', ''),
    mk('rft', 'RFT', stats.rftPct, req.rft, '>=', '%'),
    mk('otd', 'OTD', stats.otdPct, req.otd, '>=', '%'),
    mk('eff', 'Efficacité', stats.effPct, req.eff, '>=', '%'),
    mk('rework', 'Rework moy.', stats.reworkAvg, req.rework, '<=', 'h')
  ];
}

function cpMemberStats_(list) {
  let rftT = 0, rftF = 0, otdT = 0, otdF = 0, tg = 0, ac = 0, rw = 0;
  list.forEach(it => {
    if (it.rft === true) rftT++; else if (it.rft === false) rftF++;
    if (it.otd === true) otdT++; else if (it.otd === false) otdF++;
    tg += it.target; ac += it.actual; rw += it.rework;
  });
  const n = list.length;
  return {
    n: n,
    rftPct: (rftT + rftF) > 0 ? cpRound_(rftT / (rftT + rftF) * 100, 1) : 0,
    otdPct: (otdT + otdF) > 0 ? cpRound_(otdT / (otdT + otdF) * 100, 1) : 0,
    effPct: ac > 0 ? cpRound_(tg / ac * 100, 1) : 0,
    reworkAvg: n > 0 ? cpRound_(rw / n, 2) : 0
  };
}

function cpEvaluateMember_(m, stats, rules) {
  const L = m.level;
  const res = { status: 'stable', proposal: { type: 'none', to: L }, criteria: [], progress: null, gap: '' };
  if (L < 4) {
    res.criteria = cpCriteria_(stats, rules.thr[L + 1] || { vol: 0, rft: 0, otd: 0, eff: 0, rework: 99 });
    const ok = res.criteria.filter(c => c.ok).length;
    res.progress = Math.round(ok / res.criteria.length * 100);
    const failing = res.criteria.filter(c => !c.ok);
    res.gap = failing.length ? failing.slice(0, 2).map(c => c.label + ' ' + cpRound_(c.value, 2) + c.unit + (c.dir === '>=' ? ' < ' : ' > ') + c.need + c.unit).join(' · ') : 'Tous les critères du niveau ' + (L + 1) + ' sont atteints';
  } else {
    res.gap = 'Niveau maximum';
  }
  const floor = rules.thr[L];
  if (stats.n === 0) { res.status = 'donnees'; res.gap = 'Aucun livrable sur la fenêtre'; return res; }
  if (L >= 2 && stats.n >= CP_MIN_EVAL && floor && (stats.rftPct < floor.floorRft || stats.otdPct < floor.floorOtd)) {
    res.status = 'risque';
    res.proposal = { type: 'down', to: L - 1 };
    const parts = [];
    if (stats.rftPct < floor.floorRft) parts.push('RFT ' + stats.rftPct + '% < plancher ' + floor.floorRft + '%');
    if (stats.otdPct < floor.floorOtd) parts.push('OTD ' + stats.otdPct + '% < plancher ' + floor.floorOtd + '%');
    res.gap = parts.join(' · ');
  } else if (L < 4 && res.criteria.every(c => c.ok)) {
    res.status = 'promotion';
    res.proposal = { type: 'up', to: L + 1 };
  }
  return res;
}

function cpProposalText_(p) {
  return p.type === 'up' ? 'Montée → N' + p.to : (p.type === 'down' ? 'Descente → N' + p.to : '—');
}

function cpControlChart_(gamme, items, rules) {
  const k = rules.params.sigma;
  const pts = items.filter(it => it.target > 0 && it.actual > 0).sort((a, b) => a.date - b.date)
    .map(it => ({ date: cpIso_(it.date), exec: it.exec, v: cpRound_(it.actual / it.target * 100, 1), flag: '' }));
  const out = { gamme: gamme, n: pts.length, mean: 0, sd: 0, ucl: 0, lcl: 0, status: 'insuffisant', reasons: [], points: [], outCount: 0, runCount: 0 };
  if (pts.length < rules.params.minPts) { out.points = pts.slice(-CP_CC_SHOW); return out; }
  const stat = arr => {
    const m = arr.reduce((s, x) => s + x, 0) / arr.length;
    const sd = arr.length > 1 ? Math.sqrt(arr.reduce((s, x) => s + (x - m) * (x - m), 0) / (arr.length - 1)) : 0;
    return { m: m, sd: sd };
  };
  let s = stat(pts.map(p => p.v));
  let base = pts.filter(p => Math.abs(p.v - s.m) <= k * s.sd).map(p => p.v);       // 2e passe : limites sans les points aberrants
  if (base.length >= 3) s = stat(base);
  out.mean = cpRound_(s.m, 1); out.sd = cpRound_(s.sd, 1);
  out.ucl = cpRound_(s.m + k * s.sd, 1); out.lcl = cpRound_(Math.max(0, s.m - k * s.sd), 1);
  pts.forEach(p => { if (p.v > out.ucl) { p.flag = 'high'; out.outCount++; } else if (p.v < out.lcl) { p.flag = 'low'; out.outCount++; } });
  let runLen = 0, side = 0;
  pts.forEach((p, i) => {
    const sd = p.v > s.m ? 1 : (p.v < s.m ? -1 : 0);
    if (sd !== 0 && sd === side) runLen++; else { runLen = sd !== 0 ? 1 : 0; side = sd; }
    if (runLen >= rules.params.run) {
      out.runCount++;
      for (let j = i - runLen + 1; j <= i; j++) if (!pts[j].flag) pts[j].flag = 'run';
    }
  });
  out.points = pts.slice(-CP_CC_SHOW);
  // Les anomalies sont comptées sur les points AFFICHÉS, pour que le chiffre corresponde à ce que l'on voit
  out.outCount = out.points.filter(p => p.flag === 'high' || p.flag === 'low').length;
  out.runCount = out.points.filter(p => p.flag === 'run').length;
  if (out.outCount) out.reasons.push(out.outCount + ' point(s) hors limites sur ' + out.points.length);
  if (out.runCount) out.reasons.push('série de ' + rules.params.run + '+ points du même côté de la moyenne (' + out.runCount + ' points)');
  out.status = out.reasons.length ? 'hors' : 'ok';
  return out;
}


// ---------------------------------------------------------------------
//  ANOVA à un facteur : le ratio temps réel ÷ target diffère-t-il entre les types (groupes) de gammes ?
// ---------------------------------------------------------------------
function cpLogGamma_(x) {
  const c = [76.18009172947146, -86.50532032941677, 24.01409824083091, -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
  let y = x, t = x + 5.5; t -= (x + 0.5) * Math.log(t);
  let ser = 1.000000000190015;
  for (let j = 0; j < 6; j++) ser += c[j] / ++y;
  return -t + Math.log(2.5066282746310005 * ser / x);
}
function cpBetaCf_(a, b, x) {
  const FPMIN = 1e-30;
  let qab = a + b, qap = a + 1, qam = a - 1, c = 1, d = 1 - qab * x / qap;
  if (Math.abs(d) < FPMIN) d = FPMIN;
  d = 1 / d; let h = d;
  for (let m = 1; m <= 300; m++) {
    const m2 = 2 * m;
    let aa = m * (b - m) * x / ((qam + m2) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d; h *= d * c;
    aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const del = d * c; h *= del;
    if (Math.abs(del - 1) < 3e-12) break;
  }
  return h;
}
function cpIncBeta_(a, b, x) {          // bêta incomplète régularisée I_x(a, b)
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const bt = Math.exp(cpLogGamma_(a + b) - cpLogGamma_(a) - cpLogGamma_(b) + a * Math.log(x) + b * Math.log(1 - x));
  return x < (a + 1) / (a + b + 2) ? bt * cpBetaCf_(a, b, x) / a : 1 - bt * cpBetaCf_(b, a, 1 - x) / b;
}
function cpFPvalue_(F, d1, d2) { return cpIncBeta_(d2 / 2, d1 / 2, d2 / (d2 + d1 * F)); }
function cpTPvalue_(t, df) { return cpIncBeta_(df / 2, 0.5, df / (df + t * t)); }   // bilatérale

function cpAnova_(ccItems, rules) {
  const minN = Math.max(2, rules.params.minPts);
  const byType = {};
  ccItems.forEach(it => {
    if (!(it.target > 0 && it.actual > 0)) return;
    (byType[it.type] = byType[it.type] || []).push(it.actual / it.target * 100);
  });
  const names = Object.keys(byType);
  const used = names.filter(g => byType[g].length >= minN);
  const excluded = names.filter(g => byType[g].length < minN).map(g => ({ gamme: g, n: byType[g].length }));
  const out = { ok: false, alpha: CP_ANOVA_ALPHA, minN: minN, groups: [], excluded: excluded, k: used.length, N: 0, dfB: 0, dfW: 0,
                F: 0, p: 1, eta2: 0, msb: 0, ssb: 0, ssw: 0, significant: false, grand: 0, mse: 0, pairs: [], varRatio: 1, conclusion: '' };
  if (used.length < 2) { out.conclusion = 'Il faut au moins 2 types de gammes avec ' + minN + ' points chacun pour lancer l’ANOVA.'; return out; }
  let N = 0, sum = 0;
  used.forEach(g => { N += byType[g].length; byType[g].forEach(v => { sum += v; }); });
  const grand = sum / N;
  let ssb = 0, ssw = 0;
  const groups = used.map(g => {
    const a = byType[g], n = a.length, m = a.reduce((x, y) => x + y, 0) / n;
    const ss = a.reduce((x, y) => x + (y - m) * (y - m), 0);
    ssb += n * (m - grand) * (m - grand); ssw += ss;
    return { gamme: g, n: n, mean: m, sd: n > 1 ? Math.sqrt(ss / (n - 1)) : 0 };
  });
  const dfB = used.length - 1, dfW = N - used.length;
  if (dfW <= 0 || ssw <= 0) { out.conclusion = 'Variance intra-groupe nulle : ANOVA impossible.'; return out; }
  const mse = ssw / dfW, F = (ssb / dfB) / mse, pv = cpFPvalue_(F, dfB, dfW);
  groups.forEach(x => { x.se = Math.sqrt(mse / x.n); x.lo = x.mean - 1.96 * x.se; x.hi = x.mean + 1.96 * x.se; });
  // comparaisons deux à deux (t sur la variance résiduelle, correction de Bonferroni)
  const pairs = [], nPairs = used.length * (used.length - 1) / 2;
  for (let i = 0; i < groups.length; i++) for (let j = i + 1; j < groups.length; j++) {
    const a = groups[i], b = groups[j];
    const t = (a.mean - b.mean) / Math.sqrt(mse * (1 / a.n + 1 / b.n));
    const pr = Math.min(1, cpTPvalue_(t, dfW) * nPairs);
    pairs.push({ a: a.gamme, b: b.gamme, diff: cpRound_(a.mean - b.mean, 1), p: pr, significant: pr < CP_ANOVA_ALPHA });
  }
  pairs.sort((x, y) => x.p - y.p);
  const sds = groups.map(x => x.sd).filter(x => x > 0);
  out.ok = true; out.N = N; out.dfB = dfB; out.dfW = dfW; out.F = cpRound_(F, 2); out.p = pv; out.eta2 = cpRound_(ssb / (ssb + ssw), 3);
  out.significant = pv < CP_ANOVA_ALPHA; out.grand = cpRound_(grand, 1); out.mse = cpRound_(mse, 1);
  const sample = a => { if (a.length <= 120) return a.map(v => cpRound_(v, 1)); const o = []; for (let i = 0; i < 120; i++) o.push(cpRound_(a[Math.floor(i * a.length / 120)], 1)); return o; };   // échantillon pour le nuage de points
  out.groups = groups.map(x => ({ gamme: x.gamme, n: x.n, mean: cpRound_(x.mean, 1), sd: cpRound_(x.sd, 1), lo: cpRound_(x.lo, 1), hi: cpRound_(x.hi, 1), pts: sample(byType[x.gamme]) }));
  out.msb = cpRound_(ssb / dfB, 1); out.ssb = cpRound_(ssb, 0); out.ssw = cpRound_(ssw, 0);
  out.pairs = pairs.map(x => ({ a: x.a, b: x.b, diff: x.diff, p: x.p, significant: x.significant }));
  out.varRatio = sds.length > 1 ? cpRound_(Math.max.apply(null, sds) / Math.min.apply(null, sds), 2) : 1;
  const sig = pairs.filter(x => x.significant);
  out.conclusion = out.significant
    ? 'Différence significative entre les types de gammes (p ' + (pv < 0.001 ? '< 0,001' : '= ' + pv.toFixed(3).replace('.', ',')) + ').' + (sig.length ? ' Paires distinctes : ' + sig.map(x => x.a + ' ≠ ' + x.b).join(', ') + '.' : '')
    : 'Pas de différence significative entre les types de gammes (p = ' + pv.toFixed(3).replace('.', ',') + ') : les gammes se comportent de façon comparable.';
  return out;
}

// ---------------------------------------------------------------------
//  Calcul global (utilisé par la feuille ET par la web app)
// ---------------------------------------------------------------------
function cpCompute_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rules = cpReadRules_(ss);
  const tr = cpReadTeam_(ss);
  const team = tr.team;
  const today = new Date();
  const demo = rules.params.demo || (!rules.exists && !ss.getSheetByName('DashBoard'));
  const all = demo ? cpDemoItems_(team, today) : cpLoadItems_(ss);
  const from = new Date(today.getTime() - rules.params.window * 7 * 86400000);
  const to = new Date(today.getTime() + 86400000);
  const win = all.filter(it => it.date >= from && it.date <= to);
  const key = s => String(s).trim().toLowerCase();

  const counts = {};
  win.forEach(it => { counts[it.type] = (counts[it.type] || 0) + 1; });
  const gammes = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);

  const members = team.map(m => {
    const mine = win.filter(it => key(it.exec) === key(m.name));
    const stats = cpMemberStats_(mine);
    const ev = cpEvaluateMember_(m, stats, rules);
    const newProp = cpProposalText_(ev.proposal);
    let decision = m.decision;
    if (newProp !== m.proposal) decision = newProp === '—' ? '—' : 'En attente';
    const perGamme = {};
    gammes.forEach(g => {
      const sub = mine.filter(it => it.type === g);
      if (sub.length) { const s = cpMemberStats_(sub); perGamme[g] = { n: s.n, rftPct: s.rftPct, effPct: s.effPct }; }
    });
    const lv = rules.levels.filter(l => l.level === m.level)[0] || { name: '', rights: '', tasks: '', gammes: '' };
    return {
      name: m.name, team: m.team, level: m.level, levelName: lv.name, rights: lv.rights, since: m.since, row: m.row,
      stats: stats, status: ev.status, criteria: ev.criteria, progress: ev.progress, gap: ev.gap,
      proposal: newProp, proposalType: ev.proposal.type, proposalTo: ev.proposal.to, decision: decision, comment: m.comment, perGamme: perGamme
    };
  });

  const ccFrom = new Date(today.getTime() - CP_CC_WEEKS * 7 * 86400000);
  const ccItems = all.filter(it => it.date >= ccFrom && it.date <= to);
  const ccGammes = {};
  ccItems.forEach(it => { ccGammes[it.type] = true; });
  const control = Object.keys(ccGammes).sort().map(g => cpControlChart_(g, ccItems.filter(it => it.type === g), rules));

  return {
    demo: demo, setupDone: rules.exists && tr.exists, generatedAt: new Date().toISOString(),
    window: rules.params.window, params: rules.params, levels: rules.levels, thresholds: rules.thr,
    gammes: gammes, members: members, control: control, anova: cpAnova_(ccItems, rules),
    pending: members.filter(m => m.proposalType !== 'none' && m.decision === 'En attente').length
  };
}

function getCompetencyData() {
  return cpCompute_();
}

// ---------------------------------------------------------------------
//  Création des 3 onglets (données fictives)
// ---------------------------------------------------------------------
function cpHeader_(range) {
  range.setBackground('#1c5cab').setFontColor('#ffffff').setFontWeight('bold').setVerticalAlignment('middle').setWrap(true);
}
function cpTitle_(sh, text, sub) {
  sh.getRange(1, 1).setValue(text).setFontSize(16).setFontWeight('bold').setFontColor('#0f172a');
  if (sub !== undefined) sh.getRange(2, 1).setValue(sub).setFontColor('#64748b').setFontStyle('italic');
}
function cpSection_(sh, row, text, width) {
  sh.getRange(row, 1, 1, width).setBackground('#e8f1fc').setFontColor('#1c5cab').setFontWeight('bold');
  sh.getRange(row, 1).setValue(text);
}

function setupCompetencyTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];

  // ---- Seuils & Règles ----
  let sh = ss.getSheetByName(CP_SHEETS.rules);
  if (!sh) {
    sh = ss.insertSheet(CP_SHEETS.rules);
    cpTitle_(sh, 'Seuils & Règles — niveaux de compétence', 'Valeurs fictives (hypothèses) à remplacer par les seuils fournis par Mouad. Les cellules bleu clair sont modifiables ; la web app lit cet onglet.');
    cpSection_(sh, 3, '1. Définition des niveaux (1 = le plus bas, 4 = le plus haut)', 8);
    sh.getRange(4, 1, 1, 5).setValues([['Niveau', 'Intitulé', 'Droits', 'Tâches autorisées', 'Gammes accessibles']]);
    cpHeader_(sh.getRange(4, 1, 1, 5));
    sh.getRange(CP_R.lvlRow, 1, 4, 5).setValues(CP_DEFAULT_LEVELS);
    cpSection_(sh, 10, '2. Seuils de changement de niveau (fournis par Mouad) — indicateurs calculés sur la fenêtre d\'évaluation', 8);
    sh.getRange(11, 1, 1, 8).setValues([['Niveau', 'Volume min. (livrables)', 'RFT min. % (pour atteindre ce niveau)', 'OTD min. % (pour atteindre ce niveau)', 'Efficacité min. % (target ÷ réel)', 'Rework moy. max (h / livrable)', 'RFT plancher % (pour rester à ce niveau)', 'OTD plancher % (pour rester à ce niveau)']]);
    cpHeader_(sh.getRange(11, 1, 1, 8));
    sh.getRange(CP_R.thrRow, 1, 4, 8).setValues(CP_DEFAULT_THR);
    cpSection_(sh, 17, '3. Paramètres généraux', 8);
    sh.getRange(CP_R.parRow, 1, CP_R.nPar, 3).setValues(CP_DEFAULT_PARAMS);
    sh.getRange(CP_R.lvlRow, 2, 4, 4).setBackground('#eef5fd').setWrap(true).setVerticalAlignment('top');
    sh.getRange(CP_R.thrRow, 2, 4, 7).setBackground('#eef5fd').setHorizontalAlignment('center');
    sh.getRange(CP_R.parRow, 2, CP_R.nPar, 1).setBackground('#eef5fd');
    sh.getRange(CP_R.parRow + 3, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['OUI', 'NON'], true).build());
    sh.getRange(CP_R.parRow + 7, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['OUI', 'NON'], true).build());
    sh.getRange(27, 1).setValue('Règles d\'évaluation').setFontWeight('bold');
    sh.getRange(28, 1).setValue('• Montée : tous les critères du niveau SUIVANT (volume, RFT, OTD, efficacité, rework) sont atteints sur la fenêtre → proposition envoyée à Mouad.   • Descente : RFT ou OTD sous les planchers du niveau ACTUEL (au moins ' + CP_MIN_EVAL + ' livrables) → proposition de descente.   • Rien ne change sans validation de Mouad.');
    sh.getRange(29, 1).setValue('• Cartes de contrôle par gamme : ratio temps réel ÷ target (%) de chaque livrable ; limites = moyenne ± k σ (calculées sans les points aberrants) ; anomalie = point hors limites ou série de N points du même côté de la moyenne.');
    sh.getRange(28, 1, 2, 1).setWrap(false).setFontColor('#475569');
    sh.setColumnWidths(1, 1, 210); sh.setColumnWidths(2, 1, 170); sh.setColumnWidths(3, 1, 330); sh.setColumnWidths(4, 1, 330); sh.setColumnWidths(5, 4, 200);
    sh.setRowHeight(11, 48);
    sh.setFrozenRows(2);
    created.push(CP_SHEETS.rules);
  }

  // ---- Niveaux Équipe ----
  let st = ss.getSheetByName(CP_SHEETS.team);
  if (!st) {
    st = ss.insertSheet(CP_SHEETS.team);
    cpTitle_(st, 'Niveaux Équipe — membres et niveaux de compétence', 'Source de vérité des niveaux (1 à 4). « Proposition » et « Dernière évaluation » sont mises à jour par Suivi Progression ; « Décision » est à renseigner par Mouad.');
    st.getRange(3, 1, 1, 10).setValues([['Membre', 'Équipe', 'Niveau (1-4)', 'Intitulé', 'Droits', 'Niveau depuis le', 'Proposition', 'Décision', 'Commentaire', 'Dernière évaluation']]);
    cpHeader_(st.getRange(3, 1, 1, 10));
    const team = cpDefaultTeam_(ss);
    const rows = team.map((m, i) => [m.name, m.team, m.level, '', '', cpParseDate_(m.since), '—', '—', '', '']);
    st.getRange(4, 1, rows.length, 10).setValues(rows);
    const lr = rows.length;
    for (let i = 0; i < lr; i++) {
      const r = 4 + i;
      st.getRange(r, 4).setFormula("=IFERROR(INDEX('" + CP_SHEETS.rules + "'!$B$5:$B$8,MATCH(C" + r + ",'" + CP_SHEETS.rules + "'!$A$5:$A$8,0)),\"\")");
      st.getRange(r, 5).setFormula("=IFERROR(INDEX('" + CP_SHEETS.rules + "'!$C$5:$C$8,MATCH(C" + r + ",'" + CP_SHEETS.rules + "'!$A$5:$A$8,0)),\"\")");
    }
    st.getRange(4, 3, 60, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(1, 4).setAllowInvalid(false).build());
    st.getRange(4, 2, 60, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['FO', 'BO'], true).build());
    st.getRange(4, 8, 60, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['En attente', 'Validé', 'Refusé', '—'], true).build());
    st.getRange(4, 6, 60, 1).setNumberFormat('dd/mm/yyyy');
    st.getRange(4, 10, 60, 1).setNumberFormat('dd/mm/yyyy');
    st.getRange(4, 3, 60, 1).setHorizontalAlignment('center').setFontWeight('bold');
    st.getRange(4, 8, 60, 1).setBackground('#fff8e6');
    st.getRange(4, 9, 60, 1).setBackground('#fff8e6');
    st.setColumnWidths(1, 1, 200); st.setColumnWidths(2, 1, 70); st.setColumnWidths(3, 1, 90); st.setColumnWidths(4, 1, 110);
    st.setColumnWidths(5, 1, 420); st.setColumnWidths(6, 1, 110); st.setColumnWidths(7, 1, 130); st.setColumnWidths(8, 1, 110); st.setColumnWidths(9, 1, 300); st.setColumnWidths(10, 1, 130);
    st.setFrozenRows(3);
    created.push(CP_SHEETS.team);
  }

  // ---- Suivi Progression ----
  let sp = ss.getSheetByName(CP_SHEETS.prog);
  if (!sp) { sp = ss.insertSheet(CP_SHEETS.prog); created.push(CP_SHEETS.prog); }

  const r = refreshCompetencyTabs();
  return (created.length ? 'Onglets créés : ' + created.join(', ') + '.' : 'Les 3 onglets existent déjà (non modifiés).') + '\n' + r.members + ' membre(s) évalué(s).\nMode démo : ' + (r.demo ? 'OUI (données fictives)' : 'NON') + '.';
}

// ---------------------------------------------------------------------
//  Écriture de « Suivi Progression » + mise à jour de « Niveaux Équipe »
// ---------------------------------------------------------------------
const CP_STATUS_LABEL = { promotion: 'Éligible à la montée', risque: 'Risque de descente', stable: 'Stable', donnees: 'Données insuffisantes' };
const CP_STATUS_COLOR = { promotion: '#e6f6e6', risque: '#fde7e7', stable: '#eef5fd', donnees: '#f1f5f9' };

function refreshCompetencyTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = cpCompute_();
    let sp = ss.getSheetByName(CP_SHEETS.prog);
    if (!sp) sp = ss.insertSheet(CP_SHEETS.prog);
    sp.clear();
    const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm');
    cpTitle_(sp, 'Suivi Progression — indicateurs vs seuils', 'Actualisé le ' + stamp + ' · fenêtre : ' + d.window + ' semaines · ' + (d.demo ? 'MODE DÉMO (données fictives)' : 'données réelles (onglet DashBoard)') + ' · valeurs écrites par script (menu Compétences ▸ Actualiser)');

    // Bloc 1 : membres
    cpSection_(sp, 3, '1. Progression de chaque membre vers le niveau suivant', 12);
    sp.getRange(4, 1, 1, 12).setValues([['Membre', 'Équipe', 'Niveau', 'Volume', 'RFT %', 'OTD %', 'Efficacité %', 'Rework moy. (h)', 'Critères du niveau suivant', 'Progression %', 'Statut', 'Écart / prochaine étape']]);
    cpHeader_(sp.getRange(4, 1, 1, 12));
    const rows = d.members.map(m => [m.name, m.team, m.level, m.stats.n, m.stats.rftPct, m.stats.otdPct, m.stats.effPct, m.stats.reworkAvg,
      m.criteria.length ? m.criteria.filter(c => c.ok).length + ' / ' + m.criteria.length : '—', m.progress == null ? '—' : m.progress, CP_STATUS_LABEL[m.status], m.gap]);
    if (rows.length) {
      sp.getRange(5, 1, rows.length, 12).setValues(rows);
      sp.getRange(5, 5, rows.length, 3).setNumberFormat('0.0"%"');
      sp.getRange(5, 10, rows.length, 1).setNumberFormat('0"%"');
      sp.getRange(5, 3, rows.length, 8).setHorizontalAlignment('center');
      d.members.forEach((m, i) => {
        sp.getRange(5 + i, 11).setBackground(CP_STATUS_COLOR[m.status]).setFontWeight('bold');
        m.criteria.forEach(c => {
          const idx = { vol: 4, rft: 5, otd: 6, eff: 7, rework: 8 }[c.key];
          sp.getRange(5 + i, idx).setFontColor(c.ok ? '#0a7a0a' : '#b42323');
        });
      });
    }
    let row = 5 + rows.length + 1;

    // Bloc 2 : matrice membre × gamme
    cpSection_(sp, row, '2. Statistiques par personne et par gamme (volume · RFT %) sur la fenêtre', Math.max(12, d.gammes.length + 1));
    row++;
    const head = ['Membre'].concat(d.gammes);
    sp.getRange(row, 1, 1, head.length).setValues([head]);
    cpHeader_(sp.getRange(row, 1, 1, head.length));
    row++;
    const mrows = d.members.map(m => [m.name].concat(d.gammes.map(g => m.perGamme[g] ? m.perGamme[g].n + ' · ' + m.perGamme[g].rftPct + '%' : '—')));
    if (mrows.length && d.gammes.length) { sp.getRange(row, 1, mrows.length, head.length).setValues(mrows); sp.getRange(row, 2, mrows.length, d.gammes.length).setHorizontalAlignment('center'); }
    row += mrows.length + 1;

    // Bloc 3 : cartes de contrôle
    cpSection_(sp, row, '3. Cartes de contrôle par gamme — ratio temps réel ÷ target (%), ' + CP_CC_WEEKS + ' dernières semaines', 12);
    row++;
    sp.getRange(row, 1, 1, 9).setValues([['Gamme', 'Points', 'Moyenne %', 'Écart-type', 'Limite basse', 'Limite haute', 'Hors limites', 'Statut', 'Anomalies']]);
    cpHeader_(sp.getRange(row, 1, 1, 9));
    row++;
    const crow = d.control.map(c => [c.gamme, c.n, c.n >= d.params.minPts ? c.mean : '—', c.n >= d.params.minPts ? c.sd : '—', c.n >= d.params.minPts ? c.lcl : '—', c.n >= d.params.minPts ? c.ucl : '—',
      c.outCount, c.status === 'ok' ? 'Sous contrôle' : (c.status === 'hors' ? 'Hors contrôle' : 'Données insuffisantes'), c.reasons.join(' ; ') || '—']);
    if (crow.length) {
      sp.getRange(row, 1, crow.length, 9).setValues(crow);
      sp.getRange(row, 2, crow.length, 7).setHorizontalAlignment('center');
      d.control.forEach((c, i) => sp.getRange(row + i, 8).setFontWeight('bold').setBackground(c.status === 'ok' ? '#e6f6e6' : (c.status === 'hors' ? '#fde7e7' : '#f1f5f9')));
    }
    row += crow.length + 1;

    // Bloc 4 : ANOVA entre types de gammes
    const an = d.anova;
    cpSection_(sp, row, '4. Analyse ANOVA sur les types (groupes) de gammes — ratio temps réel ÷ target (%)', 12);
    row++;
    sp.getRange(row, 1).setValue(an.conclusion).setFontWeight('bold').setFontColor(an.ok && an.significant ? '#b42323' : '#334155');
    row++;
    if (an.ok) {
      sp.getRange(row, 1, 1, 6).setValues([['Groupes (k)', 'Observations (N)', 'F', 'p-value', 'η² (effet)', 'Moyenne générale %']]);
      cpHeader_(sp.getRange(row, 1, 1, 6)); row++;
      sp.getRange(row, 1, 1, 6).setValues([[an.k, an.N, an.F, an.p < 0.001 ? '< 0,001' : cpRound_(an.p, 3), an.eta2, an.grand]]).setHorizontalAlignment('center'); row += 2;
      sp.getRange(row, 1, 1, 6).setValues([['Type de gamme', 'Points', 'Moyenne %', 'Écart-type', 'IC 95% bas', 'IC 95% haut']]);
      cpHeader_(sp.getRange(row, 1, 1, 6)); row++;
      const grows = an.groups.map(g => [g.gamme, g.n, g.mean, g.sd, g.lo, g.hi]);
      sp.getRange(row, 1, grows.length, 6).setValues(grows); sp.getRange(row, 2, grows.length, 5).setHorizontalAlignment('center'); row += grows.length + 1;
      sp.getRange(row, 1, 1, 4).setValues([['Paire', 'Écart de moyennes (pts)', 'p ajustée (Bonferroni)', 'Verdict']]);
      cpHeader_(sp.getRange(row, 1, 1, 4)); row++;
      const prow = an.pairs.map(x => [x.a + ' vs ' + x.b, x.diff, x.p < 0.001 ? '< 0,001' : cpRound_(x.p, 3), x.significant ? 'Différent' : 'Comparable']);
      sp.getRange(row, 1, prow.length, 4).setValues(prow); sp.getRange(row, 2, prow.length, 3).setHorizontalAlignment('center');
      an.pairs.forEach((x, i) => sp.getRange(row + i, 4).setFontWeight('bold').setBackground(x.significant ? '#fde7e7' : '#e6f6e6'));
    }
    sp.setColumnWidths(1, 1, 200); sp.setColumnWidths(2, 10, 120); sp.setColumnWidths(9, 1, 170); sp.setColumnWidths(11, 1, 170); sp.setColumnWidths(12, 1, 380);
    sp.setFrozenRows(2);

    // Mise à jour de « Niveaux Équipe »
    const st = ss.getSheetByName(CP_SHEETS.team);
    if (st) {
      const today = new Date();
      d.members.forEach(m => {
        if (!m.row) return;
        st.getRange(m.row, 7, 1, 2).setValues([[m.proposal, m.decision]]);
        st.getRange(m.row, 10).setValue(today);
        st.getRange(m.row, 7).setBackground(m.proposalType === 'up' ? '#e6f6e6' : (m.proposalType === 'down' ? '#fde7e7' : null));
      });
    }
    return { members: d.members.length, pending: d.pending, demo: d.demo };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------
//  Notification & décision (web app + menu)
// ---------------------------------------------------------------------
function notifyResponsible() {
  const d = cpCompute_();
  const list = d.members.filter(m => m.proposalType !== 'none' && m.decision === 'En attente');
  if (!list.length) return { sent: false, count: 0, message: 'Aucune proposition en attente de validation.' };
  const lines = list.map(m => '• ' + m.name + ' (N' + m.level + ') : ' + m.proposal + ' — ' + m.gap +
    ' [vol ' + m.stats.n + ', RFT ' + m.stats.rftPct + '%, OTD ' + m.stats.otdPct + '%, eff. ' + m.stats.effPct + '%]');
  const body = 'Bonjour ' + d.params.resp + ',\n\n' + list.length + ' proposition(s) de changement de niveau attendent votre décision :\n\n' + lines.join('\n') +
    '\n\nValidez ou refusez dans l\'onglet « Niveaux Équipe » (colonne Décision) ou dans la web app (Compétences ▸ Équipe & Niveaux).';
  if (!d.params.notify) {
    return { sent: false, count: list.length, message: 'Notifications désactivées (Seuils & Règles ▸ « Notifications email actives » = NON). Simulation : ' + list.length + ' alerte(s) pour ' + d.params.resp + ' (' + d.params.email + ').' };
  }
  MailApp.sendEmail(d.params.email, 'Compétences — ' + list.length + ' changement(s) de niveau à valider', body);
  return { sent: true, count: list.length, message: 'Email envoyé à ' + d.params.resp + ' (' + d.params.email + ') : ' + list.length + ' proposition(s).' };
}

function decideLevelChange(name, decision, comment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = ss.getSheetByName(CP_SHEETS.team);
  if (!st) throw new Error('Onglet « ' + CP_SHEETS.team + ' » introuvable : lancez Compétences ▸ Initialiser les 3 onglets.');
  if (['Validé', 'Refusé', 'En attente'].indexOf(decision) === -1) throw new Error('Décision invalide.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = cpCompute_();
    const m = d.members.filter(x => x.name === name)[0];
    if (!m || !m.row) throw new Error('Membre introuvable : ' + name);
    const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy');
    const note = (comment ? comment + ' — ' : '') + decision + ' le ' + stamp + ' par ' + d.params.resp;
    if (decision === 'Validé') {
      if (m.proposalType === 'none') throw new Error('Aucune proposition à valider pour ' + name + '.');
      st.getRange(m.row, 3).setValue(m.proposalTo);
      st.getRange(m.row, 6).setValue(new Date());
      st.getRange(m.row, 7, 1, 2).setValues([['—', '—']]);
      st.getRange(m.row, 9).setValue('N' + m.level + ' → N' + m.proposalTo + ' · ' + note);
    } else {
      st.getRange(m.row, 7, 1, 2).setValues([[m.proposal, decision]]);
      st.getRange(m.row, 9).setValue(note);
    }
  } finally {
    lock.releaseLock();
  }
  return cpCompute_();
}
