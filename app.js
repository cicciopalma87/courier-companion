'use strict';
/* Courier Companion – app corriere (PWA). Dati solo sul telefono, database in file .xlsx */
const VER = '1.0';
const $ = s => document.querySelector(s);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad = (n, l = 2) => String(n).padStart(l, '0');
const todayStr = () => { const d = new Date(); return pad(d.getDate()) + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getFullYear() % 100); };
const nowIso = () => new Date().toISOString();

/* ---------- memoria locale ---------- */
const K = { db: 'sda.db', day: 'sda.day', snaps: 'sda.snaps', wa: 'sda.wa' };
function lsGet(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }
function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { toast('Memoria piena: fai un backup'); return false; } }

/* ---------- priorità di default (foglio del collega) ---------- */
const E = (n, zona, com, loc, testo, keys) => ({ n, zona, com, loc: loc || '', testo, keys });
const DEFAULT_PRIO = [
  E(1, 'Arpaise', 'arpaise', '', 'Mignolli, zona Pasquarielli, area Miranda, via delle Madonnine', ['mignolli', 'pasquarielli', 'miranda', 'madonnine']),
  E(2, 'Arpaise', 'arpaise', '', 'Zona Terranova, casino Papa', ['terranova', 'casino papa']),
  E(3, 'Arpaise', 'arpaise', '', 'Via Urbanelli, via delle Noci, via Russi, via Crocevia', ['urbanelli', 'noci', 'russi', 'crocevia']),
  E(4, 'Arpaise', 'arpaise', '', 'Via Toti, via Casalpreti', ['toti', 'casalpreti']),
  E(5, 'Arpaise', 'arpaise', '', 'Corso Emilio Capone, via G. Capone', ['emilio capone', 'capone']),
  E(6, 'San Giovanni di Ceppaloni', 'ceppaloni', 'sg', 'Via prov. San Giovanni dal 15 al 25 lato sx, via prov 9', ['provinciale san giovanni@9-25']),
  E(7, 'San Giovanni di Ceppaloni', 'ceppaloni', 'sg', 'Via Paradisi, via Chiesa', ['paradisi', 'chiesa']),
  E(8, 'San Giovanni di Ceppaloni', 'ceppaloni', 'sg', 'Corso Umberto, via Tre Santi', ['umberto', 'tre santi']),
  E(9, 'San Giovanni di Ceppaloni', 'ceppaloni', 'sg', 'Via Roma, via Mernoni, piazza Martiri, via 28 Ottobre', ['roma', 'mernoni', 'martiri', '28 ottobre']),
  E(10, 'San Giovanni di Ceppaloni', 'ceppaloni', 'sg', 'Via Venagli, via Olivella', ['venagli', 'venaglie', 'olivella']),
  E(11, 'San Giovanni di Ceppaloni', 'ceppaloni', 'sg', 'Via Provinciale da 145 al 77 tornando indietro a San Giovanni', ['provinciale san giovanni@77-145']),
  E(12, 'San Giovanni di Ceppaloni', 'ceppaloni', 'sg', 'Santa Croce (via Paduli, via Brecciale)', ['santa croce', 'paduli', 'brecciale']),
  E(13, 'San Giovanni di Ceppaloni', 'ceppaloni', 'sg', 'Barba (via Portovico, via prov Barba)', ['barba', 'portovico', 'provinciale barba']),
  E(14, 'Ceppaloni paese', 'ceppaloni', 'cp', 'Via Bosco, via Croce', ['bosco', 'croce']),
  E(15, 'Ceppaloni paese', 'ceppaloni', 'cp', 'Piazza Carmine Rossi (bar Kaffa), via Cortile, via Piazzetella', ['carmine rossi', 'cortile', 'piazzetella']),
  E(16, 'Ceppaloni paese', 'ceppaloni', 'cp', 'Via Stazione', ['stazione']),
  E(17, 'Ceppaloni paese', 'ceppaloni', 'cp', 'Via Donisi, via Cretazzo', ['donisi', 'cretazzo']),
  E(18, 'Ceppaloni paese', 'ceppaloni', 'cp', 'Via Porcari, via Manni', ['porcari', 'manni']),
  E(19, 'Ceppaloni paese', 'ceppaloni', 'bt', 'Via prov. Beltiglio', ['provinciale beltiglio']),
  E(20, 'Ceppaloni paese', 'ceppaloni', 'bt', 'Via Testi, via Cuonti, via Zerelli', ['testi', 'cuonti', 'zerelli']),
  E(21, 'Ceppaloni paese', 'ceppaloni', 'bt', 'Via Maielli, via Linguecane, via Masseriola', ['maielli', 'linguecane', 'masseriola']),
  E(22, 'Ceppaloni paese', 'ceppaloni', 'cp', 'Via Roma', ['roma']),
  E(23, 'Ceppaloni paese', 'ceppaloni', 'bt', 'Via Catalani, via Trocchia', ['catalani', 'trocchia']),
  E(24, 'Ceppaloni paese', 'ceppaloni', 'bt', 'Via Streppe, via Pizzirussi, via Marotti, via Masserie Alfieri', ['streppe', 'pizzirussi', 'marotti', 'masserie alfieri']),
  E(25, 'Ceppaloni paese', 'ceppaloni', 'bt', 'Prov. Beltiglio dal 161 al 171', ['provinciale beltiglio@161-171']),
  E(26, 'San Leucio del Sannio', 'san leucio', '', 'Via Mazzini, via Amicoli, via prov (pari fino al 68, dispari fino al 17), poi tornare indietro', ['mazzini', 'amicoli', 'provinciale']),
  E(27, 'San Leucio del Sannio', 'san leucio', '', 'Via Garibaldi, traversa Pepicelli, via Cardoni', ['garibaldi', 'pepicelli', 'cardoni']),
  E(28, 'San Leucio del Sannio', 'san leucio', '', 'Via Vittorio Emanuele fino al 38, (torna indietro) Carmine Zerella (posta)', ['vittorio emanuele@1-38', 'zerella']),
  E(29, 'San Leucio del Sannio', 'san leucio', '', 'Via De Longis, via Rizzi', ['de longis', 'rizzi']),
  E(30, 'San Leucio del Sannio', 'san leucio', '', 'Via Roma', ['roma']),
  E(31, 'San Leucio del Sannio', 'san leucio', '', 'Via Vittorio Emanuele', ['vittorio emanuele']),
  E(32, 'San Leucio del Sannio', 'san leucio', '', 'Via Angelo Zamparelli, via Piano di Zona, via Tuoppo', ['zamparelli', 'piano di zona', 'tuoppo']),
  E(33, 'San Leucio del Sannio', 'san leucio', '', 'Via prov 164, via Paparano, via Masserie', ['provinciale@164-164', 'paparano', 'masserie', 'casale maccabei']),
  E(34, 'San Leucio del Sannio', 'san leucio', '', 'Via Santa Maria 80 (carabinieri), via San Marcello', ['santa maria@80-80', 'san marcello']),
  E(35, 'San Leucio del Sannio', 'san leucio', '', 'Via Maccabei, Taverna Rotola (zona industriale), Fondovalle Sabato', ['maccabei', 'taverna rotola', 'rotola', '*fondovalle sabato']),
  E(36, 'San Leucio del Sannio', 'san leucio', '', 'Via Sauco', ['sauco']),
  E(37, 'San Leucio del Sannio', 'san leucio', '', 'Via Santa Maria numeri alti', ['santa maria@60-9999']),
  E(38, 'San Leucio del Sannio', 'san leucio', '', 'Via Vigne, via Merici', ['vigne', 'merici']),
  E(39, 'San Leucio del Sannio', 'san leucio', '', 'Via Santa Maria numeri bassi, via Abate Andrea, via Pecci', ['santa maria', 'abate andrea', 'pecci']),
  E(40, 'San Leucio del Sannio', 'san leucio', '', 'Contrade: via Vardaro, Alvanelle, Boschi, Salato, Macchiarotonda', ['vardaro', 'alvanelle', 'boschi', 'salato', 'macchiarotonda']),
  E(41, 'San Leucio del Sannio', 'san leucio', '', 'Via Pitangelo, via Ferrovia, via Varricchio, via Feleppi', ['pitangelo', 'ferrovia', 'varricchio', 'feleppi']),
  E(42, 'San Leucio del Sannio', 'san leucio', '', 'Via Ariarella, via Festola', ['ariarella', 'festola']),
  E(43, 'Apollosa', 'apollosa', '', 'Appia Apollosa', ['appia']),
  E(44, 'Apollosa', 'apollosa', '', 'Epitaffio', ['epitaffio']),
  E(45, 'Apollosa', 'apollosa', '', 'Appia antica, via Lapillusia', ['appia antica', 'lapillusia']),
  E(46, 'Apollosa', 'apollosa', '', 'Via Cancellonica', ['cancellonica']),
  E(47, 'Apollosa', 'apollosa', '', 'Via Vaccari, via Marini', ['vaccari', 'marini']),
  E(48, 'Apollosa', 'apollosa', '', 'Via Savoiari, via Francesca di Sotto', ['savoiari', 'francesca di sotto']),
  E(49, 'Apollosa', 'apollosa', '', 'Via Guadagni, via Aria di Caruso, via Stefanucci', ['guadagni', 'aria di caruso', 'stefanucci']),
  E(50, 'Apollosa', 'apollosa', '', 'Via Roma fino al 33, via Tuoro, via Carpine', ['roma@1-33', 'tuoro', 'carpine']),
  E(51, 'Apollosa', 'apollosa', '', 'Via Roma, piazza Saponaro', ['roma', 'saponaro']),
  E(52, 'Apollosa', 'apollosa', '', 'Via Petrera fino all\'azienda Suimi, via Fontana Spina', ['petrera', 'fontana spina']),
  E(53, 'Apollosa', 'apollosa', '', 'Mazzoni Curti, via Verdini', ['mazzoni curti', 'verdini']),
  E(54, 'Apollosa', 'apollosa', '', 'Via Pacilli, via Mazzoni, via Meoli', ['pacilli', 'mazzoni', 'meoli']),
  E(55, 'Apollosa', 'apollosa', '', 'Via Triduonno, Monte Mauro, via Fiego, via Patriello', ['triduonno', 'monte mauro', 'fiego', 'patriello']),
  E(56, 'Apollosa', 'apollosa', '', 'San Giovanni di Apollosa (provinciale, via Tudini, via Pirozzi, via Cesinali, via Case Grasso, via Ariomanno)', ['provinciale', 'provinciale san giovanni', 'tudini', 'pirozzi', 'cesinali', 'case grasso', 'ariomanno']),
  E(57, 'Ceppaloni zona industriale e Appia', 'ceppaloni', '', 'Via Petrara di Ceppaloni, via Taverna, via Cortoffo', ['petrara', 'taverna', 'cortoffo']),
  E(58, 'Ceppaloni zona industriale e Appia', 'ceppaloni', '', 'Appia di Ceppaloni (Pinauto, Duedi, Dematech, Montepol, carrozzerie)', ['appia']),
  E(59, 'Ceppaloni zona industriale e Appia', 'ceppaloni', '', 'Via Tre Santi di Tufara', ['tre santi', 'tufara']),
  E(60, 'Ceppaloni zona industriale e Appia', 'roccabascerana', '', 'Via Nazionale Appia (Roccabascerana) fino al bar Rendez Vous, via Fiego Basso', ['nazionale appia', 'appia', 'fiego basso']),
  E(61, 'Ceppaloni zona industriale e Appia', 'roccabascerana', '', 'Strada statale 7 di Montesarchio-San Martino fino al 36', ['statale 7@1-36', 'montesarchio']),
  E(62, 'Ceppaloni zona industriale e Appia', 'roccabascerana', '', 'Via Caraccioli (zona industriale di Rocca), via Olivella, via Boscovico fino al 18', ['caraccioli', 'olivella', 'boscovico@1-18']),
  E(63, 'Ceppaloni zona industriale e Appia', 'ceppaloni', '', 'Case sparse', []),
  E(64, 'Roccabascerana', 'roccabascerana', '', 'Zolli (via Vittorio Veneto)', ['zolli', 'vittorio veneto']),
  E(65, 'Roccabascerana', 'roccabascerana', '', 'Squillani, via Cotugno, via Girolamo del Balzo, via Cantone', ['squillani', 'cotugno', 'girolamo del balzo', 'cantone']),
  E(66, 'Roccabascerana', 'roccabascerana', '', 'Via Severini, via Guerrere, via Vernilli', ['severini', 'guerrere', 'vernilli']),
  E(67, 'Cassano (Roccabascerana)', 'roccabascerana', '', 'Via Giordano Bruno dal 231 al 225, via Cafari', ['giordano bruno@225-231', 'cafari']),
  E(68, 'Cassano (Roccabascerana)', 'roccabascerana', '', 'Via Giordano Bruno, via Arnaldo da Brescia', ['giordano bruno', 'arnaldo da brescia']),
  E(69, 'Cassano (Roccabascerana)', 'roccabascerana', '', 'Via Caselle, via Principi', ['caselle', 'principi']),
  E(70, 'Cassano (Roccabascerana)', 'roccabascerana', '', 'Via Bottilli', ['bottilli']),
  E(71, 'Tuoro (Roccabascerana)', 'roccabascerana', '', 'Via Termiti, via Piantedosi, piazza Covino, via Masiello', ['termiti', 'piantedosi', 'covino', 'masiello']),
  E(72, 'Tuoro (Roccabascerana)', 'roccabascerana', '', 'Via Fontana, via Campi', ['fontana', 'campi']),
  E(73, 'Rocca paese', 'roccabascerana', '', 'Piazza Imbriani, parco San Giorgio, via Marmora', ['imbriani', 'parco san giorgio', 'marmora']),
  E(74, 'Rocca paese', 'roccabascerana', '', 'Via Roma, via Orsini', ['roma', 'orsini']),
  E(75, 'Rocca paese', 'roccabascerana', '', 'Via Miranda fino al 81, via Alvaneta, via Santoro', ['miranda@1-81', 'alvaneta', 'santoro']),
  E(76, 'Rocca paese', 'roccabascerana', '', 'Via Miranda fino al 101', ['miranda@82-101']),
  E(77, 'Pannarano', 'pannarano', '', 'Via Irpinia fino al 27', ['irpinia@1-27']),
  E(78, 'Pannarano', 'pannarano', '', 'Via Piano, via Papa Giovanni 23, via Rampe Fontana 7-10', ['piano', 'papa giovanni@23-23', 'rampe fontana@7-10']),
  E(79, 'Pannarano', 'pannarano', '', 'Via Municipio fino al 46, via Torre, via Antonio Dalessio 1-2', ['municipio@1-46', 'torre', 'dalessio@1-2']),
  E(80, 'Pannarano', 'pannarano', '', 'Piazza dei Martiri, via Indipendenza fino al 89, via Vico Torre', ['martiri', 'indipendenza@1-89', 'vico torre']),
  E(81, 'Pannarano', 'pannarano', '', 'Via Valle, via Dalessio, via Campetiello, via Vigne Vecchie fino al 10', ['valle', 'dalessio', 'campetiello', 'vigne vecchie@1-10']),
  E(82, 'Pannarano', 'pannarano', '', 'Via Aqualta, via Santa Maria, piazza Eugenio Dalessio', ['aqualta', 'santa maria', 'eugenio dalessio']),
  E(83, 'Pannarano', 'pannarano', '', 'Via Selvetelle Pezze, via Franchi Pezze fino al 15 (poi tornare indietro)', ['selvetelle pezze', 'franchi pezze@1-15', 'pezze']),
  E(84, 'Pannarano', 'pannarano', '', 'Via Selvetelle, via Ponte, via Indipendenza numeri alti fino al 157, largo Taverna', ['selvetelle', 'ponte', 'indipendenza@90-157', 'taverna']),
  E(85, 'Pannarano', 'pannarano', '', 'Via Pietrastornina', ['pietrastornina']),
  E(86, 'Pannarano', 'pannarano', '', 'Via Sorbo San Pietro, via Cerreto, via Rampe Fontana 32, via Giacchini', ['sorbo san pietro', 'cerreto', 'rampe fontana@32-32', 'giacchini']),
  E(87, 'Pannarano', 'pannarano', '', 'Via Irpinia dal 29 al 80, via San Nicola, via Vigne Vecchie 13, via Conti', ['irpinia@29-80', 'san nicola', 'vigne vecchie@13-13', 'conti'])
];

/* ---------- stato ---------- */
let DB = (() => { const d = lsGet(K.db, null); return d && d.addr ? d : { addr: {}, prio: JSON.parse(JSON.stringify(DEFAULT_PRIO)) }; })();
let DAY = lsGet(K.day, null) || { date: todayStr(), parcels: [], nextId: 1, lastBackup: 0, doneSince: 0, dirty: 0, snooze: 0 };
let SNAPS = lsGet(K.snaps, []);
let WA = lsGet(K.wa, 'Buongiorno, sono il corriere SDA. Ho lasciato il suo pacco in un luogo sicuro, le invio la foto.');
function saveDB() { lsSet(K.db, DB); }
function saveDay() { lsSet(K.day, DAY); }
function saveSnaps() { lsSet(K.snaps, SNAPS); }

/* ---------- normalizzazione indirizzi ---------- */
const TYPE_WORDS = new Set(['via', 'viale', 'vico', 'vicolo', 'corso', 'piazza', 'piazzale', 'largo', 'strada', 'str', 'contrada', 'c', 'da', 'loc', 'localita', 'traversa', 'trav', 'cda']);
function norm(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim().replace(/\s+/g, ' ');
}
function normStreet(s) {
  let t = norm(s).split(' ').map(w => ({ prov: 'provinciale', sp: 'provinciale', vitt: 'vittorio', eman: 'emanuele', s: 'san' }[w] || w));
  while (t.length > 1 && TYPE_WORDS.has(t[0])) t.shift();
  return t.join(' ');
}
const KNOWN_COM = ['ceppaloni', 'apollosa', 'arpaise', 'san leucio', 'roccabascerana', 'pannarano', 'bucciano', 'sant angelo a cupolo', 'montesarchio'];
function normCom(c) {
  const n = norm(c).replace(/ campania$/, ''); const m = n.match(/^.+? di (.+)$/);
  if (m && KNOWN_COM.some(k => m[1].startsWith(k))) return m[1];
  return n;
}
function locOf(c, cap) { const n = norm(c); if (/^beltiglio/.test(n)) return 'bt'; if (/^san giovanni di ceppaloni/.test(n)) return 'sg'; if (/^ceppaloni/.test(n) && String(cap || '').trim() === '82010') return 'sgbt'; return ''; }
function baseComune(c) { const n = normCom(c); return n.replace(/\b\w/g, x => x.toUpperCase()); }
const sqStreet = s => normStreet(s).replace(/ /g, '');
function akey(via, civ, com) { return sqStreet(via) + '|' + norm(civ) + '|' + normCom(com); }
const akeyP = p => akey(p.via, p.civico, p.comune);
/* le chiavi del database ignorano gli spazi nel nome della via (abate andrea = abateandrea): riallineo quelle vecchie */
function rekeyDB() {
  const out = {}; let ch = false;
  for (const [k, a] of Object.entries(DB.addr)) {
    const nk = akey(a.via, a.civico, a.comune); if (nk !== k) ch = true;
    const c = out[nk];
    if (!c) { out[nk] = a; continue; }
    const [n, o] = (a.mod || '') >= (c.mod || '') ? [a, c] : [c, a];
    if (n.lat == null && o.lat != null) { n.lat = o.lat; n.lng = o.lng; }
    if (!n.note && o.note) n.note = o.note; else if (o.note && n.note !== o.note && !n.note.includes(o.note)) n.note += ' · ' + o.note;
    if (!n.prio && o.prio) n.prio = o.prio;
    const nm = new Set([].concat(n.nomi || [], o.nomi || [])); if (nm.size) n.nomi = [...nm];
    out[nk] = n;
  }
  if (ch || Object.keys(out).length !== Object.keys(DB.addr).length) { DB.addr = out; try { saveDB(); } catch (e) { } }
}
try { rekeyDB(); } catch (e) { }

/* ---------- priorità: abbinamento ---------- */
function matchPrio0(p) {
  const a = DB.addr[akeyP(p)];
  if (a && a.prio) return { n: +a.prio, amb: false };
  const st = normStreet(p.via), civ = parseInt(p.civico, 10), com = normCom(p.comune), loc = locOf(p.comune, p.cap);
  if (!st) return { n: 0, amb: false };
  const padded = ' ' + st + ' ';
  let best = [], bs = 0;
  for (const e of DB.prio) {
    const comOk = !e.com || com.startsWith(e.com);
    if (comOk && loc && e.loc && e.loc !== '' && (loc === 'sgbt' ? (e.loc !== 'sg' && e.loc !== 'bt') : e.loc !== loc)) continue;
    for (const raw of (e.keys || [])) {
      const star = String(raw).startsWith('*'); if (!comOk && !star) continue;
      const [kk, r] = String(raw).replace(/^\*/, '').split('@'); const k = normStreet(kk);
      if (!k) continue;
      if (!padded.includes(' ' + k + ' ') && !(k.replace(/ /g, '').length >= 7 && sqStreet(st).includes(k.replace(/ /g, '')))) continue;
      if (r) { const [lo, hi] = r.split('-').map(Number); if (!(civ >= lo && civ <= hi)) continue; }
      const score = k.length + (r ? 100 : 0);
      if (score > bs) { bs = score; best = [e]; } else if (score === bs && !best.includes(e)) best.push(e);
    }
  }
  if (!best.length) return { n: 0, amb: false };
  if (best.length > 1 && (!loc || loc === 'sgbt')) { const pref = loc === 'sgbt' ? best.filter(e => e.loc === 'sg' || e.loc === 'bt') : best.filter(e => e.loc === 'cp'); if (pref.length) best = pref; }
  best.sort((x, y) => x.n - y.n);
  return { n: best[0].n, amb: best.length > 1 };
}


/* se la via non è nel foglio: stessa via degli altri pacchi, altrimenti la priorità mediana del paese (segnata con ?) */
function estPrio(p) {
  const st = sqStreet(p.via), com = normCom(p.comune);
  const same = DAY.parcels.filter(q => q !== p && q.prio && !q.amb && normCom(q.comune) === com && st && sqStreet(q.via) === st).map(q => q.prio);
  if (same.length) return Math.min(...same);
  const ns = DB.prio.filter(e => e.com && com && com.startsWith(e.com)).map(e => e.n).sort((a, b) => a - b);
  return ns.length ? ns[Math.floor(ns.length / 2)] : 0;
}
function matchPrio(p) {
  const m = matchPrio0(p); if (m.n) return m;
  const e = estPrio(p); return e ? { n: e, amb: true } : m;
}

/* ---------- utilità interfaccia ---------- */
let toastT;
function toast(m, ms = 2600) { const t = $('#toast'); t.textContent = m; t.style.display = 'block'; clearTimeout(toastT); toastT = setTimeout(() => t.style.display = 'none', ms); }
function busy(msg) { const b = $('#busy'); if (msg) { b.innerHTML = '<div>' + msg + '</div>'; b.style.display = 'flex'; } else b.style.display = 'none'; }
let sheetOnClose = null;
function sheet(html, onClose) {
  $('#sheet').innerHTML = html; $('#sheet').style.display = 'block'; $('#sheetbg').style.display = 'block';
  sheetOnClose = onClose || null; $('#sheet').scrollTop = 0;
}
function closeSheet() {
  $('#sheet').style.display = 'none'; $('#sheetbg').style.display = 'none';
  const f = sheetOnClose; sheetOnClose = null; if (f) f();
}
$('#sheetbg').addEventListener('click', closeSheet);

function download(name, blob) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}
function xlsxBlob(wb) { const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' }); return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); }

/* ---------- database: righe <-> xlsx ---------- */
function addrRows() {
  return Object.entries(DB.addr).map(([k, a]) => ({ chiave: k, via: a.via, civico: a.civico, cap: a.cap || '', comune: a.comune, lat: a.lat, lng: a.lng, note: a.note || '', nomi: (a.nomi || []).join('; '), priorita: a.prio || '', modificato: a.mod || '' }));
}
function prioRows() { return DB.prio.map(e => ({ n: e.n, zona: e.zona, comune: e.com, sottozona: e.loc, descrizione: e.testo, chiavi: (e.keys || []).join('; ') })); }
function dbWorkbook() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(addrRows(), { header: ['chiave', 'via', 'civico', 'cap', 'comune', 'lat', 'lng', 'note', 'nomi', 'priorita', 'modificato'] }), 'Indirizzi');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(prioRows(), { header: ['n', 'zona', 'comune', 'sottozona', 'descrizione', 'chiavi'] }), 'Priorità');
  return wb;
}
function mergeAddrRows(rows) {
  let n = 0;
  for (const r of rows) {
    const via = String(r.via || '').trim(); if (!via) continue;
    const lat = parseFloat(String(r.lat).replace(',', '.')), lng = parseFloat(String(r.lng).replace(',', '.'));
    const key = akey(via, r.civico, r.comune);
    const inc = { via, civico: String(r.civico == null ? '' : r.civico), cap: String(r.cap || ''), comune: String(r.comune || ''), lat: isFinite(lat) ? lat : null, lng: isFinite(lng) ? lng : null, note: String(r.note || ''), prio: r.priorita ? +r.priorita : 0, nomi: String(r.nomi || '').split(';').map(x => x.trim()).filter(Boolean), mod: String(r.modificato || '') };
    const cur = DB.addr[key];
    if (!cur || (inc.mod || '') >= (cur.mod || '')) { DB.addr[key] = inc; n++; }
  }
  return n;
}
function mergePrioRows(rows) {
  const out = [];
  for (const r of rows) {
    const n = parseInt(r.n, 10); if (!n) continue;
    out.push({ n, zona: String(r.zona || ''), com: String(r.comune || ''), loc: String(r.sottozona || ''), testo: String(r.descrizione || ''), keys: String(r.chiavi || '').split(';').map(s => s.trim()).filter(Boolean) });
  }
  if (out.length) { out.sort((a, b) => a.n - b.n); DB.prio = out; }
  return out.length;
}

/* ---------- snapshot (copie progressive pin/note) ---------- */
function touch() { DAY.dirty = (DAY.dirty || 0) + 1; if (!DAY.lastBackup) DAY.lastBackup = Date.now(); saveDay(); }
function snapshot() {
  const t = todayStr(); const nn = SNAPS.filter(s => s.name.startsWith('data-' + t)).length;
  const today = nowIso().slice(0, 10);
  const mod = Object.entries(DB.addr).filter(([k, a]) => (a.mod || '').slice(0, 10) === today);
  SNAPS.push({ name: 'data-' + t + '-' + pad(nn, 3) + '.temp.xlsx', t: nowIso(), addr: mod.map(([k, a]) => a) });
  if (SNAPS.length > 40) SNAPS.splice(0, SNAPS.length - 40);
  saveSnaps(); touch();
}
/* archivio: ricordo il destinatario insieme all'indirizzo (solo sul tuo telefono e nel file che esporti) */
function archiveParcel(p) {
  if (!p || !p.via) return;
  const key = akeyP(p); const cur = DB.addr[key] || { via: p.via, civico: p.civico, cap: p.cap, comune: p.comune, lat: p.lat != null ? p.lat : null, lng: p.lng != null ? p.lng : null, note: '', prio: 0 };
  if (cur.lat == null && p.lat != null && (p.geo === 'pin' || p.geo === 'db')) { cur.lat = p.lat; cur.lng = p.lng; }
  const nm = String(p.name || '').trim();
  if (nm) { const l = cur.nomi || []; if (!l.some(x => norm(x) === norm(nm))) { l.push(nm); cur.nomi = l.slice(-6); cur.mod = nowIso(); } }
  DB.addr[key] = cur; saveDB();
}
function saveAddr(p, patch) {
  const key = akeyP(p); const cur = DB.addr[key] || { via: p.via, civico: p.civico, cap: p.cap, comune: p.comune, lat: null, lng: null, note: '', prio: 0 };
  Object.assign(cur, patch, { mod: nowIso() }); DB.addr[key] = cur; saveDB(); snapshot(); return cur;
}
function backupNow(silent) {
  if (!SNAPS.length) snapshot();
  const wb = dbWorkbook(); const name = SNAPS[SNAPS.length - 1].name;
  download(name, xlsxBlob(wb));
  DAY.lastBackup = Date.now(); DAY.doneSince = 0; DAY.dirty = 0; saveDay();
  if (!silent) toast('Backup scaricato: ' + name);
}
function checkRemind() {
  const b = $('#banner');
  const due = (DAY.dirty > 0) && (DAY.doneSince >= 10 || (DAY.lastBackup && Date.now() - DAY.lastBackup > 3600e3)) && Date.now() > (DAY.snooze || 0);
  b.style.display = due ? 'block' : 'none';
}
$('#bkYes').onclick = () => { backupNow(); checkRemind(); };
$('#bkLater').onclick = () => { DAY.snooze = Date.now() + 15 * 60e3; saveDay(); checkRemind(); };
setInterval(checkRemind, 30000);

/* ---------- pacchi ---------- */
const STATI = { consegna: { done: 'consegnato', no: 'non_consegnato' }, ritiro: { done: 'ritirato', no: 'non_ritirato' } };
function stClass(p) {
  if (p.stato === 'consegnato') return 'st-cd'; if (p.stato === 'non_consegnato') return 'st-cn';
  if (p.stato === 'ritirato') return 'st-rd'; if (p.stato === 'non_ritirato') return 'st-rn';
  return p.tipo === 'ritiro' ? 'st-r' : 'st-c';
}
function stLabel(s) { return { consegnato: 'Consegnato', non_consegnato: 'Non consegnato / assente', ritirato: 'Ritirato', non_ritirato: 'Non ritirato', da_fare: 'Da fare' }[s] || s; }
function newParcel(d) {
  return Object.assign({ id: DAY.nextId++, track: '', name: '', via: '', civico: '', cap: '', comune: '', prov: '', tel: '', cod: '', tipo: 'consegna', lat: null, lng: null, geo: '', prio: 0, amb: false, stato: 'da_fare', ord: 0 }, d);
}
function fmtAddr(p) { return (p.via + ' ' + (p.civico || '')).trim() + (p.comune ? ', ' + p.comune : ''); }
function counters() {
  const c = DAY.parcels.filter(p => p.tipo === 'consegna'), r = DAY.parcels.filter(p => p.tipo === 'ritiro');
  const cd = c.filter(p => p.stato !== 'da_fare').length, rd = r.filter(p => p.stato !== 'da_fare').length;
  $('#cnt').innerHTML = 'Consegne ' + cd + '/' + c.length + '<br>Ritiri ' + rd + '/' + r.length;
}

/* ---------- ordine di consegna (numero progressivo mostrato sui pin) ---------- */
function cmpAddr(a, b) { return normStreet(a.via).localeCompare(normStreet(b.via)) || ((parseInt(a.civico) || 0) - (parseInt(b.civico) || 0)); }
function ensureOrder() {
  const ps = DAY.parcels; if (!ps.length) return;
  const have = ps.filter(p => p.ord > 0).sort((a, b) => a.ord - b.ord);
  const fresh = ps.filter(p => !(p.ord > 0)).sort((a, b) => ((a.prio || 999) - (b.prio || 999)) || cmpAddr(a, b));
  for (const p of fresh) { let i = have.findIndex(q => (q.prio || 999) > (p.prio || 999)); if (i < 0) i = have.length; have.splice(i, 0, p); }
  have.forEach((p, i) => { p.ord = i + 1; });
}
/* ordine automatico: per priorità del foglio; a parità, catena del più vicino partendo dalla tua posizione */
/* posizione "di ordinamento" anche per i pacchi senza pin: stessa via, oppure baricentro del paese */
function sortPos(p) {
  if (p.lat != null) return p;
  const st = sqStreet(p.via), com = normCom(p.comune);
  const pts = DAY.parcels.filter(q => q.lat != null && normCom(q.comune) === com && (!st || sqStreet(q.via) === st));
  const pool = pts.length ? pts : DAY.parcels.filter(q => q.lat != null && normCom(q.comune) === com);
  if (!pool.length) return null;
  return { lat: pool.reduce((a, q) => a + q.lat, 0) / pool.length, lng: pool.reduce((a, q) => a + q.lng, 0) / pool.length };
}
function autoOrder() {
  let cur = meMarker ? meMarker.getLatLng() : null; const by = {};
  DAY.parcels.forEach(p => (by[p.prio || 999] = by[p.prio || 999] || []).push(p));
  const seq = [];
  for (const k of Object.keys(by).map(Number).sort((a, b) => a - b)) {
    const pool = by[k].filter(p => sortPos(p)).sort(cmpAddr), noLL = by[k].filter(p => !sortPos(p)).sort(cmpAddr);
    while (pool.length) {
      let bi = 0; if (cur) { let bd = 1e12; pool.forEach((p, i) => { const d = distM(cur, sortPos(p)); if (d < bd) { bd = d; bi = i; } }); }
      const p = pool.splice(bi, 1)[0]; seq.push(p); cur = sortPos(p);
    }
    seq.push(...noLL);
  }
  seq.forEach((p, i) => { p.ord = i + 1; });
}
function sortedParcels() { ensureOrder(); return DAY.parcels.slice().sort((a, b) => a.ord - b.ord); }
function setStato(id, st) {
  const p = DAY.parcels.find(x => x.id === id); if (!p) return;
  p.stato = st; p.t = nowIso(); archiveParcel(p);
  if (st !== 'da_fare') DAY.doneSince = (DAY.doneSince || 0) + 1;
  touch(); renderPins(); counters(); checkRemind();
}

/* ---------- mappa ---------- */
const map = L.map('map', { zoomControl: false, attributionControl: true }).setView([41.06, 14.74], 12);
const osmLayer = L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, maxNativeZoom: 19, attribution: '© OpenStreetMap' }).addTo(map);
const satLayer = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', { maxZoom: 19, maxNativeZoom: 19, attribution: 'Esri, Maxar, Earthstar Geographics' });
function setNative() { osmLayer.options.maxNativeZoom = navigator.onLine ? 19 : 17; osmLayer.redraw(); }
let satOn = false;
function toggleSat() {
  if (!satOn && !navigator.onLine) return toast('Il satellite richiede la rete');
  satOn = !satOn; if (satOn) { map.addLayer(satLayer); map.removeLayer(osmLayer); } else { map.addLayer(osmLayer); map.removeLayer(satLayer); }
  $('#bSat').textContent = satOn ? '🗺 Mappa' : '🛰 Satellite';
}
L.control.zoom({ position: 'topleft' }).addTo(map);
const pinLayer = L.layerGroup().addTo(map);
let meMarker = null, meCircle = null, firstFix = false;
function startGPS() {
  if (!navigator.geolocation) return;
  navigator.geolocation.watchPosition(pos => {
    const ll = [pos.coords.latitude, pos.coords.longitude];
    if (!meMarker) {
      meMarker = L.marker(ll, { icon: L.divIcon({ className: '', html: '<div class="me"></div>', iconSize: [18, 18], iconAnchor: [9, 9] }), zIndexOffset: 1000, interactive: false }).addTo(map);
      meCircle = L.circle(ll, { radius: pos.coords.accuracy, weight: 1, color: '#0a84ff', fillOpacity: 0.08, interactive: false }).addTo(map);
    } else { meMarker.setLatLng(ll); meCircle.setLatLng(ll).setRadius(pos.coords.accuracy); }
    if (!firstFix) { firstFix = true; if (!DAY.parcels.some(p => p.lat)) map.setView(ll, 14); }
  }, err => { toast('GPS non disponibile: controlla i permessi'); }, { enableHighAccuracy: true, maximumAge: 5000, timeout: 20000 });
}
$('#bSat').onclick = toggleSat;
$('#bLoc').onclick = () => { if (meMarker) map.setView(meMarker.getLatLng(), Math.max(map.getZoom(), 16)); else toast('Posizione non ancora trovata'); };

function groups() {
  const g = {};
  for (const p of DAY.parcels) { if (p.lat == null || p.lng == null) continue; const k = p.lat.toFixed(4) + ',' + p.lng.toFixed(4); (g[k] = g[k] || []).push(p); }
  return g;
}
function renderPins() {
  ensureOrder(); pinLayer.clearLayers();
  const g = groups();
  for (const k in g) {
    const ps = g[k]; const open = ps.filter(p => p.stato === 'da_fare');
    const lead = open[0] || ps[0]; const n = Math.min(...(open.length ? open : ps).map(p => p.ord || 999));
    const cls = open.length ? stClass(lead) : stClass(ps[0]);
    const approx = ps.some(p => p.geo && p.geo !== 'db' && p.geo !== 'pin');
    const html = '<div class="pin ' + cls + (approx ? ' approx' : '') + '">' + (n < 999 ? n : '?') + (ps.length > 1 ? '<sup>' + ps.length + '</sup>' : '') + '</div>';
    const m = L.marker([ps[0].lat, ps[0].lng], { icon: L.divIcon({ className: '', html, iconSize: [30, 30], iconAnchor: [15, 15] }) });
    m.on('click', () => openGroup(ps.map(p => p.id)));
    pinLayer.addLayer(m);
  }
}
function fitPins() {
  const pts = DAY.parcels.filter(p => p.lat != null).map(p => [p.lat, p.lng]);
  if (pts.length) map.fitBounds(pts, { padding: [30, 30], maxZoom: 15 });
}
let NAV = lsGet('sda.nav', 'google');
function navUrl(lat, lng, name) {
  const nm = encodeURIComponent(name || 'Destinazione');
  if (NAV === 'organic') {
    if (meMarker) { const m = meMarker.getLatLng(); return 'om://route?sll=' + m.lat + ',' + m.lng + '&saddr=' + encodeURIComponent('Posizione attuale') + '&dll=' + lat + ',' + lng + '&daddr=' + nm + '&type=vehicle'; }
    return 'om://map?v=1&ll=' + lat + ',' + lng + '&n=' + nm;
  }
  if (NAV === 'geo') return 'geo:' + lat + ',' + lng + '?q=' + lat + ',' + lng + '(' + nm + ')';
  return 'https://www.google.com/maps/dir/?api=1&travelmode=driving&destination=' + lat + ',' + lng;
}
const NAV_LABEL = { google: 'Google Maps', organic: 'Organic Maps', geo: 'Scegli ogni volta (Android)' };
function waUrl(tel) {
  let d = String(tel || '').replace(/[^\d+]/g, ''); if (d.startsWith('+')) d = d.slice(1); else if (d.startsWith('00')) d = d.slice(2); else if (!d.startsWith('39')) d = '39' + d;
  return 'https://wa.me/' + d + '?text=' + encodeURIComponent(WA);
}
function parcelCard(p) {
  const a = DB.addr[akeyP(p)]; const note = a && a.note ? '<div class="note">📝 ' + esc(a.note) + '</div>' : '';
  const done = p.stato !== 'da_fare';
  const S = STATI[p.tipo] || STATI.consegna;
  const act = done
    ? '<div class="row"><div class="tag" style="align-self:center;text-align:center;padding:12px">' + stLabel(p.stato) + '</div><button class="btn gray" data-a="undo" data-id="' + p.id + '">Annulla</button></div>'
    : '<div class="row"><button class="btn ok" data-a="done" data-id="' + p.id + '">' + (p.tipo === 'ritiro' ? 'Ritirato' : 'Consegnato') + '</button><button class="btn no" data-a="no" data-id="' + p.id + '">' + (p.tipo === 'ritiro' ? 'Non ritirato' : 'Non consegnato') + '</button></div>';
  return '<div class="card"><div class="t">' + esc(p.name || '(senza nome)') + '</div><div class="s">' + esc(fmtAddr(p)) + '</div><div style="margin-top:6px">' +
    '<span class="tag">N° ' + p.ord + '</span>' + (p.prio ? '<span class="tag">Priorità ' + p.prio + (p.amb ? ' ?' : '') + '</span>' : '<span class="tag">Da collocare</span>') +
    (p.tipo === 'ritiro' ? '<span class="tag rit">RITIRO</span>' : '') + (p.cod ? '<span class="tag cod">COD ' + esc(p.cod) + ' €</span>' : '') +
    (p.geo && p.geo !== 'db' && p.geo !== 'pin' ? '<span class="tag">posizione da verificare</span>' : '') + '</div>' + note + act +
    '<div class="row">' + (p.tel ? '<a class="btn sec sm" style="text-decoration:none;text-align:center" target="_blank" href="' + waUrl(p.tel) + '">WhatsApp</a>' : '') +
    '<button class="btn sec sm" data-a="note" data-id="' + p.id + '">Note</button><button class="btn sec sm" data-a="more" data-id="' + p.id + '">⋮ Altro</button></div></div>';
}
function openGroup(ids) {
  const ps = ids.map(id => DAY.parcels.find(p => p.id === id)).filter(Boolean); if (!ps.length) return closeSheet();
  const p0 = ps[0];
  const nav = p0.lat != null ? '<a class="btn" style="display:block;text-align:center;text-decoration:none;margin-bottom:6px"' + (NAV === 'google' ? ' target="_blank"' : '') + ' href="' + navUrl(p0.lat, p0.lng, fmtAddr(p0)) + '">➜ Naviga (' + NAV_LABEL[NAV].split(' (')[0] + ')</a>' : '<button class="btn warn" data-a="place" data-id="' + p0.id + '" style="width:100%">Posiziona sulla mappa</button>';
  sheet('<h3>' + esc(p0.via + ' ' + (p0.civico || '')) + '</h3>' + nav + ps.map(parcelCard).join('') + '<button class="btn sec" style="width:100%" data-a="close">Chiudi</button>');
  const root = $('#sheet');
  root.querySelectorAll('[data-a]').forEach(b => b.onclick = ev => {
    const a = b.dataset.a, id = +b.dataset.id;
    if (a === 'close') return closeSheet();
    if (a === 'done') { const p = DAY.parcels.find(x => x.id === id); setStato(id, STATI[p.tipo].done); return nextOrReopen(ids); }
    if (a === 'no') { const p = DAY.parcels.find(x => x.id === id); setStato(id, STATI[p.tipo].no); return nextOrReopen(ids); }
    if (a === 'undo') { setStato(id, 'da_fare'); return openGroup(ids); }
    if (a === 'note') return editNote(id, ids);
    if (a === 'more') return moreMenu(id, ids);
    if (a === 'place') { closeSheet(); return startMove(id); }
  });
}
function nextOrReopen(ids) {
  const left = ids.map(id => DAY.parcels.find(p => p.id === id)).filter(p => p && p.stato === 'da_fare');
  if (left.length) openGroup(ids); else closeSheet();
}
function editNote(id, ids) {
  const p = DAY.parcels.find(x => x.id === id); const a = DB.addr[akeyP(p)];
  sheet('<h3>Note – ' + esc(fmtAddr(p)) + '</h3><p style="font-size:13px;color:#51627d;margin:0">La nota resta sull\'indirizzo e riappare ai prossimi passaggi.</p><textarea id="nt">' + esc(a ? a.note : '') + '</textarea><div class="row"><button class="btn sec" id="c">Annulla</button><button class="btn ok" id="o">Salva</button></div>');
  $('#c').onclick = () => openGroup(ids);
  $('#o').onclick = () => { saveAddr(p, { note: $('#nt').value.trim(), lat: a ? a.lat : p.lat, lng: a ? a.lng : p.lng }); toast('Nota salvata'); openGroup(ids); };
}
function moreMenu(id, ids) {
  const p = DAY.parcels.find(x => x.id === id);
  sheet('<h3>' + esc(fmtAddr(p)) + '</h3><div style="display:flex;flex-direction:column;gap:8px">' +
    '<button class="btn" data-m="move">Sposta / correggi pin</button><button class="btn sec" data-m="edit">Modifica dati</button><button class="btn sec" data-m="prio">Assegna priorità</button>' +
    '<button class="btn sec" data-m="geo">Cerca posizione online</button><button class="btn no" data-m="del">Elimina pacco</button><button class="btn gray" data-m="back">Indietro</button></div>');
  $('#sheet').querySelectorAll('[data-m]').forEach(b => b.onclick = () => {
    const m = b.dataset.m;
    if (m === 'back') return openGroup(ids);
    if (m === 'move') { closeSheet(); return startMove(id); }
    if (m === 'edit') return parcelForm(p, () => { renderPins(); openGroup(ids); });
    if (m === 'prio') return prioPicker(p, ids);
    if (m === 'geo') { closeSheet(); return geocodeOne(p); }
    if (m === 'del') { if (confirm('Eliminare questo pacco dalla lista?')) { DAY.parcels = DAY.parcels.filter(x => x.id !== id); touch(); renderPins(); counters(); closeSheet(); } }
  });
}
function prioPicker(p, ids) {
  sheet('<h3>Priorità per ' + esc(fmtAddr(p)) + '</h3><label>Numero del foglio (1–66, vuoto = da collocare)</label><input id="pn" type="number" value="' + (p.prio || '') + '"><div class="row"><button class="btn sec" id="c">Annulla</button><button class="btn ok" id="o">Salva</button></div>');
  $('#c').onclick = () => openGroup(ids);
  $('#o').onclick = () => { const n = parseInt($('#pn').value, 10) || 0; p.prio = n; p.amb = false; saveAddr(p, { prio: n, lat: p.lat, lng: p.lng }); renderPins(); openGroup(ids); };
}

/* ---------- sposta / posiziona pin ---------- */
let moving = null;
async function startMove(id) {
  const p = DAY.parcels.find(x => x.id === id); if (!p) return;
  showView('map');
  let start;
  if (p.lat != null) { start = [p.lat, p.lng]; map.setView(start, Math.max(map.getZoom(), 17)); }
  else {
    let g = guessCenter(p); if (!g && navigator.onLine) { busy('Cerco ' + baseComune(p.comune) + '…'); g = await fetchComuneCenter(p); busy(); }
    if (g) { start = g; map.setView(g, 16); toast('Indirizzo non trovato da solo: usa Cerca via… oppure trascina il pin', 5000); }
    else start = [map.getCenter().lat, map.getCenter().lng];
  }
  const m = L.marker(start, { draggable: true, icon: L.divIcon({ className: '', html: '<div class="pin st-x" style="width:36px;height:36px">⊕</div>', iconSize: [36, 36], iconAnchor: [18, 18] }), zIndexOffset: 2000 }).addTo(map);
  moving = { p, m };
  const tb = $('#topbar'); tb.style.display = 'block';
  tb.innerHTML = '<b>' + esc(fmtAddr(p)) + '</b><div style="font-size:13px;color:#51627d">Trascina il pin (o tocca la mappa). Non vedi la via? Cercala:</div>' +
    '<div class="row" style="margin:6px 0"><input id="mvQ" placeholder="Cerca via o luogo…" value="' + esc(p.via) + '"><button class="btn sm" id="mvG" style="flex:.45">Cerca</button></div><div class="row" style="margin:4px 0"><a class="btn sec sm" style="text-decoration:none;text-align:center" target="_blank" href="https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(fmtAddr(p) + ', Italia') + '">Cerca su Google Maps</a><button class="btn sec sm" id="mvP">Incolla coordinate, link o Plus code</button></div><div class="row" style="margin:4px 0"><button class="btn sec sm" id="mvMe">📍 Usa la mia posizione</button></div><div id="mvR" style="max-height:30vh;overflow:auto"></div>' +
    '<div class="row"><button class="btn sec sm" id="mvC">Annulla</button><button class="btn ok sm" id="mvS">Salva</button></div>';
  const doSearch = async () => {
    const q = $('#mvQ').value.trim(); if (!q) return; const box = $('#mvR');
    const cc0 = parseCoords(q) || parsePlus(q, [m.getLatLng().lat, m.getLatLng().lng]);
    if (cc0) { m.setLatLng(cc0); map.setView(cc0, 19); box.innerHTML = '<div class="s" style="padding:6px">📍 Pin spostato su ' + cc0[0].toFixed(6) + ', ' + cc0[1].toFixed(6) + '. Controlla e premi Salva.</div>'; return; } box.innerHTML = '<div class="s" style="padding:6px">Cerco…</div>';
    const res = await searchPlaces(q, p);
    box.innerHTML = res.length ? res.map((r, i) => '<div class="it" data-i="' + i + '" style="padding:8px;border-bottom:1px solid #dde3ec"><div class="x"><b style="font-size:14px;white-space:normal">' + esc(r.label) + '</b><span>' + (r.src === 'db' ? 'dal database' : r.src === 'gps' ? 'da Google Maps / GPS' : r.src === 'info' ? '' : 'online') + '</span></div></div>').join('') : '<div class="s" style="padding:6px">Nessun risultato' + (navigator.onLine ? '' : ' (senza rete cerco solo nel database)') + '</div>';
    box.querySelectorAll('[data-i]').forEach(el => el.onclick = () => { const r = res[+el.dataset.i]; if (r.lat == null) return; m.setLatLng([r.lat, r.lng]); map.setView([r.lat, r.lng], 18); box.innerHTML = ''; });
  };
  $('#mvP').onclick = async () => { try { const t = await navigator.clipboard.readText(); $('#mvQ').value = t; doSearch(); } catch (e) { toast('Incolla a mano nel campo di ricerca (tieni premuto → Incolla)'); } };
  $('#mvMe').onclick = () => { if (!meMarker) return toast('GPS non ancora attivo: premi il tasto Io e aspetta il punto blu', 4000); const ll = meMarker.getLatLng(); m.setLatLng(ll); map.setView(ll, 19); toast('Pin portato sulla tua posizione: premi Salva', 3500); };
  let taT = null, taSeq = 0;
  $('#mvQ').addEventListener('input', () => {
    const q = $('#mvQ').value.trim(), box = $('#mvR'), seq = ++taSeq; clearTimeout(taT);
    if (q.length < 3 || parseCoords(q) || /^[0-9A-Z]{4,8}\+/i.test(q)) { if (q.length < 3) box.innerHTML = ''; return; }
    const row = (lab, sub, i) => '<div class="it" data-t="' + i + '" style="padding:8px;border-bottom:1px solid #dde3ec"><div class="x"><b style="font-size:14px;white-space:normal">' + esc(lab) + '</b><span>' + sub + '</span></div></div>';
    const loc = suggestStreets(q); const items = [];
    loc.forEach(e => items.push({ label: e.via + (e.com ? ', ' + e.com : ''), lat: e.pts.length ? e.pts[0].lat : null, lng: e.pts.length ? e.pts[0].lng : null, txt: e.via + (e.com ? ' ' + e.com : '') }));
    const draw = (extra) => { const all = items.concat(extra || []); box.innerHTML = all.map((r, i) => row(r.label, r.lat == null ? 'tocca per cercare' : (r.on ? 'online' : 'dal database'), i)).join(''); box.querySelectorAll('[data-t]').forEach(el => el.onclick = () => { const r = all[+el.dataset.t]; if (r.lat != null) { m.setLatLng([r.lat, r.lng]); map.setView([r.lat, r.lng], 18); box.innerHTML = ''; } else { $('#mvQ').value = r.txt || r.label; doSearch(); } }); };
    if (items.length) draw();
    if (!navigator.onLine || q.length < 4) return;
    taT = setTimeout(async () => { try { const ex = (await photonSuggest(q + ' ' + baseComune(p.comune))).map(x => Object.assign(x, { on: 1 })); if (seq === taSeq) draw(ex); } catch (e) { } }, 400);
  });
  $('#mvG').onclick = doSearch; $('#mvQ').addEventListener('keydown', e => { if (e.key === 'Enter') doSearch(); });
  moving.h = e => { if (moving) m.setLatLng(e.latlng); }; map.on('click', moving.h);
  $('#mvC').onclick = endMove;
  $('#mvS').onclick = () => {
    const ll = m.getLatLng(); const key = akeyP(p);
    DAY.parcels.filter(x => akeyP(x) === key).forEach(x => { x.lat = ll.lat; x.lng = ll.lng; x.geo = 'pin'; });
    saveAddr(p, { lat: ll.lat, lng: ll.lng, via: p.via, civico: p.civico, cap: p.cap, comune: p.comune, prio: (DB.addr[key] || {}).prio || 0 });
    endMove(); renderPins(); toast('Posizione salvata nel database');
  };
}
function endMove() { if (moving) { map.off('click', moving.h); map.removeLayer(moving.m); moving = null; } $('#topbar').style.display = 'none'; }

/* ---------- geocoding online (OpenStreetMap) ---------- */
const BBOX = { s: 40.7, n: 41.4, w: 14.3, e: 15.2 };
const inBox = (la, lo) => la > BBOX.s && la < BBOX.n && lo > BBOX.w && lo < BBOX.e;
const sleep = ms => new Promise(r => setTimeout(r, ms));
async function nominatim(params) {
  const r = await fetch('https://nominatim.openstreetmap.org/search?format=json&limit=5&addressdetails=1&countrycodes=it&' + params);
  return r.json();
}
function placeMatches(item, com) {
  const a = item.address || {};
  return [a.town, a.city, a.village, a.municipality, a.hamlet, a.suburb].filter(Boolean).map(norm).some(n => n.includes(com) || com.includes(n));
}
const ADMIN_TYPES = new Set(['city', 'town', 'village', 'hamlet', 'suburb', 'municipality', 'quarter', 'neighbourhood', 'county', 'state', 'region', 'locality', 'administrative']);
function validHit(it, com) {
  if (!inBox(+it.lat, +it.lon)) return false;
  if (it.class === 'boundary' || (it.class === 'place' && ADMIN_TYPES.has(it.type))) return false;
  return placeMatches(it, com);
}
async function geocode(p) {
  const com = normCom(p.comune), cname = baseComune(p.comune);
  const civ = p.civico && !/^s\.?n\.?c\.?$/i.test(p.civico) ? p.civico.trim() : '';
  const tries = [
    () => nominatim('q=' + encodeURIComponent(p.via + ' ' + civ + ', ' + cname + ', Italia')),
    () => nominatim('street=' + encodeURIComponent((civ ? civ + ' ' : '') + p.via) + '&city=' + encodeURIComponent(cname)),
    () => nominatim('q=' + encodeURIComponent(p.via + ', ' + cname + ', Italia'))
  ];
  for (let i = 0; i < tries.length; i++) {
    try { const j = await tries[i](); const h = j.find(it => validHit(it, com)); if (h) return [+h.lat, +h.lon, (h.class === 'building' || h.type === 'house') ? 'exact' : 'street']; } catch (e) { }
    await sleep(1100);
  }
  return null;
}

/* ---------- Plus code (Open Location Code) ---------- */
const OLC = '23456789CFGHJMPQRVWX';
function olcEncode(lat, lng) {
  let la = Math.floor((lat + 90) * 8000), lo = Math.floor((lng + 180) * 8000), c = '';
  for (let i = 0; i < 5; i++) { c = OLC[la % 20] + OLC[lo % 20] + c; la = Math.floor(la / 20); lo = Math.floor(lo / 20); }
  return c;
}
function olcDecode(code) {
  let la = -90, lo = -180, res = 20;
  for (let i = 0; i < Math.min(code.length, 10); i += 2) {
    la += OLC.indexOf(code[i]) * res; lo += OLC.indexOf(code[i + 1]) * res; if (i < 8) res /= 20;
  }
  const size = code.length >= 10 ? 0.000125 : 20 / Math.pow(20, Math.floor(code.length / 2) - 1);
  return [la + size / 2, lo + size / 2];
}
function parsePlus(t, ref) {
  const m = String(t || '').toUpperCase().match(/(?:^|[\s,])([23456789CFGHJMPQRVWX]{2,8})\+([23456789CFGHJMPQRVWX]{2,3})(?![23456789CFGHJMPQRVWX])/);
  if (!m) return null;
  let head = m[1]; const tail = m[2];
  if (head.length < 8) {
    if (!ref) return null;
    const pad = 8 - head.length, resolution = Math.pow(20, 2 - pad / 2);
    head = olcEncode(ref[0], ref[1]).slice(0, pad) + head;
    let [la, lo] = olcDecode(head + tail);
    if (ref[0] + resolution / 2 < la && la - resolution >= -90) la -= resolution; else if (ref[0] - resolution / 2 > la && la + resolution <= 90) la += resolution;
    if (ref[1] + resolution / 2 < lo) lo -= resolution; else if (ref[1] - resolution / 2 > lo) lo += resolution;
    return inBox(la, lo) ? [la, lo] : null;
  }
  const [la, lo] = olcDecode(head + tail); return inBox(la, lo) ? [la, lo] : null;
}
function parseCoords(t) {
  t = String(t || '').trim(); if (!t) return null;
  const m = t.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/) || t.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) || t.match(/[?&](?:q|ll|query|destination|center)=(-?\d+\.\d+)(?:,|%2C)\s*(-?\d+\.\d+)/i) || t.match(/(-?\d{1,2}[.,]\d{3,})\s*[,; ]\s*(-?\d{1,3}[.,]\d{3,})/);
  if (!m) return null;
  const la = parseFloat(m[1].replace(',', '.')), lo = parseFloat(m[2].replace(',', '.'));
  return isFinite(la) && isFinite(lo) && inBox(la, lo) ? [la, lo] : null;
}

/* ---------- suggerimenti mentre scrivi (database offline + Photon online) ---------- */
function lev(a, b) { const m = a.length, n = b.length; if (Math.abs(m - n) > 2) return 9; let pr = Array.from({ length: n + 1 }, (_, j) => j); for (let i = 1; i <= m; i++) { const cu = [i]; for (let j = 1; j <= n; j++) cu[j] = Math.min(pr[j] + 1, cu[j - 1] + 1, pr[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)); pr = cu; } return pr[n]; }
function knownStreets() {
  const m = new Map();
  const add = (via, com, a) => { if (!via) return; const k = sqStreet(via) + '|' + normCom(com); let e = m.get(k); if (!e) { e = { via: String(via).trim(), com: com || '', pts: [] }; m.set(k, e); } if (a && a.lat != null) e.pts.push(a); };
  for (const a of Object.values(DB.addr)) add(a.via, a.comune, a);
  for (const q of (DAY.parcels || [])) add(q.via, q.comune, q.lat != null ? q : null);
  return [...m.values()];
}
function suggestStreets(q) {
  const sq = normStreet(q).replace(/ /g, ''); if (sq.length < 3) return [];
  const sc = [];
  for (const e of knownStreets()) {
    const t = sqStreet(e.via); let s = 0;
    if (t.startsWith(sq)) s = 3; else if (t.includes(sq)) s = 2; else if (sq.length >= 5 && lev(t.slice(0, sq.length + 1), sq) <= 1) s = 1; else if (sq.length >= 5 && lev(t, sq) <= 2) s = 1;
    if (s) sc.push([s, e]);
  }
  return sc.sort((a, b) => b[0] - a[0]).slice(0, 6).map(x => x[1]);
}
async function photonSuggest(q) {
  const r = await fetch('https://photon.komoot.io/api/?limit=5&lang=it&lat=41.05&lon=14.75&q=' + encodeURIComponent(q));
  const j = await r.json(); const out = [];
  for (const f of (j.features || [])) { const [lng, lat] = f.geometry.coordinates; if (!inBox(lat, lng)) continue; const pr = f.properties || {}; const lab = [pr.name || pr.street, pr.housenumber, pr.city || pr.town || pr.village || pr.county].filter(Boolean).join(' '); if (lab) out.push({ label: lab, lat, lng }); }
  return out;
}
async function searchPlaces(q, p) {
  const out = [], nq = norm(q);
  const cc = parseCoords(q); if (cc) { out.push({ label: 'Coordinate incollate: ' + cc[0].toFixed(6) + ', ' + cc[1].toFixed(6), lat: cc[0], lng: cc[1], src: 'gps' }); return out; }
  if (/(maps\.app\.goo\.gl|goo\.gl\/maps)/i.test(q)) { out.push({ label: 'Link breve di Google Maps: non lo posso aprire. Apri il link in Chrome e incolla qui l\'indirizzo lungo, oppure copia le coordinate (tieni premuto sul punto in Maps).', lat: null, lng: null, src: 'info' }); return out; }
  for (const a of Object.values(DB.addr)) if (a.lat != null && norm(a.via + ' ' + a.civico + ' ' + a.comune).includes(nq)) out.push({ label: (a.via + ' ' + a.civico + ', ' + a.comune).trim(), lat: a.lat, lng: a.lng, src: 'db' });
  if (navigator.onLine) {
    const com = baseComune(p.comune); const hasCom = norm(q).includes(normCom(p.comune));
    const add = j => { for (const it of j) { if (!inBox(+it.lat, +it.lon) || (it.class === 'boundary')) continue; const lab = String(it.display_name || '').split(',').slice(0, 3).join(',').trim(); if (!out.some(o => Math.abs(o.lat - it.lat) < 1e-5 && Math.abs(o.lng - it.lon) < 1e-5)) out.push({ label: lab, lat: +it.lat, lng: +it.lon, src: 'osm' }); } };
    try { add(await nominatim('q=' + encodeURIComponent(hasCom ? q : q + ', ' + com + ', Italia'))); } catch (e) { }
    if (out.filter(o => o.src === 'osm').length === 0) { try { add(await nominatim('viewbox=14.3,41.4,15.2,40.7&bounded=1&q=' + encodeURIComponent(q))); } catch (e) { } }
  }
  return out.slice(0, 8);
}
const CC_FALLBACK = { 'ceppaloni': [41.0553, 14.7347], 'san leucio': [41.0647, 14.7519], 'apollosa': [41.085, 14.700] };
function guessCenter(p) {
  const com = normCom(p.comune); const pts = [];
  if (/^ceppaloni/.test(com) && String(p.cap || '').trim() === '82010') { const sb = [41.0462, 14.7407]; const known = Object.values(DB.addr).filter(a => a.lat != null && normCom(a.comune) === com && Math.abs(a.lat - sb[0]) < 0.02 && Math.abs(a.lng - sb[1]) < 0.02); if (!known.length) return sb; }
  for (const a of Object.values(DB.addr)) if (a.lat != null && normCom(a.comune) === com) pts.push([a.lat, a.lng]);
  for (const x of DAY.parcels) if (x.lat != null && x !== p && normCom(x.comune) === com) pts.push([x.lat, x.lng]);
  if (pts.length) return [pts.reduce((s, q) => s + q[0], 0) / pts.length, pts.reduce((s, q) => s + q[1], 0) / pts.length];
  const cc = lsGet('sda.cc', {}); if (cc[com]) return cc[com];
  for (const k in CC_FALLBACK) if (com.startsWith(k)) return CC_FALLBACK[k];
  return null;
}
async function fetchComuneCenter(p) {
  try {
    const com = normCom(p.comune); const j = await nominatim('city=' + encodeURIComponent(baseComune(p.comune)));
    const h = j.find(it => inBox(+it.lat, +it.lon)); if (!h) return null;
    const cc = lsGet('sda.cc', {}); cc[com] = [+h.lat, +h.lon]; lsSet('sda.cc', cc); return cc[com];
  } catch (e) { return null; }
}
let gq = [], gbusy = false;
function queueGeo(p) { if (p.lat != null || !navigator.onLine) return; gq.push(p.id); if (!gbusy) runGeo(); }
async function runGeo() {
  gbusy = true;
  while (gq.length) {
    const p = DAY.parcels.find(x => x.id === gq.shift()); if (!p || p.lat != null) continue;
    const g = await geocode(p); if (g) { p.lat = g[0]; p.lng = g[1]; p.geo = 'osm'; touch(); renderPins(); renderList(); }
    await sleep(1100);
  }
  gbusy = false;
}
async function geocodeOne(p) {
  if (!navigator.onLine) { toast('Senza rete: posiziona a mano'); return startMove(p.id); }
  busy('Cerco la posizione…'); const g = await geocode(p); busy();
  if (!g) { toast('Non trovato: posiziona a mano'); return startMove(p.id); }
  p.lat = g[0]; p.lng = g[1]; p.geo = 'osm'; touch(); renderPins(); map.setView(g, 17); toast('Trovato: controlla il pin e correggi se serve');
}
function lookupDB(p) {
  const exact = DB.addr[akeyP(p)];
  if (exact && exact.lat != null) return { lat: exact.lat, lng: exact.lng, geo: 'db' };
  const st = sqStreet(p.via), com = normCom(p.comune), civ = parseInt(p.civico, 10);
  if (!isFinite(civ)) { /* senza numero (snc): centro dei punti già noti di quella via */
    const pts = Object.values(DB.addr).filter(a => a.lat != null && sqStreet(a.via) === st && normCom(a.comune) === com);
    return pts.length ? { lat: pts.reduce((x, a) => x + a.lat, 0) / pts.length, lng: pts.reduce((x, a) => x + a.lng, 0) / pts.length, geo: 'near' } : null;
  }
  /* punti noti della stessa via con numero: se il civico cade tra due numeri noti lo interpolo, altrimenti uso il più vicino */
  const kn = [];
  for (const a of Object.values(DB.addr)) {
    if (a.lat == null || sqStreet(a.via) !== st || normCom(a.comune) !== com) continue;
    const c = parseInt(a.civico, 10); if (isFinite(c)) kn.push({ c, lat: a.lat, lng: a.lng });
  }
  if (!kn.length) return null;
  kn.sort((x, y) => x.c - y.c);
  const lo = kn.filter(k => k.c <= civ).pop(), hi = kn.find(k => k.c >= civ);
  if (lo && hi && lo.c !== hi.c) { const t = (civ - lo.c) / (hi.c - lo.c); return { lat: lo.lat + (hi.lat - lo.lat) * t, lng: lo.lng + (hi.lng - lo.lng) * t, geo: 'near' }; }
  const best = kn.reduce((m, k) => Math.abs(k.c - civ) < Math.abs(m.c - civ) ? k : m);
  return { lat: best.lat, lng: best.lng, geo: 'near' };
}
async function elabora() {
  const ps = DAY.parcels; let n = 0, tot = ps.filter(p => p.lat == null).length;
  for (const p of ps) {
    if (p.lat != null) continue;
    n++; busy('Elaboro ' + n + '/' + tot + '…');
    const f = lookupDB(p);
    if (f) { Object.assign(p, f); continue; }
    if (navigator.onLine) { const g = await geocode(p); if (g) { p.lat = g[0]; p.lng = g[1]; p.geo = 'osm'; } await new Promise(r => setTimeout(r, 1100)); }
  }
  for (const p of ps) { if (!p.prio) { const m = matchPrio(p); p.prio = m.n; p.amb = m.amb; } }
  if (!DAY.manual) autoOrder();
  touch(); busy(); renderPins(); counters(); renderList();
  const miss = ps.filter(p => p.lat == null).length;
  toast('Pronto. ' + (miss ? miss + ' da posizionare a mano.' : 'Tutti posizionati.'), 4000);
  showView('map'); fitPins();
}

/* ---------- viste ---------- */
let curView = 'map', listFilter = 'tutti';
function showView(v) {
  if (v === 'add') return addMenu();
  curView = v;
  document.querySelectorAll('.view').forEach(x => x.classList.toggle('on', x.id === 'v-' + v));
  document.querySelectorAll('nav button').forEach(b => b.classList.toggle('on', b.dataset.v === v));
  if (v === 'map') setTimeout(() => map.invalidateSize(), 50);
  if (v === 'list') renderList();
  if (v === 'menu') renderMenu();
  if (v === 'scan') startCam(); else stopCam();
}
document.querySelectorAll('nav button').forEach(b => b.onclick = () => showView(b.dataset.v));

function renderList() {
  const el = $('#v-list'); const f = listFilter;
  let ps = sortedParcels();
  if (f === 'da fare') ps = ps.filter(p => p.stato === 'da_fare');
  if (f === 'fatti') ps = ps.filter(p => p.stato !== 'da_fare');
  if (f === 'ritiri') ps = ps.filter(p => p.tipo === 'ritiro');
  if (f === 'da posizionare') ps = ps.filter(p => p.lat == null);
  const chips = ['tutti', 'da fare', 'fatti', 'ritiri', 'da posizionare'].map(c => '<button class="' + (c === f ? 'on' : '') + '" data-f="' + c + '">' + c + '</button>').join('');
  el.innerHTML = '<div class="chips">' + chips + '</div><div class="list">' + (ps.map(p => '<div class="it" data-id="' + p.id + '"><div class="n ' + stClass(p) + '">' + p.ord + '</div><div class="x"><b>' + esc(p.name || '(senza nome)') + '</b><span>' + esc(fmtAddr(p)) + (p.tipo === 'ritiro' ? ' · RITIRO' : '') + (p.lat == null ? ' · da posizionare' : '') + '</span></div><span style="font-size:12px;color:#51627d">' + (p.stato !== 'da_fare' ? '✔' : '') + '</span>' + (f === 'tutti' ? '<span class="drag">☰</span>' : '') + '</div>').join('') || '<div style="padding:24px;text-align:center;color:#51627d">Nessun pacco. Usa lo Scanner, il tasto + oppure Menu → Importa file.</div>') + '</div>' +
    '<div class="legend">Legenda:<span class="st-c"></span>consegna<span class="st-cd"></span>consegnato<span class="st-cn"></span>assente<span class="st-r"></span>ritiro<span class="st-rd"></span>ritirato<span class="st-rn"></span>non ritirato</div>' +
    '<div class="menu"><button class="btn sec" id="bAuto">Riordina in automatico (priorità + vicinanza)</button><button class="btn" id="bEl">Elabora (crea i pin sulla mappa)</button><button class="btn no sm" id="bClr">Cancella tutti i pacchi di oggi</button></div>';
  el.querySelectorAll('[data-f]').forEach(b => b.onclick = () => { listFilter = b.dataset.f; renderList(); });
  el.querySelectorAll('.it').forEach(it => it.onclick = () => {
    const p = DAY.parcels.find(x => x.id === +it.dataset.id);
    if (p.lat == null) return openGroup([p.id]);
    showView('map'); map.setView([p.lat, p.lng], 17);
    const g = groups()[p.lat.toFixed(4) + ',' + p.lng.toFixed(4)] || [p]; openGroup(g.map(x => x.id));
  });
  $('#bAuto').onclick = () => { if (!confirm('Rifaccio l\'ordine di consegna: prima per priorità del foglio, poi il più vicino a te. Perdi il riordino fatto a mano. Continuo?')) return; DAY.manual = false; autoOrder(); touch(); renderPins(); renderList(); toast('Ordine rifatto', 2500); };
  enableDrag(el);
  $('#bEl').onclick = elabora;
  $('#bClr').onclick = () => { if (DAY.parcels.length && confirm('Cancellare tutti i ' + DAY.parcels.length + ' pacchi di oggi dalla lista? Il database indirizzi resta.')) { DAY.parcels = []; DAY.nextId = 1; touch(); renderPins(); counters(); renderList(); } };
}

function enableDrag(el) {
  el.querySelectorAll('.drag').forEach(h => {
    h.addEventListener('pointerdown', e => {
      e.preventDefault(); e.stopPropagation();
      const it = h.closest('.it'), list = it.parentNode; it.classList.add('dragging');
      const mv = ev => {
        if (ev.clientY < 110) { el.scrollBy(0, -14); window.scrollBy(0, -14); } else if (ev.clientY > innerHeight - 130) { el.scrollBy(0, 14); window.scrollBy(0, 14); }
        const t = document.elementFromPoint(ev.clientX, ev.clientY), o = t && t.closest && t.closest('.it');
        if (o && o !== it && o.parentNode === list) { const r = o.getBoundingClientRect(); list.insertBefore(it, ev.clientY < r.top + r.height / 2 ? o : o.nextSibling); }
      };
      const up = () => {
        document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); document.removeEventListener('pointercancel', up); it.classList.remove('dragging');
        [...list.querySelectorAll('.it')].forEach((x, i) => { const p = DAY.parcels.find(q => q.id === +x.dataset.id); if (p) p.ord = i + 1; });
        DAY.manual = true; touch(); renderPins(); renderList();
      };
      document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up); document.addEventListener('pointercancel', up);
    });
    h.addEventListener('click', e => e.stopPropagation());
  });
}
function distM(a, b) { const R = 6371000, r = Math.PI / 180, dLa = (b.lat - a.lat) * r, dLo = (b.lng - a.lng) * r; const x = Math.sin(dLa / 2) ** 2 + Math.cos(a.lat * r) * Math.cos(b.lat * r) * Math.sin(dLo / 2) ** 2; return 2 * R * Math.asin(Math.sqrt(x)); }
/* prossima tappa: la priorità più bassa tra i pacchi aperti; a parità sceglie il più vicino alla tua posizione */
function nextStop() {
  ensureOrder();
  const open = DAY.parcels.filter(p => p.stato === 'da_fare' && p.lat != null).sort((a, b) => a.ord - b.ord);
  if (!open.length) return null;
  const me = meMarker ? meMarker.getLatLng() : null;
  return { p: open[0], d: me ? distM(me, open[0]) : null, prio: open[0].ord };
}
$('#bNext').onclick = () => {
  const r = nextStop();
  if (!r) return toast('Nessuna tappa da fare con posizione');
  const p = r.p; map.setView([p.lat, p.lng], Math.max(map.getZoom(), 16));
  const g = groups()[p.lat.toFixed(4) + ',' + p.lng.toFixed(4)] || [p]; openGroup(g.map(x => x.id));
  toast('Consegna n° ' + r.prio + (r.d != null ? ' · a ' + (r.d < 1000 ? Math.round(r.d) + ' m' : (r.d / 1000).toFixed(1) + ' km') : ''), 3500);
};

/* ---------- moduli di inserimento ---------- */
function parcelForm(p, onOk, isNew) {
  const a = DB.addr[akeyP(p)];
  sheet('<h3>' + (isNew ? 'Nuovo ' + (p.tipo === 'ritiro' ? 'ritiro' : 'pacco') : 'Dati pacco') + '</h3>' +
    '<label>Destinatario</label><input id="f_name" value="' + esc(p.name) + '">' +
    '<div class="row"><div><label>Via</label><input id="f_via" list="dl_via" autocomplete="off" value="' + esc(p.via) + '"><datalist id="dl_via">' + knownStreets().map(e => '<option value="' + esc(e.via) + '">').join('') + '</datalist></div><div style="flex:.5"><label>Civico</label><input id="f_civ" value="' + esc(p.civico) + '"></div></div>' +
    '<div class="row"><div style="flex:.6"><label>CAP</label><input id="f_cap" value="' + esc(p.cap) + '" inputmode="numeric"></div><div><label>Città</label><input id="f_com" value="' + esc(p.comune) + '"></div></div>' +
    '<div class="row"><div><label>Telefono (verifica che sia del destinatario)</label><input id="f_tel" value="' + esc(p.tel) + '" inputmode="tel"></div><div style="flex:.6"><label>COD €</label><input id="f_cod" value="' + esc(p.cod) + '"></div></div>' +
    '<label>Tipo</label><select id="f_tipo"><option value="consegna"' + (p.tipo === 'consegna' ? ' selected' : '') + '>Consegna</option><option value="ritiro"' + (p.tipo === 'ritiro' ? ' selected' : '') + '>Ritiro</option></select>' +
    '<label>Note sull\'indirizzo (restano nel database)</label><textarea id="f_note">' + esc(a ? a.note : '') + '</textarea>' +
    (a && a.lat != null ? '<div class="tag" style="margin-top:6px">posizione già nel database</div>' : '') +
    '<div class="row"><button class="btn sec" id="f_c">Annulla</button><button class="btn ok" id="f_o">OK</button></div>', () => { if (typeof p._cancel === 'function') p._cancel(); });
  let ok = false;
  $('#f_via').addEventListener('change', () => { /* se è la stessa via scritta diversa (spazi), uso il nome già nel database */
    const v = $('#f_via').value.trim(), k = sqStreet(v); if (!k) return;
    const hit = knownStreets().find(e => sqStreet(e.via) === k && e.via !== v && normCom(e.com) === normCom($('#f_com').value));
    if (hit) { $('#f_via').value = hit.via; toast('Via già nota: ' + hit.via, 2500); }
  });
  $('#f_c').onclick = () => { closeSheet(); };
  $('#f_o').onclick = () => {
    const q = { name: $('#f_name').value.trim(), via: $('#f_via').value.trim(), civico: $('#f_civ').value.trim(), cap: $('#f_cap').value.trim(), comune: $('#f_com').value.trim(), tel: $('#f_tel').value.trim(), cod: $('#f_cod').value.trim(), tipo: $('#f_tipo').value };
    if (!q.via) return toast('Manca la via');
    const moved = akey(q.via, q.civico, q.comune) !== akeyP(p);
    Object.assign(p, q); if (moved) { p.lat = null; p.lng = null; p.geo = ''; p.prio = 0; }
    const note = $('#f_note').value.trim(); const cur = DB.addr[akeyP(p)];
    if ((cur ? cur.note : '') !== note) saveAddr(p, { note, lat: cur ? cur.lat : p.lat, lng: cur ? cur.lng : p.lng, prio: cur ? cur.prio : 0 });
    if (isNew) {
      p._cancel = null;
      if (DAY.parcels.some(x => x.track && x.track === p.track && x !== p)) { toast('Pacco già in lista'); }
      else { DAY.parcels.push(p); }
      const f = lookupDB(p); if (f && p.lat == null) Object.assign(p, f);
      if (!p.prio) { const m = matchPrio(p); p.prio = m.n; p.amb = m.amb; }
      queueGeo(p);
    }
    ok = true; touch(); counters(); renderPins(); renderList(); sheetOnClose = null; closeSheet(); if (onOk) onOk(p);
  };
}
function addMenu() {
  sheet('<h3>Aggiungi</h3><div style="display:flex;flex-direction:column;gap:8px"><button class="btn" id="a1">Nuovo pacco da consegnare</button><button class="btn warn" id="a2">Nuovo ritiro</button><button class="btn sec" id="a3">Annulla</button></div>');
  $('#a1').onclick = () => parcelForm(newParcel({}), null, true);
  $('#a2').onclick = () => parcelForm(newParcel({ tipo: 'ritiro' }), null, true);
  $('#a3').onclick = closeSheet;
}
$('#bManual').onclick = () => parcelForm(newParcel({}), () => toast('Aggiunto'), true);

/* ---------- scanner ---------- */
let scanStart = 0, scanN = 0, torchOn = false, oneD = {};
let stream = null, scanTimer = null, fast = false, scanning = false, detector = null, lastCode = '', lastTime = 0;
/* è l'indirizzo SDA solo se ha i campi separati da '|' e un CAP a 5 cifre: QR interni e codici a barre vengono ignorati */
function isAddr(t) { const f = String(t || '').split('|'); return f.length >= 6 && f.some((x, i) => i >= 2 && /^\s*\d{5}\s*$/.test(x)); }
function parsePayload(txt) {
  const f = String(txt).split('|').map(s => s.trim());
  const d = { track: f[0] || '', name: '', via: '', civico: '', cap: '', comune: '', prov: '', tel: '', cod: '', tipo: 'consegna' };
  if (f.length < 6) return d;
  const ci = f.findIndex((x, i) => i >= 2 && /^\d{5}$/.test(x));
  if (ci > 0) {
    d.cap = f[ci]; d.comune = (f[ci + 1] || '').toLowerCase().replace(/(^|[\s'-])\p{L}/gu, m => m.toUpperCase()).replace(/ Di /g, ' di ');
    if (/^[A-Za-z]{2}$/.test(f[ci + 2] || '')) d.prov = f[ci + 2].toUpperCase();
    const addr = f[ci - 1] || '';
    const m = addr.match(/^(.*?)[,\s]+(?:n\.?|nr\.?|civ\.?)?\s*(\d+\s*[a-zA-Z]?(?:\/\w+)?|s\.?n\.?c\.?)\s*$/i);
    if (m) { d.via = m[1].replace(/[,\s]+$/, ''); d.civico = m[2].replace(/\s+/g, ''); } else d.via = addr;
    d.via = d.via.toLowerCase().replace(/(^|[\s'.-])\p{L}/gu, x => x.toUpperCase());
    for (let i = ci - 2; i >= 1; i--) { if (f[i] && /\p{L}/u.test(f[i]) && !/^\d+$/.test(f[i])) { d.name = f[i].toLowerCase().replace(/(^|[\s'-])\p{L}/gu, x => x.toUpperCase()); break; } }
  }
  for (let i = (ci > 0 ? ci + 3 : 6); i < f.length; i++) { if (/^(\+?39)?3\d{8,9}$/.test(f[i].replace(/\s/g, ''))) { d.tel = f[i].replace(/\s/g, ''); break; } }
  const ic = f.findIndex(x => /COD/i.test(x));
  if (ic >= 0) for (let i = ic + 1; i <= ic + 3 && i < f.length; i++) { if (/^\d+([.,]\d{1,2})?$/.test(f[i])) { d.cod = f[i]; break; } }
  return d;
}
async function startCam() {
  if (stream) return;
  const msg = $('#scanmsg');
  try {
    stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1080 } }, audio: false });
    const v = $('#cam'); v.srcObject = stream; await v.play();
    if ('BarcodeDetector' in window) { try { const fm = await BarcodeDetector.getSupportedFormats(); if (fm.includes('data_matrix')) detector = new BarcodeDetector({ formats: ['data_matrix', 'code_128', 'qr_code'] }); } catch (e) { } }
    try {
      const tr = stream.getVideoTracks()[0], caps = tr.getCapabilities ? tr.getCapabilities() : {};
      if (caps.focusMode && caps.focusMode.includes('continuous')) await tr.applyConstraints({ advanced: [{ focusMode: 'continuous' }] });
      $('#bTorch').style.display = caps.torch ? '' : 'none';
    } catch (e) { }
    scanStart = Date.now(); scanN = 0;
    msg.textContent = 'Inquadra il Data Matrix dell\'etichetta (15–25 cm, senza riflessi)'; scanning = true; loopScan();
  } catch (e) { msg.textContent = 'Fotocamera non disponibile (permesso negato?). Usa "Da foto" o "+ A mano".'; }
}
function stopCam() { scanning = false; clearTimeout(scanTimer); if (stream) { stream.getTracks().forEach(t => t.stop()); stream = null; } }
let zxReady = false;
function zxInit() {
  if (zxReady || typeof ZXingWASM === 'undefined') return;
  zxReady = true;
  try { ZXingWASM.prepareZXingModule({ overrides: { locateFile: (f, pre) => f.endsWith('.wasm') ? 'lib/zxing_reader.wasm' : pre + f }, fireImmediately: true }); } catch (e) { }
}
const ZX_BASE = { formats: ['DataMatrix', 'Code128', 'QRCode'], tryHarder: true, tryRotate: true, tryInvert: true, tryDownscale: true, maxNumberOfSymbols: 4 };
async function zxRun(img, extra) {
  try { const r = await ZXingWASM.readBarcodes(img, Object.assign({}, ZX_BASE, extra || {})); if (r.length) return r.map(z => z.text).sort((a, b) => ((isAddr(b) ? 1000 : 0) + b.length) - ((isAddr(a) ? 1000 : 0) + a.length))[0]; } catch (e) { window.zxErr = String(e && e.message || e).slice(0, 60); }
  return null;
}
/* ritaglio (sx,sy,sw,sh) della sorgente, scalato in modo che il lato lungo sia maxSide (ingrandisce anche, fino a 'up' volte) */
function grab(src, sx, sy, sw, sh, maxSide, up, enhance) {
  const sc = Math.min(up || 1, maxSide / Math.max(sw, sh)); const w = Math.max(8, Math.round(sw * sc)), h = Math.max(8, Math.round(sh * sc));
  const c = document.createElement('canvas'); c.width = w; c.height = h; const x = c.getContext('2d', { willReadFrequently: true });
  x.imageSmoothingQuality = 'high'; x.drawImage(src, sx, sy, sw, sh, 0, 0, w, h);
  const id = x.getImageData(0, 0, w, h);
  if (enhance) {
    const d = id.data, n = w * h, hist = new Uint32Array(256);
    for (let i = 0; i < n; i++) { const y = (d[i * 4] * 77 + d[i * 4 + 1] * 150 + d[i * 4 + 2] * 29) >> 8; d[i * 4] = y; hist[y]++; }
    let lo = 0, hi = 255, acc = 0; while (lo < 254 && (acc += hist[lo]) < n * 0.02) lo++; acc = 0; while (hi > lo + 1 && (acc += hist[hi]) < n * 0.02) hi--;
    const k = 255 / Math.max(1, hi - lo);
    const g = new Uint8ClampedArray(n); for (let i = 0; i < n; i++) g[i] = (d[i * 4] - lo) * k;
    const o = new Uint8ClampedArray(n); /* nitidezza 3x3 */
    for (let yy = 1; yy < h - 1; yy++) for (let xx = 1; xx < w - 1; xx++) { const i = yy * w + xx; o[i] = 5 * g[i] - g[i - 1] - g[i + 1] - g[i - w] - g[i + w]; }
    for (let i = 0; i < n; i++) { const v = o[i] || g[i]; d[i * 4] = d[i * 4 + 1] = d[i * 4 + 2] = v; d[i * 4 + 3] = 255; }
  }
  return id;
}
/* tentativi, dal più veloce al più pesante */
function decodePasses(w, h) {
  const P = [
    s => zxRun(grab(s, 0, 0, w, h, 1280, 1)),
    s => zxRun(grab(s, 0, 0, w, h, 2200, 1)),
    s => zxRun(grab(s, 0, 0, w, h, 1920, 1, true)),
    s => zxRun(grab(s, w * 0.2, h * 0.2, w * 0.6, h * 0.6, 1600, 2.5)),
    s => zxRun(grab(s, w * 0.2, h * 0.2, w * 0.6, h * 0.6, 1600, 2.5, true), { binarizer: 'GlobalHistogram' }),
    s => zxRun(grab(s, 0, 0, w, h, 3000, 1), { binarizer: 'FixedThreshold' }),
    s => zxRun(grab(s, 0, 0, w, h, 800, 1)),
    s => zxRun(grab(s, 0, 0, w, h, 560, 1)),
    s => zxRun(grab(s, 0, 0, w, h, 960, 1, true), { binarizer: 'GlobalHistogram' }),
    s => zxRun(grab(s, w * 0.25, h * 0.25, w * 0.5, h * 0.5, 1400, 3, true)),
    s => zxRun(grab(s, w * 0.15, h * 0.15, w * 0.7, h * 0.7, 1000, 1.4, true), { binarizer: 'FixedThreshold' })
  ];
  /* finestre 3x3 sovrapposte */
  for (let r = 0; r < 3; r++) for (let c = 0; c < 3; c++) { const sx = w * c * 0.25, sy = h * r * 0.25; P.push(s => zxRun(grab(s, sx, sy, w * 0.5, h * 0.5, 1400, 2, r === 1))); }
  return P;
}
let passIdx = 0;
async function decodeSource(src, w, h, full) {
  let fb = null;
  if (detector) { try { const r = await detector.detect(src); if (r.length) { const t = r.map(z => z.rawValue).sort((a, b) => ((isAddr(b) ? 1000 : 0) + b.length) - ((isAddr(a) ? 1000 : 0) + a.length))[0]; if (isAddr(t)) return t; fb = t; } } catch (e) { } }
  zxInit(); if (typeof ZXingWASM === 'undefined') return fb;
  const P = decodePasses(w, h);
  /* cerco il Data Matrix con l'indirizzo (contiene '|'); un semplice codice a barre del tracking vale solo come ripiego */
  const run = async list => { for (const f of list) { const t = await f(src); if (t) { if (isAddr(t)) return t; if (!fb || t.length > fb.length) fb = t; } } return null; };
  if (full) { return (await run(P)) || fb; }
  passIdx = (passIdx + 1) % (P.length - 2);
  const q = await run([P[0], P[1], P[2 + passIdx]]); if (q) return q;
  if (fb) return (await run(P)) || fb; /* ho visto un codice a barre: c'è un'etichetta, cerco a fondo il Data Matrix */
  return null;
}
window.decodeFrame = (src, w, h) => decodeSource(src, w, h, true);
let shooting = false, lastShot = 0;
async function hiResShot() {
  if (!stream || typeof ImageCapture === 'undefined' || shooting) return null;
  shooting = true; lastShot = Date.now();
  try {
    const ic = new ImageCapture(stream.getVideoTracks()[0]); const opt = {};
    try { const c = await ic.getPhotoCapabilities(); if (c.imageWidth && c.imageWidth.max) opt.imageWidth = Math.min(c.imageWidth.max, 3264); } catch (e) { }
    const blob = await ic.takePhoto(opt); const bmp = await createImageBitmap(blob);
    const t = await decodeSource(bmp, bmp.width, bmp.height, true); if (bmp.close) bmp.close(); return t;
  } catch (e) { window.zxErr = 'foto: ' + String(e && e.message || e).slice(0, 40); return null; }
  finally { shooting = false; }
}
function scanDbg() {
  const v = $('#cam'); const eng = (detector ? 'ML Kit+' : '') + (typeof ZXingWASM === 'undefined' ? 'ZXing NON caricato' : 'ZXing');
  $('#scandbg').textContent = eng + ' · ' + (v.videoWidth || '?') + '×' + (v.videoHeight || '?') + ' · analisi ' + scanN + (window.zxErr ? ' · errore: ' + window.zxErr : '');
  if (scanning && scanN > 8 && Date.now() - scanStart > 9000) $('#scanmsg').textContent = 'Non leggo il codice. Avvicina a 15–25 cm, togli i riflessi, prova la 🔦 oppure usa 📸 Scatta e leggi.';
}
function loopScan() {
  if (!scanning) return;
  const v = $('#cam');
  if (v.videoWidth) {
    decodeSource(v, v.videoWidth, v.videoHeight).then(t => { scanN++; scanDbg(); if (!t || !scanning) return; if (isAddr(t)) { oneD = {}; onCode(t); } else { oneD[t] = (oneD[t] || 0) + 1;
      const n = oneD[t]; $('#scanmsg').textContent = n >= 8 ? 'Vedo solo il codice a barre/QR (senza indirizzo): avvicinati al quadrato Data Matrix, tienilo fermo e a fuoco' : 'Letto il tracking ' + t + ' – cerco il Data Matrix con l\'indirizzo…';
      if (n >= 2 && !shooting && Date.now() - lastShot > 4000) { $('#scanmsg').textContent = 'Etichetta vista: scatto una foto ad alta risoluzione…'; hiResShot().then(h => { if (h && scanning && isAddr(h)) { oneD = {}; onCode(h); } }); } } }).catch(e => { window.zxErr = String(e && e.message || e).slice(0, 60); }).finally(() => { if (scanning) scanTimer = setTimeout(loopScan, 150); });
  } else scanTimer = setTimeout(loopScan, 250);
}
function onCode(txt) {
  if (txt === lastCode && Date.now() - lastTime < 4000) return; lastCode = txt; lastTime = Date.now();
  if (navigator.vibrate) navigator.vibrate(70);
  const d = parsePayload(txt); const p = newParcel(d);
  if (d.track && DAY.parcels.some(x => x.track === d.track)) { toast('Già in lista: ' + d.track); return; }
  const a = DB.addr[akeyP(p)];
  if (fast && p.via) {
    DAY.parcels.push(p); if (a && a.lat != null) { p.lat = a.lat; p.lng = a.lng; p.geo = 'db'; } const m = matchPrio(p); p.prio = m.n; p.amb = m.amb; queueGeo(p);
    touch(); counters(); toast('✔ ' + (p.name || p.track) + ' – ' + fmtAddr(p) + (a && a.note ? '  📝 nota presente' : ''), 2200); return;
  }
  scanning = false; clearTimeout(scanTimer);
  p._cancel = () => { if (curView === 'scan') { scanning = true; loopScan(); } };
  parcelForm(p, () => { toast('Aggiunto: ' + (p.name || p.track)); if (curView === 'scan') { scanning = true; loopScan(); } }, true);
  if (!d.via) toast('Etichetta senza indirizzo: completa a mano');
}

/* ---------- lettura del testo dell'etichetta (OCR, funziona anche offline dopo il primo uso) ---------- */
let ocrW = null;
function loadScript(src) { return new Promise((res, rej) => { const e = document.createElement('script'); e.src = src; e.onload = res; e.onerror = () => rej(new Error('script ' + src)); document.head.appendChild(e); }); }
async function ocrWorker() {
  if (ocrW) return ocrW;
  if (typeof Tesseract === 'undefined') await loadScript('lib/ocr/tesseract.min.js');
  const base = new URL('lib/ocr/', location.href).href;
  ocrW = await Tesseract.createWorker('ita', 1, { workerPath: base + 'worker.min.js', corePath: base, langPath: base, cacheMethod: 'none' });
  return ocrW;
}
function rotCanvas(src, deg, maxSide) {
  const w0 = src.width, h0 = src.height, sc = Math.min(1, maxSide / Math.max(w0, h0)), w = Math.round(w0 * sc), h = Math.round(h0 * sc);
  const r = deg * Math.PI / 180, cs = Math.abs(Math.cos(r)), sn = Math.abs(Math.sin(r));
  const W = Math.ceil(w * cs + h * sn), H = Math.ceil(w * sn + h * cs);
  const c = document.createElement('canvas'); c.width = W; c.height = H; const x = c.getContext('2d');
  x.fillStyle = '#fff'; x.fillRect(0, 0, W, H); x.translate(W / 2, H / 2); x.rotate(r); x.filter = 'grayscale(1) contrast(1.35)'; x.drawImage(src, -w / 2, -h / 2, w, h);
  return c;
}
const VIA_RE = /^(via|viale|v\.le|corso|c\.so|piazza|p\.zza|piazzale|largo|vico|vicolo|strada|s\.p\.|sp|s\.s\.|contrada|c\.da|localit[àa]?|loc\.|frazione|fraz\.|traversa|rione|salita|discesa|borgo|parco)\b/i;
const tcase = s => String(s).toLowerCase().replace(/(^|[\s'.-])\p{L}/gu, m => m.toUpperCase()).replace(/ Di /g, ' di ').replace(/ Del /g, ' del ').replace(/ Della /g, ' della ');
function ocrScore(txt, conf) {
  let sc = conf || 0; if (/destinat/i.test(txt)) sc += 60; if (/\bvia\b|viale|corso|piazza|contrada|strada/i.test(txt)) sc += 25; if (/\b8\d{4}\b/.test(txt)) sc += 25; if (/spedizione|mittente|lettera di vettura/i.test(txt)) sc += 15; return sc;
}
function parseLabelText(text) {
  const lines = text.split(/\n/).map(raw => {
    let tk = raw.replace(/[|_=~<>«»]+/g, ' ').split(/\s+/).filter(x => x && !/^[^\p{L}\d]+$/u.test(x));
    const vi0 = tk.findIndex((x, i) => i < 5 && /^(\\?ia|via|viale|corso|piazza|piazzale|largo|vico|vicolo|strada|contrada|localit[àa]|frazione|traversa)$/i.test(x));
    if (vi0 >= 0) { tk = tk.slice(vi0); if (/^\\?ia$/i.test(tk[0])) tk[0] = 'Via'; }
    return tk.join(' ').replace(/^[^\p{L}\d]+/u, '').trim();
  }).filter(s => s.length > 1);
  const di = lines.findIndex(l => /destinat/i.test(l));
  const tail = lines.slice(di >= 0 ? di + 1 : 0, di >= 0 ? di + 9 : undefined);
  const vi = tail.findIndex(l => VIA_RE.test(l));
  const d = { track: '', name: '', via: '', civico: '', cap: '', comune: '', prov: '', tel: '', cod: '', tipo: 'consegna' };
  const capM = (tail.join(' ').match(/\b(8\d{4})\b/) || text.match(/\b(8\d{4})\b/)); if (capM) d.cap = capM[1];
  const t = text.replace(/(\d)\s(?=\d)/g, '$1'); const tel = t.match(/\b(3\d{9})\b/); if (tel) d.tel = tel[1];
  const trk = text.match(/\b(\d{3}[A-Z]\d{8,12}[A-Z]?|[A-Z0-9]{14,18})\b/); if (trk && /\d/.test(trk[1]) && /[A-Z]/.test(trk[1])) d.track = trk[1];
  if (vi < 0) return d;
  const addr = tail[vi].replace(/\s*[,;]\s*$/, '');
  const m = addr.match(/^(.*?)[,\s]+(?:n\.?|nr\.?|civ\.?)?\s*(\d+\s*[a-zA-Z]?(?:\/\w+)?|s\.?n\.?c\.?)\s*$/i);
  if (m) { d.via = m[1].replace(/[,\s]+$/, ''); d.civico = m[2].replace(/\s+/g, ''); } else d.via = addr;
  d.via = tcase(d.via);
  const nm = tail.slice(0, vi).reverse().map(l => l.replace(/\b\S*\d\S*\b/g, '').replace(/\s+/g, ' ').trim()).find(l => l.split(' ').filter(w => w.length > 2).length >= 1);
  if (nm) d.name = tcase(nm.split(' ').filter(w => w.length > 1).join(' '));
  const places = [];
  for (const l of tail.slice(vi + 1, vi + 5)) { const c = l.replace(/\([A-Z]{2}\)|\b[A-Z]{2}$/g, '').replace(/[^\p{L}' ]/gu, ' ').split(/\s+/).filter(w => w.length > 2 || /^(di|a|del|dei)$/i.test(w)).join(' ').trim(); if (c.length >= 4 && !/lettera|vettura|spedizione|importo|pagamento|peso|collit|contrassegno/i.test(c)) places.push(c); else if (places.length) break; }
  if (places.length) d.comune = tcase(places[Math.min(places.length, 2) - 1]); const frz = places.length >= 2 ? tcase(places[0]) : '';
  /* correzione dai nomi già noti (vie e comuni) */
  const sqv = sqStreet(d.via); if (sqv.length >= 4) { const hit = knownStreets().find(e => { const t = sqStreet(e.via); return t === sqv || lev(t, sqv) <= 1; }); if (hit) d.via = hit.via; }
  const coms = new Set(KNOWN_COM.map(baseComuneRaw)); Object.values(DB.addr).forEach(a => a.comune && coms.add(normCom(a.comune)));
  if (frz && coms.has(normCom(d.comune)) && !coms.has(normCom(frz))) d.comune = frz + ' di ' + d.comune;
  const nc = normCom(d.comune); if (nc && !/ di /.test(d.comune)) { const hit = [...coms].find(c => c === nc || (nc.length >= 5 && lev(c, nc) <= 1)); if (hit) d.comune = tcase(hit); }
  return d;
}
const baseComuneRaw = c => c;
/* direzione delle righe di testo/codici (mod 90°) dal profilo di proiezione dei pixel scuri */
function skewAngle(bmp) {
  const c = rotCanvas(bmp, 0, 480), x = c.getContext('2d', { willReadFrequently: true }), W = c.width, H = c.height, id = x.getImageData(0, 0, W, H).data;
  const g = new Uint8Array(W * H); let sum = 0, sq = 0;
  for (let i = 0; i < W * H; i++) { const v = id[i * 4]; g[i] = v; sum += v; sq += v * v; }
  const mean = sum / (W * H), sd = Math.sqrt(sq / (W * H) - mean * mean), th = mean - 0.7 * sd;
  const pts = []; for (let y = 0; y < H; y++) for (let xx = 0; xx < W; xx++) if (g[y * W + xx] < th) pts.push(xx, y);
  let best = 0, bv = -1; const R = Math.ceil(Math.hypot(W, H)) + 2;
  for (let a = -45; a < 45; a += 1) {
    const r = a * Math.PI / 180, cs = Math.cos(r), sn = Math.sin(r), h = new Float64Array(2 * R + 1);
    for (let i = 0; i < pts.length; i += 2) h[Math.round(pts[i + 1] * cs - pts[i] * sn) + R]++;
    let m = 0, q = 0, n = 0; for (let i = 0; i < h.length; i++) if (h[i]) { m += h[i]; q += h[i] * h[i]; n++; }
    const v = q / Math.max(1, n) - (m / Math.max(1, n)) ** 2 + 0 * m; /* varianza */
    let vv = 0; for (let i = 1; i < h.length; i++) vv += (h[i] - h[i - 1]) ** 2; /* righe nette = forti differenze tra righe vicine */
    if (vv > bv) { bv = vv; best = a; }
  }
  return best;
}
async function ocrLabel(file) {
  busy('Preparo la lettura (la prima volta può scaricare ~8 MB)…');
  const bmp = await createImageBitmap(file); const w = await ocrWorker();
  const sk = skewAngle(bmp); window.lastSkew = sk;
  const tries = [-sk, -sk + 180, -sk + 90, -sk + 270, -sk + 12, -sk - 12, -sk + 168, -sk + 192]; const found = [];
  for (let i = 0; i < tries.length; i++) {
    busy('Leggo l\'etichetta… ' + (i + 1) + '/' + tries.length);
    const o = await w.recognize(rotCanvas(bmp, tries[i], 1800)); const sc = ocrScore(o.data.text, o.data.confidence); const dd = parseLabelText(o.data.text); window.lastOcr = o.data.text;
    found.push({ sc, dd });
    if (/destinat/i.test(o.data.text) && dd.via) break;
  }
  found.sort((x, y) => y.sc - x.sc); if (bmp.close) bmp.close(); busy(); return found[0].dd;
}
function ocrToForm(d) {
  const p = newParcel(d); scanning = false; clearTimeout(scanTimer);
  p._cancel = () => { if (curView === 'scan') { scanning = true; loopScan(); } };
  parcelForm(p, () => { toast('Aggiunto: ' + (p.name || p.track || fmtAddr(p))); if (curView === 'scan') { scanning = true; loopScan(); } }, true);
  toast('Letto dal testo: controlla nome, via e civico prima di OK', 5000);
}
$('#bOcr').onclick = () => { shotPending = true; stopCam(); $('#fOcr').click(); };
$('#fOcr').addEventListener('change', async ev => {
  const f = ev.target.files[0]; ev.target.value = ''; shotPending = false;
  if (!f) { if (curView === 'scan') startCam(); return; }
  try {
    const d = await ocrLabel(f); busy();
    if (curView === 'scan') await startCam();
    if (d && (d.via || d.cap || d.name)) ocrToForm(d); else toast('Testo non riconosciuto: rifai la foto con l\'etichetta dritta, ben illuminata e a fuoco', 5000);
  } catch (e) { busy(); toast('Lettura testo non riuscita: ' + String(e && e.message || e).slice(0, 60) + (navigator.onLine ? '' : ' (la prima volta serve la rete)'), 5000); if (curView === 'scan') startCam(); }
});
$('#bFast').onclick = () => { fast = !fast; $('#bFast').textContent = 'Veloce: ' + (fast ? 'sì' : 'no'); };
$('#bPhoto').onclick = () => $('#fPhoto').click();
$('#bShot').onclick = () => { shotPending = true; stopCam(); $('#fShot').click(); };
$('#bTorch').onclick = async () => { try { torchOn = !torchOn; await stream.getVideoTracks()[0].applyConstraints({ advanced: [{ torch: torchOn }] }); $('#bTorch').textContent = torchOn ? '🔦 on' : '🔦'; } catch (e) { toast('Torcia non disponibile'); } };
let shotPending = false;
const photoHandler = async ev => {
  const f = ev.target.files[0]; ev.target.value = ''; shotPending = false; if (!f) { if (curView === 'scan') startCam(); return; }
  busy('Leggo la foto…');
  try {
    const bmp = await createImageBitmap(f); const t = await decodeSource(bmp, bmp.width, bmp.height, true); busy();
    if (curView === 'scan') await startCam();
    if (t && isAddr(t)) onCode(t); else if (t) toast('Ho letto solo un codice a barre/QR senza indirizzo: rifai la foto più vicina al quadrato Data Matrix', 4500); else toast('Codice non leggibile: rifai la foto più da vicino e a fuoco, oppure usa "+ A mano"', 4500);
  } catch (e) { busy(); toast('Foto non leggibile'); if (curView === 'scan') startCam(); }
};
$('#fPhoto').onchange = photoHandler; $('#fShot').onchange = photoHandler;
window.addEventListener('focus', () => { if (shotPending) setTimeout(() => { if (shotPending && curView === 'scan' && !stream) { shotPending = false; startCam(); } }, 1500); });

/* ---------- menu, import/export ---------- */
function renderMenu() {
  const nAddr = Object.keys(DB.addr).length;
  $('#v-menu').innerHTML = '<div class="menu">' +
    '<small>Database: ' + nAddr + ' indirizzi · ' + DB.prio.length + ' priorità · ' + DAY.parcels.length + ' pacchi oggi</small>' +
    '<button class="btn" id="m_imp">Importa / Unisci file (.xlsx)</button><small>Database, pacchi della giornata o entrambi. Se un indirizzo c\'è già vince la modifica più recente.</small>' +
    '<button class="btn sec" id="m_exp">Esporta database (.xlsx)</button>' +
    '<button class="btn sec" id="m_bk">Backup ora (.temp.xlsx)</button>' +
    '<button class="btn sec" id="m_ed">Editor indirizzi</button>' +
    '<button class="btn sec" id="m_pr">Editor priorità</button>' +
    '<button class="btn sec" id="m_tl">Scarica mappa Irpinia e Sannio (Wi‑Fi)</button>' +
    '<button class="btn sec" id="m_wk">Schermo sempre acceso: ' + (wake ? 'sì' : 'no') + '</button>' +
    '<button class="btn sec" id="m_nav">Navigazione con: ' + NAV_LABEL[NAV] + '</button>' +
    '<button class="btn sec" id="m_wa">Messaggio WhatsApp</button>' +
    '<button class="btn warn" id="m_fine">Fine lavoro</button>' +
    '<small>Versione ' + VER + '</small></div>';
  $('#m_imp').onclick = () => $('#fImport').click();
  $('#m_exp').onclick = () => { download('database-SDA-' + todayStr() + '.xlsx', xlsxBlob(dbWorkbook())); toast('Database esportato'); };
  $('#m_bk').onclick = () => backupNow();
  $('#m_ed').onclick = () => addrEditor('');
  $('#m_pr').onclick = prioEditor;
  $('#m_tl').onclick = downloadTiles;
  $('#m_wk').onclick = toggleWake;
  $('#m_nav').onclick = () => { NAV = { google: 'organic', organic: 'geo', geo: 'google' }[NAV]; lsSet('sda.nav', NAV); renderMenu(); toast('Il tasto Naviga ora usa: ' + NAV_LABEL[NAV], 3500); };
  $('#m_wa').onclick = waEditor;
  $('#m_fine').onclick = fineLavoro;
}
$('#fImport').onchange = async ev => {
  const f = ev.target.files[0]; ev.target.value = ''; if (!f) return;
  try {
    busy('Importo…'); const wb = XLSX.read(await f.arrayBuffer());
    let msg = [];
    const sh = n => { const k = wb.SheetNames.find(x => norm(x) === norm(n)); return k ? XLSX.utils.sheet_to_json(wb.Sheets[k], { defval: '' }) : null; };
    const ind = sh('Indirizzi'), pri = sh('Priorità') || sh('Priorita'), pac = sh('Pacchi');
    if (ind) msg.push(mergeAddrRows(ind) + ' indirizzi uniti');
    if (pri) msg.push(mergePrioRows(pri) + ' priorità');
    if (pac && pac.length) {
      busy(); if (DAY.parcels.length && !confirm('Sostituire i ' + DAY.parcels.length + ' pacchi già in lista con quelli del file?')) { pac.length = 0; }
      else { importParcels(pac); msg.push(pac.length + ' pacchi'); }
    }
    saveDB(); saveDay(); busy(); counters(); renderPins(); renderList(); renderMenu();
    toast(msg.join(' · ') || 'Nessun dato riconosciuto', 4000);
    if (pac && pac.length) { showView('map'); fitPins(); if (DAY.parcels.some(p => p.lat == null)) elabora(); }
  } catch (e) { busy(); toast('File non leggibile'); }
};
function importParcels(rows) {
  DAY.parcels = []; DAY.nextId = 1; DAY.date = todayStr();
  for (const r of rows) {
    const g = k => String(r[k] == null ? '' : r[k]).trim();
    const lat = parseFloat(g('lat').replace(',', '.')), lng = parseFloat(g('lng').replace(',', '.'));
    const p = newParcel({ track: g('tracking'), name: g('destinatario'), via: g('via'), civico: g('civico'), cap: g('cap'), comune: g('comune'), tel: g('tel'), cod: g('cod'), tipo: /rit/i.test(g('tipo')) ? 'ritiro' : 'consegna', prio: parseInt(g('prio'), 10) || 0, ord: parseInt(g('ordine'), 10) || 0, stato: g('stato') || 'da_fare' });
    if (isFinite(lat) && isFinite(lng)) { p.lat = lat; p.lng = lng; p.geo = g('geo') || 'db'; }
    if (g('nota')) { const k = akeyP(p); const cur = DB.addr[k]; if (!cur || !cur.note) { DB.addr[k] = Object.assign(cur || { via: p.via, civico: p.civico, cap: p.cap, comune: p.comune, lat: p.lat, lng: p.lng, prio: 0 }, { note: g('nota'), mod: nowIso() }); } }
    DAY.parcels.push(p);
  }
  for (const p of DAY.parcels) { if (p.lat == null) { const f = lookupDB(p); if (f) Object.assign(p, f); } if (!p.prio) { const m = matchPrio(p); p.prio = m.n; p.amb = m.amb; } }
  DAY.lastBackup = Date.now(); DAY.doneSince = 0; DAY.dirty = 0;
}
function addrEditor(q) {
  const rows = Object.entries(DB.addr).filter(([k, a]) => !q || norm(a.via + ' ' + a.comune + ' ' + a.civico).includes(norm(q))).sort((a, b) => normCom(a[1].comune).localeCompare(normCom(b[1].comune)) || normStreet(a[1].via).localeCompare(normStreet(b[1].via)) || ((parseInt(a[1].civico) || 0) - (parseInt(b[1].civico) || 0))).slice(0, 80);
  sheet('<h3>Indirizzi nel database</h3><input id="aq" placeholder="Cerca via o comune" value="' + esc(q) + '"><div class="list" style="margin-top:6px;max-height:50vh;overflow:auto">' +
    rows.map(([k, a]) => '<div class="it" data-k="' + esc(k) + '"><div class="x"><b>' + esc((a.via + ' ' + a.civico).trim()) + '</b><span>' + esc(a.comune) + (a.prio ? ' · prio ' + a.prio : '') + (a.note ? ' · 📝' : '') + (a.lat == null ? ' · senza posizione' : '') + '</span></div></div>').join('') + '</div>' +
    '<div class="row"><button class="btn sec" id="ac">Chiudi</button><button class="btn" id="an">+ Nuovo</button></div>');
  $('#aq').onchange = () => addrEditor($('#aq').value);
  $('#ac').onclick = closeSheet; $('#an').onclick = () => addrForm(null, q);
  $('#sheet').querySelectorAll('[data-k]').forEach(it => it.onclick = () => addrForm(it.dataset.k, q));
}
function addrForm(key, q) {
  const a = key ? DB.addr[key] : { via: '', civico: '', cap: '', comune: '', lat: '', lng: '', note: '', prio: '' };
  sheet('<h3>' + (key ? 'Modifica indirizzo' : 'Nuovo indirizzo') + '</h3><div class="row"><div><label>Via</label><input id="e_via" value="' + esc(a.via) + '"></div><div style="flex:.5"><label>Civico</label><input id="e_civ" value="' + esc(a.civico) + '"></div></div>' +
    '<div class="row"><div style="flex:.6"><label>CAP</label><input id="e_cap" value="' + esc(a.cap) + '"></div><div><label>Città</label><input id="e_com" value="' + esc(a.comune) + '"></div></div>' +
    '<div class="row"><div><label>Latitudine</label><input id="e_lat" value="' + esc(a.lat == null ? '' : a.lat) + '" inputmode="decimal"></div><div><label>Longitudine</label><input id="e_lng" value="' + esc(a.lng == null ? '' : a.lng) + '" inputmode="decimal"></div></div>' +
    '<div class="row"><div style="flex:.5"><label>Priorità</label><input id="e_pr" type="number" value="' + esc(a.prio || '') + '"></div></div><label>Note</label><textarea id="e_nt">' + esc(a.note) + '</textarea>' +
    '<div class="row"><button class="btn sec" id="e_c">Indietro</button>' + (key ? '<button class="btn no" id="e_d">Elimina</button>' : '') + '<button class="btn ok" id="e_o">Salva</button></div>');
  $('#e_c').onclick = () => addrEditor(q);
  if (key) $('#e_d').onclick = () => { if (confirm('Eliminare questo indirizzo dal database?')) { delete DB.addr[key]; saveDB(); snapshot(); addrEditor(q); } };
  $('#e_o').onclick = () => {
    const via = $('#e_via').value.trim(); if (!via) return toast('Manca la via');
    const la = parseFloat($('#e_lat').value.replace(',', '.')), lo = parseFloat($('#e_lng').value.replace(',', '.'));
    const na = { via, civico: $('#e_civ').value.trim(), cap: $('#e_cap').value.trim(), comune: $('#e_com').value.trim(), lat: isFinite(la) ? la : null, lng: isFinite(lo) ? lo : null, note: $('#e_nt').value.trim(), prio: parseInt($('#e_pr').value, 10) || 0, mod: nowIso() };
    if (key) delete DB.addr[key]; DB.addr[akey(na.via, na.civico, na.comune)] = na; saveDB(); snapshot(); addrEditor(q);
  };
}
function prioEditor() {
  sheet('<h3>Priorità del giro</h3><div class="list" style="max-height:55vh;overflow:auto">' + DB.prio.map(e => '<div class="it" data-n="' + e.n + '"><div class="n st-c">' + e.n + '</div><div class="x"><b>' + esc(e.zona) + '</b><span>' + esc(e.testo) + '</span></div></div>').join('') + '</div><div class="row"><button class="btn sec" id="pc">Chiudi</button><button class="btn" id="pn">+ Nuova</button></div>');
  $('#pc').onclick = closeSheet; $('#pn').onclick = () => prioForm(null);
  $('#sheet').querySelectorAll('[data-n]').forEach(it => it.onclick = () => prioForm(+it.dataset.n));
}
function prioForm(n) {
  const e = n ? DB.prio.find(x => x.n === n) : { n: (Math.max(0, ...DB.prio.map(x => x.n)) + 1), zona: '', com: '', loc: '', testo: '', keys: [] };
  sheet('<h3>Priorità ' + e.n + '</h3><label>Numero</label><input id="p_n" type="number" value="' + e.n + '"><label>Zona</label><input id="p_z" value="' + esc(e.zona) + '"><label>Comune (arpaise, ceppaloni, san leucio, apollosa, roccabascerana)</label><input id="p_c" value="' + esc(e.com) + '">' +
    '<label>Descrizione</label><textarea id="p_t">' + esc(e.testo) + '</textarea><label>Nomi di vie (separati da ;)  – es. roma@1-33 = solo civici 1–33</label><textarea id="p_k">' + esc((e.keys || []).join('; ')) + '</textarea>' +
    '<div class="row"><button class="btn sec" id="p_b">Indietro</button>' + (n ? '<button class="btn no" id="p_d">Elimina</button>' : '') + '<button class="btn ok" id="p_o">Salva</button></div>');
  $('#p_b').onclick = prioEditor;
  if (n) $('#p_d').onclick = () => { DB.prio = DB.prio.filter(x => x.n !== n); saveDB(); prioEditor(); };
  $('#p_o').onclick = () => {
    const ne = { n: parseInt($('#p_n').value, 10) || e.n, zona: $('#p_z').value.trim(), com: norm($('#p_c').value), loc: e.loc || '', testo: $('#p_t').value.trim(), keys: $('#p_k').value.split(';').map(s => s.trim()).filter(Boolean) };
    DB.prio = DB.prio.filter(x => x.n !== n && x.n !== ne.n); DB.prio.push(ne); DB.prio.sort((a, b) => a.n - b.n); saveDB(); prioEditor();
  };
}
function waEditor() {
  sheet('<h3>Messaggio WhatsApp</h3><textarea id="wt">' + esc(WA) + '</textarea><div class="row"><button class="btn sec" id="wc">Annulla</button><button class="btn ok" id="wo">Salva</button></div>');
  $('#wc').onclick = closeSheet; $('#wo').onclick = () => { WA = $('#wt').value; lsSet(K.wa, WA); closeSheet(); };
}
let wake = null;
async function toggleWake() {
  try { if (wake) { await wake.release(); wake = null; } else { wake = await navigator.wakeLock.request('screen'); } } catch (e) { toast('Non supportato su questo telefono'); }
  renderMenu();
}
document.addEventListener('visibilitychange', async () => { if (document.visibilityState === 'visible' && wake) { try { wake = await navigator.wakeLock.request('screen'); } catch (e) { } } });

/* ---------- fine lavoro ---------- */
function fineLavoro() {
  const c = DAY.parcels.filter(p => p.tipo === 'consegna'), r = DAY.parcels.filter(p => p.tipo === 'ritiro');
  const left = DAY.parcels.filter(p => p.stato === 'da_fare').length;
  sheet('<h3>Fine lavoro</h3><p>Consegne: ' + c.filter(p => p.stato === 'consegnato').length + '/' + c.length + ' · Ritiri: ' + r.filter(p => p.stato === 'ritirato').length + '/' + r.length + (left ? '<br><b>Attenzione: ' + left + ' ancora da fare.</b>' : '') + '</p><p style="font-size:14px;color:#51627d">Il database aggiornato viene salvato come file .xlsx (cartella Download).</p><div class="row"><button class="btn sec" id="fc">Annulla</button><button class="btn ok" id="fo">Salva database</button></div>');
  $('#fc').onclick = closeSheet;
  $('#fo').onclick = () => {
    download('database-SDA-' + todayStr() + '.xlsx', xlsxBlob(dbWorkbook())); DAY.lastBackup = Date.now(); DAY.dirty = 0; DAY.doneSince = 0; saveDay();
    const list = SNAPS.map(s => '<div class="card" data-s="' + esc(s.name) + '"><div class="t" style="font-size:14px">' + esc(s.name) + '</div></div>').join('') || '<div class="s">Nessuna copia temporanea</div>';
    sheet('<h3>Database salvato ✔</h3><p style="font-size:14px">Copie temporanee presenti sul telefono:</p><div style="max-height:30vh;overflow:auto">' + list + '</div><p><b>Vuoi cancellare la cache?</b><br><span style="font-size:13px;color:#51627d">Cancella copie temporanee, pacchi di oggi e database dal telefono. Controlla che il file sia in Download.</span></p><div class="row"><button class="btn sec" id="cn">No</button><button class="btn no" id="cs">Sì</button></div>');
    $('#sheet').querySelectorAll('[data-s]').forEach(c => c.onclick = () => { const s = SNAPS.find(x => x.name === c.dataset.s); if (!s) return; const keep = DB.addr; const tmp = {}; s.addr.forEach(a => tmp[akey(a.via, a.civico, a.comune)] = a); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(Object.entries(tmp).map(([k, a]) => ({ chiave: k, via: a.via, civico: a.civico, cap: a.cap || '', comune: a.comune, lat: a.lat, lng: a.lng, note: a.note || '', priorita: a.prio || '', modificato: a.mod || '' }))), 'Indirizzi'); download(s.name, xlsxBlob(wb)); });
    $('#cn').onclick = closeSheet;
    $('#cs').onclick = () => { localStorage.removeItem(K.db); localStorage.removeItem(K.day); localStorage.removeItem(K.snaps); DB = { addr: {}, prio: JSON.parse(JSON.stringify(DEFAULT_PRIO)) }; DAY = { date: todayStr(), parcels: [], nextId: 1, lastBackup: 0, doneSince: 0, dirty: 0, snooze: 0 }; SNAPS = []; closeSheet(); renderPins(); counters(); renderMenu(); toast('Cache cancellata'); };
  };
}

/* ---------- mappa offline: scarica la zona del giro ---------- */
function lon2x(lon, z) { return Math.floor((lon + 180) / 360 * Math.pow(2, z)); }
function lat2y(lat, z) { const r = lat * Math.PI / 180; return Math.floor((1 - Math.log(Math.tan(r) + 1 / Math.cos(r)) / Math.PI) / 2 * Math.pow(2, z)); }
let dlStop = false;
$('#busy').addEventListener('click', e => { if (e.target && e.target.id === 'dlStop') dlStop = true; });
async function downloadTiles() {
  if (!navigator.onLine) return toast('Serve la rete (meglio Wi‑Fi)');
  const boxes = [
    { n: 41.5, s: 40.7, w: 14.3, e: 15.4, z0: 9, z1: 13 },          // tutta Irpinia e Sannio: paesi e strade principali
    { n: 41.2, s: 40.9, w: 14.5, e: 14.9, z0: 14, z1: 15 },         // zona ampia del giro, da Atripalda alla Valle Caudina
    { n: 41.16, s: 40.96, w: 14.58, e: 14.86, z0: 16, z1: 17 }      // i tuoi paesi con tutte le stradine: Arpaise, Ceppaloni, San Leucio, Apollosa, Tufara, Roccabascerana, Pannarano
  ];
  const set = new Set();
  for (const bx of boxes) for (let z = bx.z0; z <= bx.z1; z++) {
    const x0 = lon2x(bx.w, z), x1 = lon2x(bx.e, z), y0 = lat2y(bx.n, z), y1 = lat2y(bx.s, z);
    for (let x = x0; x <= x1; x++) for (let y = y0; y <= y1; y++) set.add(z + '/' + x + '/' + y);
  }
  // corridoio di attraversamento da Atripalda (nomi delle vie visibili anche se ti chiedono di portare il pacco altrove)
  const extra = [{ n: 40.98, s: 40.89, w: 14.76, e: 14.88 }];
  // tutti i comuni dove hai pacchi o indirizzi nel database: area intorno ai loro punti, livello 16-17
  const byCom = {};
  const addPt = (com, la, lo) => { const k = normCom(com) || '?'; const b = byCom[k] || (byCom[k] = { n: -90, s: 90, w: 180, e: -180 }); b.n = Math.max(b.n, la); b.s = Math.min(b.s, la); b.w = Math.min(b.w, lo); b.e = Math.max(b.e, lo); };
  DAY.parcels.filter(p => p.lat != null).forEach(p => addPt(p.comune, p.lat, p.lng));
  Object.values(DB.addr).filter(a => a.lat != null).forEach(a => addPt(a.comune, a.lat, a.lng));
  for (const k in byCom) { const b = byCom[k], pad = 0.012; extra.push({ n: b.n + pad, s: b.s - pad, w: b.w - pad, e: b.e + pad }); }
  for (const bx of extra) for (const z of [16, 17]) {
    const x0 = lon2x(bx.w, z), x1 = lon2x(bx.e, z), y0 = lat2y(bx.n, z), y1 = lat2y(bx.s, z);
    for (let x = x0; x <= x1; x++) for (let y = y0; y <= y1; y++) set.add(z + '/' + x + '/' + y);
  }
  const c = await caches.open('tiles-v1');
  const todo = [];
  for (const t of set) { const url = 'https://tile.openstreetmap.org/' + t + '.png'; if (!(await c.match(url))) todo.push(url); }
  if (!todo.length) return toast('Mappa già scaricata');
  if (!confirm('Scarico ' + todo.length + ' tessere (circa ' + Math.round(todo.length * 0.015) + ' MB, ' + Math.round(todo.length / 14 / 60) + '–' + Math.round(todo.length / 8 / 60) + ' minuti): tutta Irpinia e Sannio e, per i tuoi paesi e per la strada da Atripalda, tutte le stradine. Serve il Wi‑Fi e la batteria carica. Se ti interrompi, riprende da dove era arrivato. Procedo?')) return;
  const total = todo.length; let done = 0, fail = 0; dlStop = false;
  const msg = () => 'Scarico la mappa… ' + Math.round(done / total * 100) + '%<br><small>' + done + ' / ' + total + ' · errori ' + fail + '</small><br><button class="btn warn sm" id="dlStop" style="margin-top:16px">Interrompi</button>';
  async function worker() {
    while (todo.length && !dlStop) {
      const url = todo.pop();
      try { const r = await fetch(url, { mode: 'cors' }); if (r.ok) await c.put(url, r); else fail++; } catch (e) { fail++; }
      await sleep(90);
      done++; if (done % 15 === 0) busy(msg());
    }
  }
  busy(msg()); await Promise.all([worker(), worker(), worker()]); busy();
  toast(dlStop ? 'Interrotto: premi di nuovo il tasto per continuare' : fail > total * 0.05 ? 'Finito con ' + fail + ' errori: riprova più tardi per completare' : 'Mappa scaricata ✔', 6000);
}

/* ---------- avvio ---------- */
function netState() { $('#net').className = navigator.onLine ? '' : 'off'; setNative(); }
window.addEventListener('online', netState); window.addEventListener('offline', netState); netState();
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => { });
if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => { });
counters(); renderPins(); startGPS(); checkRemind();
if (DAY.parcels.some(p => p.lat != null)) fitPins();
