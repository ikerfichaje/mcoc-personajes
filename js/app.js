/* MCOC Buscador — app sin dependencias. Datos: data/data.js (generado desde el Excel). */
(() => {
  'use strict';

  const RAW = self.MCOC_RAW;
  const IMG = self.MCOC_IMG || {};
  const view = document.getElementById('view');
  const titleEl = document.getElementById('title');
  const backBtn = document.getElementById('backBtn');

  // ---------- utilidades ----------
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  const none = v => !v || v === '—' || v === '-';
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* sin almacenamiento */ } }
  };
  // separa por `sep` ignorando lo que va entre paréntesis
  function splitTop(str, sep) {
    if (none(str)) return [];
    const out = []; let depth = 0, cur = '';
    for (let i = 0; i < str.length; i++) {
      const ch = str[i];
      if (ch === '(') depth++;
      else if (ch === ')') depth = Math.max(0, depth - 1);
      if (depth === 0 && str.startsWith(sep, i)) { out.push(cur); cur = ''; i += sep.length - 1; continue; }
      cur += ch;
    }
    out.push(cur);
    return out.map(s => s.trim()).filter(s => s && !none(s));
  }
  const list = v => splitTop(v, ', ');
  const groups = v => splitTop(v, '; ');
  const items = v => groups(v).flatMap(list);   // listas que a veces usan «; » en vez de «, »
  const synLabel = s => (s.imm ? `${s.e} (inmunidad)` : s.e);
  // «Efecto (dónde)» de las inmunidades condicionales. Algunos efectos llevan paréntesis en el propio nombre
  // («Atravesar (Phasing)», «todos los perjuicios (debuffs)»…): solo el último paréntesis es la condición, y
  // si el efecto es uno de esos, se conserva entero.
  const NAME_PAREN = /^(debuffs|prowess|taunt|miss|daunted|ambush|death|perjuicio|no puede recibir buffs)$/i;
  function splitCond(s) {
    // busca el paréntesis que abre el último grupo, aunque dentro haya otros paréntesis
    // («Incinerar (Winchesters Hellfire Rounds (Medium Combo Ender…))»)
    let m = null;
    if (s.endsWith(')')) {
      let depth = 0;
      for (let i = s.length - 1; i >= 0; i--) {
        if (s[i] === ')') depth++;
        else if (s[i] === '(' && --depth === 0) { if (s.slice(0, i).trim()) m = [s, s.slice(0, i).trim(), s.slice(i + 1, -1)]; break; }
      }
    }
    if (!m) return { e: s, w: '' };
    if (NAME_PAREN.test(m[2]) || (/^phasing$/i.test(m[2]) && /^atravesar$/i.test(m[1]))) return { e: s, w: '' };
    return { e: m[1], w: m[2] };
  }

  function rowsOf(name) {
    const sh = RAW.sheets[name];
    const h = sh.rows[0];
    return sh.rows.slice(1).map((r, i) => {
      const o = { _i: i + 1 };
      h.forEach((k, j) => { o[k] = r[j] == null ? '' : r[j]; });
      return o;
    });
  }

  const CLASS_KEY = { 'Cósmico': 'cosmico', 'Tecnología': 'tecnologia', 'Mutante': 'mutante', 'Habilidad': 'habilidad', 'Ciencia': 'ciencia', 'Místico': 'mistico' };
  const CLASSES = Object.keys(CLASS_KEY);
  const ccVar = cls => `--cc:var(--c-${CLASS_KEY[cls] || 'x'})`;
  const RANK = { 'S+': 0, 'S': 1, 'A': 2, 'B': 3, 'C': 4, 'D': 5, 'E': 6, 'F': 6 };
  const rank = t => (t in RANK ? RANK[t] : 9);
  const tierHtml = (t, label) => none(t)
    ? `<span class="tier" title="Sin valorar">${label ? `<small>${label}</small>` : ''}–</span>`
    : `<span class="tier t-${t === 'S+' ? 'Sp' : t}">${label ? `<small>${label}</small>` : ''}${esc(t)}</span>`;

  // ---------- modelo de datos ----------
  const champs = [];
  const byName = new Map();
  {
    const links = RAW.sheets['Campeones'].links || {};
    for (const r of rowsOf('Campeones')) {
      if (!r['Campeón']) continue;
      const c = {
        name: r['Campeón'], cls: r['Clase'], rar: r['Rareza máx.'],
        yt: r['Tier youtubers'], atk: r['Guide Blog ataque'], def: r['Guide Blog defensa'],
        immP: items(r['Inmunidades permanentes']),
        immC: items(r['Inmunidades condicionales']).map(splitCond),
        immU: items(r['Inmunidades no confirmadas']),
        // «20% menos daño de X», «Resistencia a X», «60% de resistencia a X», «resistencia parcial a X» -> efecto X
        res: groups(r['Resistencias pasivas']).flatMap(list).map(s => {
          const m = s.match(/menos daño de\s+(.+)$/i) || s.match(/^(?:\d+\s*%\s*de\s+)?resistencia(?:\s+parcial)?\s+a\s+(.+?)(\s+\d.*)?$/i);
          return { e: m ? m[1] : s, txt: s };
        }),
        prov: items(r['Efectos que provoca al rival']),
        gain: items(r['Mejoras/efectos que obtiene']),
        other: items(r['Otras mecánicas']),
        pas: groups(r['Efectos pasivos (siempre activos)']).map(g => {
          const k = g.indexOf(': ');
          return k > 0 ? { t: g.slice(0, k), items: list(g.slice(k + 2)) } : { t: 'Pasivo', items: list(g) };
        }),
        // «Veneno (inmunidad) con X» -> efecto «Veneno», marcado como inmunidad que da la sinergia
        syn: groups(r['Por sinergia']).map(g => {
          const k = g.indexOf(' con ');
          const raw = k > 0 ? g.slice(0, k) : g;
          const imm = /\s*\(inmunidad\)$/i.test(raw);
          return { e: raw.replace(/\s*\(inmunidad\)$/i, ''), imm, with: k > 0 ? list(g.slice(k + 5)) : [] };
        }),
        key: items(r['Habilidades clave (fichas parciales)']),
        ficha: r['Ficha'],
        url: links[`${r._i},17`] || (/^https?:/.test(r['Enlace texto completo']) ? r['Enlace texto completo'] : ''),
        kit: [], bc: '', scP: '', scN: '', adv: '',
        gen: '', genScore: null, genPos: null, nOps: 0, agree: '', genDef: '', ops: {}
      };
      champs.push(c);
      byName.set(c.name, c);
    }
    // Tier list general (consenso de varias fuentes; hoja generada por scripts/update-tiers.ps1)
    if (RAW.sheets['Tier list general']) {
      for (const r of rowsOf('Tier list general')) {
        const c = byName.get(r['Campeón']);
        if (!c) continue;
        const num = v => (none(v) || isNaN(+v) ? null : +v);
        c.gen = none(r['Tier general']) ? '' : r['Tier general'];
        c.genScore = num(r['Nota general (0-100)']);
        c.genPos = num(r['Puesto']);
        c.nOps = num(r['Nº opiniones']) || 0;
        c.agree = none(r['Acuerdo']) ? '' : r['Acuerdo'];
        c.genDef = none(r['Tier general defensa']) ? '' : r['Tier general defensa'];
        c.ops = {
          Vega: num(r['Vega']), Lagacy: num(r['Lagacy']), MetalSonicDude: num(r['MetalSonicDude']), Seatin: num(r['Seatin']),
          scAtk: none(r['MCOC Scout ataque']) ? '' : r['MCOC Scout ataque'], scDef: none(r['MCOC Scout defensa']) ? '' : r['MCOC Scout defensa']
        };
      }
    }
    for (const r of rowsOf('Mapa de habilidades')) {
      const c = byName.get(r['Campeón']);
      if (c) c.kit.push({ part: r['Parte del kit'], pas: r['¿Pasiva?'] === 'Sí', eff: list(r['Efectos que aparecen']) });
    }
    for (const r of rowsOf('Reliquias por campeón')) {
      const c = byName.get(r['Campeón']);
      if (!c) continue;
      c.bc = none(r['Battlecast recomendado']) ? '' : r['Battlecast recomendado'];
      c.scP = none(r['Statcast recomendado (retrato del relic)']) ? '' : r['Statcast recomendado (retrato del relic)'];
      c.scN = none(r['Statcast recomendado (nombre exacto)']) ? '' : r['Statcast recomendado (nombre exacto)'];
      c.adv = none(r['¿Cuál te conviene más?']) ? '' : r['¿Cuál te conviene más?'];
    }
  }
  // Efectos leídos directamente del texto del kit (scripts/kit-effects.ps1): incluyen los propios de cada
  // campeón y corrigen la dirección (provoca / obtiene) de lo que venía del Excel.
  {
    const KITFX = self.MCOC_KITFX || {};
    const has = (arr, e) => arr.some(x => norm(x) === norm(e));
    for (const c of champs) {
      const k = c.url ? KITFX[c.url.split('/').pop()] : null;
      if (!k) continue;
      const prov = (k.prov || []).slice(), gain = (k.gain || []).slice();
      c.prov.forEach(e => { if (!has(prov, e) && !has(gain, e)) prov.push(e); });
      c.gain.forEach(e => { if (!has(gain, e) && !has(prov, e)) gain.push(e); });
      c.prov = prov; c.gain = gain;
      c.other = c.other.filter(e => !has(prov, e) && !has(gain, e));
    }
  }
  const nameIndex = new Map(champs.map(c => [norm(c.name), c]));
  const findChamp = n => byName.get(n) || nameIndex.get(norm(n));

  // Battlecasts
  const bcBase = s => String(s || '').replace(/\s*\(.*\)\s*$/, '').trim();
  const battlecasts = rowsOf('Battlecasts').filter(r => r['Battlecast']).map(r => ({
    name: r['Battlecast'], key: bcBase(r['Battlecast']), cls: r['Clase'],
    eff: none(r['Efectos conocidos']) || /^Pendiente/i.test(r['Efectos conocidos']) ? '' : r['Efectos conocidos']
  }));
  const bcByKey = new Map(battlecasts.map(b => [norm(b.key), b]));
  const bcOf = name => bcByKey.get(norm(bcBase(name)));
  const scKey = c => c.scP || c.scN;

  // Statcasts: familias (prefijo + sufijos) por clase
  const scFamilies = [];
  for (const r of RAW.sheets['Statcasts'].rows.slice(1)) {
    if (!r[0] || !CLASS_KEY[r[0]] || !r[1] || /Retrato/.test(r[1])) { if (r[0] && /^Retratos/.test(r[0])) break; continue; }
    scFamilies.push({ cls: r[0], prefix: r[1], pattern: r[2], suffixes: list(r[3]), portrait: none(r[4]) ? '' : r[4] });
  }
  const statcasts = new Map(); // key -> {key, cls, champs[]}
  for (const c of champs) {
    const k = scKey(c);
    if (!k) continue;
    if (!statcasts.has(k)) statcasts.set(k, { key: k, cls: c.cls, exact: !c.scP, champs: [] });
    statcasts.get(k).champs.push(c);
  }

  // Matriz de inmunidades (para la vista Equipo)
  const matrix = RAW.sheets['Matriz de inmunidades'].rows;
  const MATRIX_EFFECTS = matrix[0].slice(3);
  const immMatrix = new Map();
  for (const r of matrix.slice(1)) {
    if (!byName.has(r[0])) continue;
    const m = {};
    MATRIX_EFFECTS.forEach((e, j) => { const v = r[j + 3]; if (v) m[e] = v; });
    immMatrix.set(r[0], m);
  }

  // Índice de efectos para el buscador
  const CAT = {
    imm: 'inmune', immc: 'inmune (condicional)', immu: 'inmune (sin confirmar)', res: 'resiste',
    prov: 'provoca', gain: 'obtiene', other: 'mecánica', pas: 'pasivo', syn: 'sinergia', kit: 'en el kit', key: 'habilidad clave'
  };
  const MODES = [
    ['any', 'Cualquier parte', null],
    ['imm', 'Inmune (siempre)', ['imm']],
    ['immc', 'Inmune (cualquiera)', ['imm', 'immc', 'immu']],
    ['prov', 'Lo provoca al rival', ['prov']],
    ['gain', 'Lo obtiene (mejora)', ['gain']],
    ['pas', 'Pasivo', ['pas']],
    ['res', 'Resiste su daño', ['res']]
  ];
  const MODE_CATS = Object.fromEntries(MODES.map(m => [m[0], m[2]]));
  const effects = new Map(); // normKey -> {label, champs:Set}
  for (const c of champs) {
    c.tags = new Map(); // normKey -> Set(cat)
    const add = (e, cat) => {
      const k = norm(e);
      if (!k) return;
      if (!effects.has(k)) effects.set(k, { key: k, label: e, champs: new Set() });
      effects.get(k).champs.add(c);
      if (!c.tags.has(k)) c.tags.set(k, new Set());
      c.tags.get(k).add(cat);
    };
    c.immP.forEach(e => add(e, 'imm'));
    c.immC.forEach(x => add(x.e, 'immc'));
    c.immU.forEach(e => add(e, 'immu'));
    c.res.forEach(x => add(x.e, 'res'));
    c.prov.forEach(e => add(e, 'prov'));
    c.gain.forEach(e => add(e, 'gain'));
    c.other.forEach(e => add(e, 'other'));
    c.pas.forEach(g => g.items.forEach(e => add(e, 'pas')));
    c.syn.forEach(s => add(s.e, 'syn'));
    c.kit.forEach(k => k.eff.forEach(e => add(e, 'kit')));
    c.key.forEach(e => add(e, 'key'));
  }
  const effectLabel = k => (effects.get(k) || { label: k }).label;
  self.MCOC_DEBUG = { effects, champs };   // para revisar los datos desde la consola del navegador

  // ---------- estado ----------
  const S = Object.assign({ q: '', filters: [], cls: [], rar: '', sort: 'yt', relSeg: 'bc', bestUse: 'atk', bestCls: '', best7: false, rkCls: 'Todas', rkBy: 'gen', rk7: false },
    store.get('mcoc-state', {}));
  S.q = '';
  let team = store.get('mcoc-team', []).filter(n => byName.has(n));
  const saveState = () => store.set('mcoc-state', { filters: S.filters, cls: S.cls, rar: S.rar, sort: S.sort, relSeg: S.relSeg, bestUse: S.bestUse, bestCls: S.bestCls, best7: S.best7, rkCls: S.rkCls, rkBy: S.rkBy, rk7: S.rk7, teamSeg: S.teamSeg, recUse: S.recUse, recSize: S.recSize, top7: S.top7 });
  const saveTeam = () => store.set('mcoc-team', team);

  // ---------- piezas de interfaz ----------
  function avatar(c, size = '') {
    const ini = c.name.replace(/\(.*\)/, '').split(/[\s-]+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase();
    const f = IMG[c.name];
    return `<span class="av ${size}" style="${ccVar(c.cls)}">${esc(ini)}${f ? `<img loading="lazy" decoding="async" src="img/${esc(f)}" alt="" onerror="this.remove()">` : ''}</span>`;
  }
  const clsTag = cls => `<span class="cls" style="${ccVar(cls)}">${esc(cls)}</span>`;
  const champHref = c => `#/c/${encodeURIComponent(c.name)}`;
  const tiersMini = c => `<span class="tiers" title="Tier general (consenso)">${tierHtml(c.gen || c.yt)}</span>`;
  const AGREE_TXT = { Alto: 'las fuentes coinciden', Medio: 'algo de discrepancia', Bajo: 'opiniones muy distintas' };
  // nota 0-100 -> letra, con las franjas de mcoc.app (las mismas que usa la hoja «Tier list general»)
  const scoreTier = v => v == null ? '' : v >= 96 ? 'S+' : v >= 86 ? 'S' : v >= 72 ? 'A' : v >= 60 ? 'B' : v >= 46 ? 'C' : v >= 30 ? 'D' : 'F';
  function opinionsHtml(c) {
    const rows = [
      ['Vega', 'YouTube (vía mcoc.app)', c.ops.Vega], ['Lagacy', 'YouTube (vía mcoc.app)', c.ops.Lagacy],
      ['MetalSonicDude', 'YouTube (vía mcoc.app)', c.ops.MetalSonicDude], ['Seatin', 'YouTube (vía mcoc.app)', c.ops.Seatin]
    ].map(([n, src, v]) => `<tr><th>${n}<br><small>${src}</small></th><td>${v == null ? '<span class="empty">sin valorar</span>' : `${tierHtml(scoreTier(v))} <small>${v}/100</small>`}</td></tr>`);
    rows.push(`<tr><th>MCOC Guide Blog<br><small>CaptainRaghu</small></th><td>${none(c.atk) && none(c.def) ? '<span class="empty">sin valorar</span>' : `${tierHtml(c.atk, 'Atq')} ${tierHtml(c.def, 'Def')}`}</td></tr>`);
    if (none(c.atk) && (c.ops.scAtk || c.ops.scDef)) {
      rows.push(`<tr><th>MCOC Scout<br><small>sustituye al Guide Blog</small></th><td>${tierHtml(c.ops.scAtk === 'A++' ? 'S+' : c.ops.scAtk, 'Atq')} ${tierHtml(c.ops.scDef, 'Def')}</td></tr>`);
    }
    return `<table class="tbl ops">${rows.join('')}</table>`;
  }

  function champItem(c, extra = '') {
    return `<li><a class="citem" href="${champHref(c)}">${avatar(c)}<span class="meta"><div class="nm">${esc(c.name)}</div>
      <div class="sub">${clsTag(c.cls)}<span>${esc(c.rar)}</span><span>Atq ${esc(none(c.atk) ? '–' : c.atk)} · Def ${esc(none(c.def) ? '–' : c.def)}</span></div>${extra}</span>${tiersMini(c)}</a></li>`;
  }
  const effChip = (e, cls = '', mode = 'any', extra = '') =>
    `<button class="chip ${cls}" data-eff="${esc(norm(e))}" data-mode="${mode}">${esc(e)}${extra}</button>`;
  const chipsOr = (arr, fn) => arr.length ? `<div class="chips">${arr.map(fn).join('')}</div>` : `<p class="empty">—</p>`;

  function setTitle(t, back) {
    titleEl.textContent = t;
    backBtn.hidden = !back;
    document.title = t === 'Campeones' ? 'MCOC Buscador' : `${t} · MCOC Buscador`;
  }

  // ---------- filtros de búsqueda ----------
  function matches(c) {
    if (S.cls.length && !S.cls.includes(c.cls)) return false;
    if (S.rar && c.rar !== S.rar) return false;
    for (const f of S.filters) {
      const t = c.tags.get(f.e);
      if (!t) return false;
      const cats = MODE_CATS[f.mode];
      if (cats && !cats.some(x => t.has(x))) return false;
    }
    if (S.q) {
      const q = norm(S.q);
      if (!norm(c.name).includes(q)) return false;
    }
    return true;
  }
  function sortChamps(arr, by) {
    const k = by === 'gen' ? c => [rank(c.gen) * 1000 + (c.genScore == null ? 999 : -c.genScore), -c.nOps]
      : by === 'atk' ? c => [rank(c.atk), rank(c.yt)] : by === 'def' ? c => [rank(c.def), rank(c.yt)] : by === 'name' ? null : c => [rank(c.yt), rank(c.atk)];
    return arr.sort((a, b) => {
      if (k) { const x = k(a), y = k(b); if (x[0] !== y[0]) return x[0] - y[0]; if (x[1] !== y[1]) return x[1] - y[1]; }
      return a.name.localeCompare(b.name);
    });
  }
  function whyLine(c) {
    if (!S.filters.length) return '';
    return `<div class="why">${S.filters.map(f => {
      const cats = [...(c.tags.get(f.e) || [])];
      const wanted = MODE_CATS[f.mode];
      const shown = (wanted ? cats.filter(x => wanted.includes(x)) : cats).map(x => CAT[x]);
      return `<b>${esc(effectLabel(f.e))}:</b> ${esc(shown.join(', '))}`;
    }).join(' · ')}</div>`;
  }

  // Sinónimos para el buscador: palabras coloquiales en español y nombre en inglés del juego -> efecto real.
  // Las claves son los nombres de los efectos tal como salen en los datos; si alguno no existe, se ignora.
  const SYNONYMS = {
    'Descarga': 'shock|electricidad|electrico|electrica|electrocutar|alto voltaje|voltaje|rayo|relampago|corriente|chispa|calambre|electro|trueno',
    'Neurodescarga': 'neuroshock|descarga neural|electricidad|nervios',
    'Veneno': 'poison|toxina|toxico|envenenar|ponzona|intoxicar|veneno',
    'Neurotoxina': 'neurotoxin|toxina|veneno',
    'Sangrado': 'bleed|sangre|sangrar|hemorragia|cortar|corte|herida',
    'Herida': 'injury|herir|lesion',
    'Incinerar': 'incinerate|fuego|quemar|quemadura|llamas|arder|calor|lumbre',
    'Chamuscado': 'seared|quemado|fuego',
    'Llama nova': 'nova flame|fuego|llamas',
    'Ola de frío': 'coldsnap|cold snap|frio|helar|hielo|congelar|congelado',
    'Congelación': 'frostbite|hielo|frio|escarcha|congelar|congelado',
    'Aturdimiento': 'stun|aturdir|noquear|ko|atontar|paralizar|inmovilizar',
    'Anular': 'nullify|quitar mejoras|eliminar mejoras|borrar mejoras|quitar buffs|borrar buffs|dispel',
    'Tambaleo': 'stagger|tambalear|anular siguiente mejora',
    'Rotura de armadura': 'armor break|armour break|romper armadura|quitar armadura|bajar armadura',
    'Armadura destrozada': 'armor shattered|destrozar armadura|romper armadura',
    'Aumento de armadura': 'armor up|subir armadura|blindaje|defensa|armadura',
    'Drenaje de poder': 'power drain|quitar poder|drenar poder|drenar energia|quitar barra|barra de poder',
    'Quema de poder': 'power burn|quemar poder|quemar barra|barra de poder',
    'Robo de poder': 'power steal|robar poder|robar barra|barra de poder',
    'Bloqueo de poder': 'power lock|bloquear poder|no gana poder|barra de poder',
    'Aguijón de poder': 'power sting|barra de poder',
    'Ganancia de poder': 'power gain|ganar poder|ganar barra|cargar barra|recargar|energia|barra de poder',
    'Bloqueo de curación': 'heal block|no curar|impedir curacion|anticuracion|sin curacion',
    'Regeneración': 'regeneration|regen|curar|curacion|sanar|recuperar vida|vida|heal',
    'Robo de vida': 'life steal|vampiro|chupar vida|drenar vida|curar',
    'Degeneración': 'degeneration|degen|perder vida|dano continuo',
    'Ruptura': 'rupture|romper',
    'Controles invertidos': 'reverse controls|reversed controls|invertir controles|controles al reves|reves',
    'Sello del destino': 'fate seal|sellar destino|bloquear habilidades',
    'Delirio': 'delirium|locura|alucinacion|confusion',
    'Reducción de precisión de habilidades': 'ability accuracy|precision de habilidades|fallar habilidades',
    'Imbloqueable': 'unblockable|no se puede bloquear|atravesar bloqueo|romper bloqueo|sin bloqueo',
    'Imparable': 'unstoppable|no se puede interrumpir|super armadura|sin retroceso|no para',
    'Intocable': 'untouchable|invulnerable|no recibe dano|inmune al dano|intocable',
    'Esquivar': 'dodge|dodging|esquivar|esquiva|evitar golpes',
    'Evasión': 'evade|evadir|esquivar|esquiva|evitar golpes',
    'Atravesar (Phasing)': 'phasing|phase|fase|fantasma|traspasar|atravesar',
    'Furia': 'fury|mas ataque|aumento de ataque|subir ataque|fuerza|potencia|poder de ataque',
    'Precisión': 'precision|critico|criticos|probabilidad de critico',
    'Crueldad': 'cruelty|dano critico|critico|criticos',
    'Daño crítico': 'critical damage|critico|criticos',
    'Resistencia a críticos': 'critical resistance|critico|criticos',
    'Sin críticos': 'no critical|no critico|sin criticos',
    'Debilidad': 'weakness|debilitar|menos ataque|reducir ataque|bajar ataque',
    'Vulnerabilidad': 'vulnerability|vulnerable|mas dano recibido',
    'Fragilidad': 'fragility|fragil',
    'Lentitud': 'slow|ralentizar|lento|no esquiva',
    'Provocación (Taunt)': 'taunt|provocar|atraer',
    'Daño verdadero': 'true damage|ignorar armadura|ignora armadura',
    'Golpe verdadero': 'true strike|ignorar esquiva|ignora esquiva|ignorar intocable|ignorar evasion',
    'Precisión verdadera': 'true accuracy|ignorar esquiva|ignora esquiva',
    'Sentido verdadero': 'true sense|ignorar phasing|ignora phasing|ignorar fase',
    'Enfoque verdadero': 'true focus',
    'Penetración física': 'physical penetration|atravesar armadura',
    'Penetración de energía': 'energy penetration|atravesar resistencia',
    'Bloqueo perfecto': 'perfect block|bloquear todo|sin dano al bloquear',
    'Penetración de bloqueo': 'block penetration|atravesar bloqueo',
    'Escudo de combo': 'combo shield|escudo',
    'Campo de fuerza': 'force field|escudo|barrera|proteccion',
    'Indestructible': 'indestructible|invencible|no muere|no puede morir',
    'Imperecedero': 'undying|revivir|resucitar|no morir|sobrevivir|inmortal|segunda vida',
    'Inmortalidad': 'immortality|revivir|resucitar|no morir|sobrevivir|inmortal|segunda vida',
    'Invisibilidad': 'invisibility|invisible|sigilo|camuflaje',
    'Retroceso': 'knockback|empujar|derribar|tirar al suelo',
    'Autorreparación': 'self-repair|self repair|reparar|curar',
    'Agotamiento': 'exhaustion|cansancio|cansar|agotar',
    'Agotar': 'exhaust|cansancio|cansar',
    'Fatiga': 'fatigue|cansancio|cansar',
    'Desorientar': 'disorient|confundir|mareo|marear',
    'Encanto': 'charm|hechizar|seducir|control mental',
    'Control mental': 'mind control|controlar mente|telepatia|hipnosis',
    'Sueño': 'sleep|dormir|dormido',
    'Petrificar': 'petrify|piedra|convertir en piedra',
    'Bloqueo de especiales': 'special lock|sin especiales|bloquear especiales|no puede usar especiales',
    'Desarmar': 'disarm|quitar arma',
    'Intimidar': 'intimidate|miedo|asustar',
    'Miedo': 'fear|asustar|terror',
    'Pánico': 'panic|miedo',
    'Cobardía': 'cowardice|miedo|cobarde',
    'Enfurecer': 'infuriate|enfado|enfadar',
    'Ira': 'rage|enfado|cabreo',
    'Berserker': 'berserk|furia ciega|descontrol',
    'Frenesí': 'frenzy|descontrol',
    'Corrosión': 'corrosion|acido|corroer',
    'Quemadura ácida': 'acid burn|acido|corroer',
    'Marcado': 'marked|marcar|marca',
    'Rastreo': 'tracking|perseguir|seguir',
    'Detonación de combo': 'combo detonation|explotar|explosion|detonar|bomba',
    'Detonación de poder': 'power detonation|explotar|explosion|detonar|bomba',
    'Represalia': 'retribution|devolver dano|venganza',
    'Reflejo de daño': 'damage reflection|reflejar dano|devolver dano|espinas',
    'Contragolpe': 'counterpunch|contraataque|contraatacar',
    'Bloqueo automático': 'auto-block|auto block|autobloqueo|bloquear solo',
    'Purificar': 'purify|limpiar perjuicios|quitar perjuicios|quitar debuffs|limpiar debuffs',
    'Limpiar': 'cleanse|limpiar perjuicios|quitar perjuicios|quitar debuffs',
    'todos los perjuicios (debuffs)': 'debuff|debuffs|perjuicios|estados negativos|efectos negativos',
    'Desintegración': 'disintegration|desintegrar',
    'Neutralizar': 'neutralize|neutralizar',
    'Golpe de refilón': 'glance|glancing|rozar|roce',
    'Fallo (Miss)': 'miss|fallar|fallo',
    'Aumento de resistencia': 'resistance up|subir resistencia|resistencia',
    'Resistencia física': 'physical resistance|resistencia|reducir dano fisico',
    'Resistencia a energía': 'energy resistance|resistencia|reducir dano de energia',
    'Daño de energía': 'energy damage|energia',
    'Daño físico': 'physical damage|golpes',
    'Daño directo': 'direct damage|dano directo',
    'Ataques sin contacto': 'non contact attacks|proyectil|disparo|a distancia|sin contacto',
    'Ataques de contacto': 'contact attacks|cuerpo a cuerpo|contacto',
    'Adrenalina': 'adrenaline',
    'Celo': 'zeal',
    'Envalentonar': 'embolden',
    'Aptitud': 'aptitude',
    'Pericia': 'expertise',
    'Tenacidad': 'tenacity',
    'Firmeza': 'steadfast',
    'Coraje': 'grit',
    'Supresión': 'suppression',
    'Sofocar': 'stifle',
    'Socavar': 'undermine',
    'Hendir': 'sunder',
    'Marchitar': 'wither',
    'Atrofia': 'atrophy',
    'Vacilar': 'falter',
    'Torpeza': 'fumble',
    'Enervar': 'enervate',
    'Obstaculizar': 'hinder',
    'Tranquilizar': 'tranquilize|calmar|sedar',
    'Desacelerar': 'decelerate|frenar',
    'Debilitar': 'debilitate',
    'Conmoción': 'concussion|golpe en la cabeza',
    'Trauma': 'trauma',
    'Tormento': 'torment',
    'Miseria': 'misery',
    'Aplastar': 'crush|aplastar',
    'Pulverizar': 'pulverize',
    'Inexorable': 'inexorable',
    'Implacable': 'implacable',
    'Aguante': 'endurance|resistir',
    'Reforzar': 'reinforce',
    'Salvaguarda': 'safeguard|proteger',
    'Protección': 'protection|proteger',
    'Vigilancia': 'vigilance',
    'Confianza': 'confidence',
    'Astucia': 'cunning',
    'Instinto': 'instinct',
    'Vitalidad': 'vitality|vida',
    'Robótica': 'robotics|robot',
    'Arena': 'sand'
  };
  const synIndex = [];
  for (const [label, syns] of Object.entries(SYNONYMS)) {
    const x = effects.get(norm(label));
    if (x) synIndex.push({ x, syns: syns.split('|').map(norm) });
  }
  // efectos cuyo sinónimo encaja con lo escrito: devuelve Map(efecto -> sinónimo que ha encajado)
  function synonymMatches(q) {
    const out = new Map();
    const words = q.split(/\s+/).filter(w => w.length >= 3);
    for (const { x, syns } of synIndex) {
      const hit = syns.find(s => s.includes(q) || (s.length >= 4 && q.includes(s)) ||
        (words.length > 1 && words.every(w => s.split(/\s+/).some(sw => sw.startsWith(w)))));
      if (hit) out.set(x, { s: hit, score: hit === q ? 3 : hit.includes(q) ? 2 : 1 });
    }
    return out;
  }

  // Inmunidades «clave»: las de los nodos más habituales (vista Equipo y recomendaciones)
  const KEY_IMM_LIST = ['Veneno', 'Sangrado', 'Incinerar', 'Descarga', 'Aturdimiento', 'Anular', 'Rotura de armadura', 'Bloqueo de poder', 'Bloqueo de curación', 'Ola de frío', 'Congelación', 'Degeneración'];
  // Todas las inmunidades que aparecen en los datos (permanentes, condicionales o sin confirmar), de más a menos campeones
  const IMM_CATS = ['imm', 'immc', 'immu'];
  const QUICK_IMM = [...effects.values()]
    .map(x => ({ label: x.label, n: [...x.champs].filter(c => [...(c.tags.get(x.key) || [])].some(t => IMM_CATS.includes(t))).length }))
    .filter(x => x.n > 0)
    .sort((a, b) => b.n - a.n || a.label.localeCompare(b.label, 'es'));

  // ---------- vista: buscar ----------
  function viewSearch() {
    setTitle('Campeones', false);
    const nFilt = S.cls.length + (S.rar ? 1 : 0);
    view.innerHTML = `
      <div class="search">
        <input id="q" type="search" autocomplete="off" placeholder="Nombre o efecto (p. ej. Veneno, fuego, curar…)" value="${esc(S.q)}" aria-label="Buscar">
        <button class="clear" id="qclear" aria-label="Borrar" ${S.q ? '' : 'hidden'}>×</button>
      </div>
      <div id="suggest"></div>
      <div id="fchips"></div>
      <details class="panel" id="fpanel" ${nFilt ? 'open' : ''}>
        <summary>Filtros y orden ${nFilt ? `<span class="badge-n">${nFilt}</span>` : ''}</summary>
        <div class="body">
          <div class="row"><label>Clase</label>${CLASSES.map(k => `<button class="toggle ${S.cls.includes(k) ? 'on' : ''}" data-cls="${esc(k)}" style="${ccVar(k)}">${esc(k)}</button>`).join('')}</div>
          <div class="row"><label>Rareza</label>${['', '7★', '6★'].map(r => `<button class="toggle ${S.rar === r ? 'on' : ''}" data-rar="${r}">${r || 'Todas'}</button>`).join('')}</div>
          <div class="row"><label>Ordenar</label>
            <select class="sel" id="sort">
              ${[['gen', 'Tier general (consenso)'], ['yt', 'Tier youtubers'], ['atk', 'Ataque (Guide Blog)'], ['def', 'Defensa (Guide Blog)'], ['name', 'Nombre']].map(([v, l]) => `<option value="${v}" ${S.sort === v ? 'selected' : ''}>${l}</option>`).join('')}
            </select>
          </div>
        </div>
      </details>
      <div id="results"></div>`;
    const q = view.querySelector('#q');
    q.addEventListener('input', () => { S.q = q.value; view.querySelector('#qclear').hidden = !S.q; renderSuggest(); renderResults(); });
    q.addEventListener('keydown', e => {
      if (e.key === 'Enter') { const b = view.querySelector('#suggest [data-add]'); if (b) { b.click(); e.preventDefault(); } q.blur(); }
    });
    view.querySelector('#qclear').addEventListener('click', () => { S.q = ''; q.value = ''; view.querySelector('#qclear').hidden = true; renderSuggest(); renderResults(); q.focus(); });
    view.querySelector('#sort').addEventListener('change', e => { S.sort = e.target.value; saveState(); renderResults(); });
    view.querySelector('#fpanel').addEventListener('click', e => {
      const b = e.target.closest('.toggle');
      if (!b) return;
      if (b.dataset.cls) {
        const k = b.dataset.cls;
        S.cls = S.cls.includes(k) ? S.cls.filter(x => x !== k) : [...S.cls, k];
        b.classList.toggle('on');
      } else {
        S.rar = b.dataset.rar;
        view.querySelectorAll('[data-rar]').forEach(x => x.classList.toggle('on', x === b));
      }
      const n = S.cls.length + (S.rar ? 1 : 0);
      view.querySelector('#fpanel summary').innerHTML = `Filtros y orden ${n ? `<span class="badge-n">${n}</span>` : ''}`;
      saveState(); renderResults();
    });
    renderSuggest(); renderChips(); renderResults();
  }

  function addFilter(e, mode = 'any') {
    const ex = S.filters.find(f => f.e === e);
    if (ex) ex.mode = mode; else S.filters.push({ e, mode });
    saveState();
  }

  function renderSuggest() {
    const box = view.querySelector('#suggest');
    const q = norm(S.q);
    if (q.length < 2) {
      box.innerHTML = (!S.filters.length && !S.q)
        ? `<p class="hint">Escribe un efecto para buscar campeones que lo tengan, o prueba un atajo: <b>inmune a…</b></p>
           <div class="chips">${QUICK_IMM.map(x => effChip(x.label, 'ok', 'immc', ` <small>${x.n}</small>`)).join('')}</div>
           <details class="panel all-eff"><summary>Todos los efectos (${effects.size})</summary><div class="body">
             <div class="chips">${[...effects.values()].sort((a, b) => a.label.localeCompare(b.label, 'es'))
               .map(x => effChip(x.label, '', 'any', ` <small>${x.champs.size}</small>`)).join('')}</div>
           </div></details>`
        : '';
      return;
    }
    const used = x => S.filters.some(f => f.e === x.key);
    const direct = [...effects.values()]
      .filter(x => norm(x.label).includes(q) && !used(x))
      .sort((a, b) => (norm(b.label).startsWith(q) - norm(a.label).startsWith(q)) || (b.champs.size - a.champs.size));
    // por sinónimo (p. ej. «alto voltaje» -> Descarga), detrás de las coincidencias directas
    const bySyn = synonymMatches(q);
    const syn = [...bySyn.keys()].filter(x => !used(x) && !direct.includes(x))
      .sort((a, b) => (bySyn.get(b).score - bySyn.get(a).score) || (b.champs.size - a.champs.size));
    const btn = (x, why) => `<button data-add="${esc(x.key)}">◆ Filtrar por <b>${esc(x.label)}</b>${why}<span class="k">${x.champs.size} campeones</span></button>`;
    box.innerHTML = direct.length || syn.length ? `<div class="suggest">${direct.map(x => btn(x, '')).join('')}${syn.map(x =>
      btn(x, ` <small class="syn-why">≈ «${esc(bySyn.get(x).s)}»</small>`)).join('')}</div>` : '';
  }

  function renderChips() {
    const box = view.querySelector('#fchips');
    box.innerHTML = S.filters.length ? `<div class="filters">${S.filters.map((f, i) => `
      <span class="fchip"><b>${esc(effectLabel(f.e))}</b>
        <select data-i="${i}" aria-label="Dónde">${MODES.map(([v, l]) => `<option value="${v}" ${f.mode === v ? 'selected' : ''}>${l}</option>`).join('')}</select>
        <button class="x" data-rm="${i}" aria-label="Quitar">×</button></span>`).join('')}
      ${S.filters.length > 1 ? `<button class="toggle" data-rmall>Quitar todos</button>` : ''}</div>` : '';
  }

  function renderResults() {
    const res = sortChamps(champs.filter(matches), S.sort);
    const box = view.querySelector('#results');
    const total = champs.length;
    box.innerHTML = `<div class="count"><span>${res.length === total ? `${total} campeones` : `${res.length} de ${total} campeones`}</span>
      ${S.filters.length > 1 ? '<span>cumplen todos los filtros</span>' : ''}</div>
      ${res.length ? `<ul class="list">${res.map(c => champItem(c, whyLine(c))).join('')}</ul>` : '<p class="empty">Ningún campeón cumple estos filtros.</p>'}`;
  }

  view.addEventListener('click', e => {
    const add = e.target.closest('[data-add]');
    if (add) { addFilter(add.dataset.add); S.q = ''; const q = view.querySelector('#q'); if (q) q.value = ''; refreshSearch(); return; }
    const rm = e.target.closest('[data-rm]');
    if (rm) { S.filters.splice(+rm.dataset.rm, 1); saveState(); refreshSearch(); return; }
    if (e.target.closest('[data-rmall]')) { S.filters = []; saveState(); refreshSearch(); return; }
    const chip = e.target.closest('[data-eff]');
    if (chip) {
      const onSearch = (location.hash || '#/buscar').startsWith('#/buscar');
      if (!onSearch) S.filters = [];
      addFilter(chip.dataset.eff, chip.dataset.mode || 'any');
      if (onSearch) refreshSearch(); else location.hash = '#/buscar';
    }
  });
  view.addEventListener('change', e => {
    const s = e.target.closest('.fchip select');
    if (s) { S.filters[+s.dataset.i].mode = s.value; saveState(); renderResults(); }
  });
  function refreshSearch() {
    if (!view.querySelector('#results')) return;
    const qc = view.querySelector('#qclear'); if (qc) qc.hidden = !S.q;
    renderSuggest(); renderChips(); renderResults();
  }

  // ---------- vista: ficha de campeón ----------
  function viewChamp(name) {
    const c = findChamp(name);
    if (!c) { view.innerHTML = '<p class="empty">Campeón no encontrado.</p>'; setTitle('Campeón', true); return; }
    setTitle(c.name, true);
    const inTeam = team.includes(c.name);
    const bc = c.bc ? bcOf(c.bc) : null;
    const synHtml = c.syn.length ? c.syn.map(s => `<div class="syn">${effChip(s.e, s.imm ? 'ok' : 'buf', 'any', s.imm ? ' <small>(inmunidad)</small>' : '')} con ${s.with.map(w => {
      const p = findChamp(w); return p ? `<a href="${champHref(p)}">${esc(w)}</a>` : esc(w);
    }).join(', ')}</div>`).join('') : '<p class="empty">—</p>';
    const kitHtml = c.kit.length ? c.kit.map(k => `<details class="kit"><summary>${esc(k.part)} ${k.pas ? '<span class="pas">PASIVA</span>' : ''}</summary>
      ${chipsOr(k.eff, e => effChip(e))}</details>`).join('') : '<p class="empty">Sin desglose del kit (ficha parcial).</p>';

    view.innerHTML = `
      <div class="hero">${avatar(c, 'lg')}<div>
        <h2>${esc(c.name)}</h2>
        <div class="sub">${clsTag(c.cls)}<span>${esc(c.rar)}</span><span>Ficha ${esc((c.ficha || '').toLowerCase())}</span></div>
      </div></div>
      <div class="tierbox">
        <div>${tierHtml(c.gen)}<p>General${c.genPos ? ` · nº ${c.genPos}` : ''}</p></div>
        <div>${tierHtml(c.yt)}<p>Youtubers</p></div>
        <div>${tierHtml(c.atk)}<p>Ataque (GB)</p></div>
        <div>${tierHtml(c.def)}<p>Defensa (GB)</p></div>
      </div>

      <details class="card opbox"><summary><h3>Comparar tier lists</h3>
        <small>${c.nOps ? `${c.nOps} opinion${c.nOps === 1 ? '' : 'es'}${c.genScore != null ? ` · nota ${c.genScore}/100` : ''}${c.agree ? ` · acuerdo ${c.agree.toLowerCase()} (${AGREE_TXT[c.agree]})` : ''}` : 'Sin valoraciones todavía'}</small></summary>
        ${opinionsHtml(c)}
        <p class="hint">El tier general es la media de todas las opiniones de ataque (todas las clases juntas). Detalles en Info → Leyenda y fuentes.</p>
      </details>

      <section class="card"><h3>Reliquia recomendada</h3>
        ${c.bc ? `<a class="relic" href="#/rb/${encodeURIComponent(bc ? bc.key : c.bc)}"><span class="ico" style="${ccVar(bc ? bc.cls : c.cls)}">⚔</span>
          <span class="meta"><b>Battlecast: ${esc(c.bc)}</b><small>Se activa en combate e invoca un ataque · ideal para atacar</small></span>›</a>` : ''}
        ${scKey(c) ? `<a class="relic" href="#/rs/${encodeURIComponent(scKey(c))}"><span class="ico" style="${ccVar(c.cls)}">⛨</span>
          <span class="meta"><b>Statcast: ${esc(c.scN || 'retrato ' + c.scP)}</b><small>Estadísticas pasivas · funciona también en defensa</small></span>›</a>` : ''}
        ${!c.bc && !scKey(c) ? '<p class="empty">Sin recomendación.</p>' : ''}
        ${c.adv ? `<div class="advice"><b>¿Cuál te conviene?</b> ${esc(c.adv)}</div>` : ''}
      </section>

      <section class="card"><h3>Inmunidades</h3>
        ${!c.immP.length && !c.immC.length && !c.immU.length ? '<p class="empty">Ninguna conocida.</p>' : ''}
        ${c.immP.length ? `<h4>Siempre</h4>${chipsOr(c.immP, e => effChip(e, 'ok', 'imm'))}` : ''}
        ${c.immC.length ? `<h4>Condicionales</h4>${chipsOr(c.immC, x => effChip(x.e, 'warn', 'immc', x.w ? ` <small>(${esc(x.w)})</small>` : ''))}` : ''}
        ${c.immU.length ? `<h4>Sin confirmar</h4>${chipsOr(c.immU, e => effChip(e, 'unk', 'immc'))}` : ''}
        ${c.res.length ? `<h4>Resistencias</h4>${chipsOr(c.res, x => effChip(x.e, 'ok', 'res', '').replace(`>${esc(x.e)}<`, `>${esc(x.txt)}<`))}` : ''}
      </section>

      <section class="card"><h3>Efectos</h3>
        <h4>Provoca al rival</h4>${chipsOr(c.prov, e => effChip(e, 'bad', 'prov'))}
        <h4>Obtiene (mejoras)</h4>${chipsOr(c.gain, e => effChip(e, 'buf', 'gain'))}
        ${c.pas.length ? `<h4>Pasivos (siempre activos)</h4>${c.pas.map(g => `<p class="empty" style="margin:4px 0">${esc(g.t)}</p>${chipsOr(g.items, e => effChip(e, 'ok', 'pas'))}`).join('')}` : ''}
        ${c.other.length ? `<h4>Otras mecánicas</h4>${chipsOr(c.other, e => effChip(e))}` : ''}
        ${c.key.length ? `<h4>Habilidades clave</h4>${chipsOr(c.key, e => effChip(e))}` : ''}
      </section>

      <section class="card"><h3>Sinergias</h3>${synHtml}</section>
      <section class="card"><h3>Mapa de habilidades</h3>${kitHtml}</section>

      <div class="btns">
        <button class="btn ${inTeam ? 'sec' : ''}" id="teamBtn">${inTeam ? 'Quitar del equipo' : '+ Añadir al equipo'}</button>
        ${c.url ? `<a class="btn sec" href="${esc(c.url)}" target="_blank" rel="noopener">Kit completo en MCOCHUB ↗</a>` : ''}
      </div>
      <p class="legal">Toca cualquier efecto para ver qué otros campeones lo tienen. Clasificación automática a partir del texto del kit en inglés: ante la duda, revisa el juego o MCOCHUB.</p>`;
    view.querySelector('#teamBtn').addEventListener('click', () => {
      if (team.includes(c.name)) team = team.filter(n => n !== c.name);
      else if (team.length >= 5) { alertSheet('El equipo ya tiene 5 campeones. Quita uno en la pestaña Equipo.'); return; }
      else team.push(c.name);
      saveTeam(); viewChamp(c.name);
    });
  }

  function isStandaloneNow() {
    return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
  }
  function alertSheet(msg) {
    const bg = document.createElement('div');
    bg.className = 'sheet-bg';
    bg.innerHTML = `<div class="sheet"><p>${esc(msg)}</p><div class="btns"><button class="btn">Vale</button></div></div>`;
    bg.addEventListener('click', e => { if (e.target === bg || e.target.closest('.btn')) bg.remove(); });
    document.body.appendChild(bg);
  }

  // ---------- vista: equipo ----------
  function pickChampion(onPick) {
    const bg = document.createElement('div');
    bg.className = 'sheet-bg';
    bg.innerHTML = `<div class="sheet"><h3>Añadir campeón</h3>
      <div class="search"><input type="search" placeholder="Buscar campeón…" autocomplete="off"></div>
      <ul class="list" style="margin-top:8px"></ul></div>`;
    const input = bg.querySelector('input'), ul = bg.querySelector('ul');
    const draw = () => {
      const q = norm(input.value);
      const arr = sortChamps(champs.filter(c => !team.includes(c.name) && (!q || norm(c.name).includes(q))), 'yt').slice(0, 60);
      ul.innerHTML = arr.map(c => `<li><button class="citem" data-pick="${esc(c.name)}">${avatar(c, 'sm')}<span class="meta"><div class="nm">${esc(c.name)}</div><div class="sub">${clsTag(c.cls)}<span>${esc(c.rar)}</span></div></span>${tiersMini(c)}</button></li>`).join('');
    };
    input.addEventListener('input', draw);
    bg.addEventListener('click', e => {
      if (e.target === bg) { bg.remove(); return; }
      const b = e.target.closest('[data-pick]');
      if (b) { bg.remove(); onPick(b.dataset.pick); }
    });
    document.body.appendChild(bg);
    draw();
    setTimeout(() => input.focus(), 50);
  }

  // ---------- Mi roster y recomendaciones de equipo ----------
  const ROSTER_MAX = 30;
  let roster = store.get('mcoc-roster', []).filter(n => byName.has(n));
  const saveRoster = () => store.set('mcoc-roster', roster);

  // Sinergias en ambos sentidos: nombre -> Map(compañero -> Set(efectos))
  const synMap = new Map();
  const addSyn = (a, b, e) => {
    if (!synMap.has(a)) synMap.set(a, new Map());
    if (!synMap.get(a).has(b)) synMap.get(a).set(b, new Set());
    synMap.get(a).get(b).add(e);
  };
  champs.forEach(c => c.syn.forEach(s => s.with.forEach(w => {
    const p = findChamp(w);
    if (p && p !== c) { addSyn(c.name, p.name, synLabel(s)); addSyn(p.name, c.name, synLabel(s)); }
  })));
  // Inmunidades «clave» (las de los nodos más habituales) que tiene cada campeón (siempre o condicional)
  const KEY_IMM = new Map(KEY_IMM_LIST.map(e => [norm(e), e]));
  champs.forEach(c => {
    c.keyImm = new Set([...c.immP, ...c.immC.map(x => x.e)].map(norm).filter(k => KEY_IMM.has(k)));
  });

  const TV = { 'S+': 10, 'S': 8, 'A': 6, 'B': 4, 'C': 2.5, 'D': 1.5, 'E': 1, 'F': 1 };
  // Valor individual (0-10) según el uso: ataque = media de tier ataque y youtubers; defensa = sobre todo tier defensa
  function champValue(c, use) {
    const main = use === 'def' ? c.def : c.atk;
    const a = TV[main], b = TV[c.yt];
    if (a == null && b == null) return 4; // sin valorar (p. ej. recién salidos)
    if (a == null) return b;
    if (b == null) return a;
    return use === 'def' ? a * 0.7 + b * 0.3 : (a + b) / 2;
  }
  const W = { syn: 2, synMax: 4, imm: 0.6, cls: 0.8 };
  function teamScore(t, use) {
    let power = 0;
    const syn = [], imm = new Set(), classes = new Set();
    t.forEach((c, i) => {
      power += champValue(c, use);
      c.keyImm.forEach(k => imm.add(k));
      classes.add(c.cls);
      for (let j = i + 1; j < t.length; j++) {
        const e = synMap.get(c.name) && synMap.get(c.name).get(t[j].name);
        if (e) syn.push({ a: c, b: t[j], e: [...e] });
      }
    });
    const total = power + W.syn * Math.min(syn.length, W.synMax) + W.imm * imm.size + W.cls * classes.size;
    return { total, power, syn, imm, classes };
  }
  function improveTeam(t, pool, use) {
    for (let pass = 0; pass < 3; pass++) {
      let changed = false;
      for (let i = 0; i < t.length; i++) {
        let best = teamScore(t, use).total;
        for (const c of pool) {
          if (t.includes(c)) continue;
          const tt = t.slice(); tt[i] = c;
          const s = teamScore(tt, use).total;
          if (s > best + 0.01) { best = s; t[i] = c; changed = true; }
        }
      }
      if (!changed) break;
    }
    return t;
  }
  // Mejores equipos dentro de `pool` (greedy desde varias semillas + mejora local); equipos distintos entre sí
  function bestTeams(pool, size, use, n = 3) {
    const seeds = pool.slice().sort((a, b) => champValue(b, use) - champValue(a, use)).slice(0, 15);
    const cands = [];
    for (const s of seeds) {
      const t = [s];
      while (t.length < size) {
        let best = null, bs = -Infinity;
        for (const c of pool) {
          if (t.includes(c)) continue;
          const sc = teamScore([...t, c], use).total;
          if (sc > bs) { bs = sc; best = c; }
        }
        t.push(best);
      }
      improveTeam(t, pool, use);
      cands.push({ t, s: teamScore(t, use) });
    }
    cands.sort((a, b) => b.s.total - a.s.total);
    const out = [];
    for (const c of cands) {
      // que cada equipo propuesto se diferencie en al menos 2 campeones de los anteriores
      if (out.every(o => c.t.filter(x => !o.t.includes(x)).length >= Math.min(2, size - 1))) out.push(c);
      if (out.length >= n) break;
    }
    return out;
  }
  // Campeones de fuera del roster que más mejorarían el equipo t (cambiando a uno de sus miembros)
  function upgradesFor(t, use, inRoster) {
    const base = teamScore(t, use);
    const res = [];
    for (const o of champs) {
      if (inRoster.has(o.name)) continue;
      let best = null;
      t.forEach((m, i) => {
        const tt = t.slice(); tt[i] = o;
        const s = teamScore(tt, use);
        if (!best || s.total > best.s.total) best = { s, out: m };
      });
      if (best.s.total > base.total + 0.5) res.push({ o, out: best.out, gain: best.s.total - base.total, s: best.s });
    }
    res.sort((a, b) => b.gain - a.gain);
    return res.slice(0, 3).map(r => {
      const why = [];
      const tierT = use === 'def' ? r.o.def : r.o.atk;
      if (!none(tierT) || !none(r.o.yt)) why.push(`tier ${use === 'def' ? 'defensa' : 'ataque'} ${none(tierT) ? '–' : tierT} (youtubers ${none(r.o.yt) ? '–' : r.o.yt})`);
      const newSyn = r.s.syn.filter(x => x.a === r.o || x.b === r.o).map(x => `${x.e.join(', ')} con ${x.a === r.o ? x.b.name : x.a.name}`);
      if (newSyn.length) why.push('sinergia: ' + newSyn.join('; '));
      const newImm = [...r.s.imm].filter(k => !base.imm.has(k)).map(k => KEY_IMM.get(k));
      if (newImm.length) why.push('cubre ' + newImm.join(', '));
      if (r.s.classes.size > base.classes.size) why.push('añade la clase ' + r.o.cls);
      const lost = base.syn.filter(x => x.a === r.out || x.b === r.out).map(x => (x.a === r.out ? x.b.name : x.a.name));
      if (lost.length) why.push('pierdes la sinergia con ' + lost.join(', '));
      return Object.assign(r, { why });
    });
  }

  function pickMany() {
    const bg = document.createElement('div');
    bg.className = 'sheet-bg';
    bg.innerHTML = `<div class="sheet"><h3>Mi roster <small class="empty" id="rcount"></small></h3>
      <div class="search"><input type="search" placeholder="Buscar campeón…" autocomplete="off"></div>
      <ul class="list" style="margin-top:8px"></ul>
      <div class="btns"><button class="btn" data-done>Listo</button></div></div>`;
    const input = bg.querySelector('input'), ul = bg.querySelector('ul');
    const draw = () => {
      bg.querySelector('#rcount').textContent = `${roster.length}/${ROSTER_MAX}`;
      const q = norm(input.value);
      const arr = sortChamps(champs.filter(c => !q || norm(c.name).includes(q)), 'yt');
      ul.innerHTML = arr.map(c => {
        const on = roster.includes(c.name);
        return `<li><button class="citem" data-tog="${esc(c.name)}">${avatar(c, 'sm')}<span class="meta"><div class="nm">${esc(c.name)}</div><div class="sub">${clsTag(c.cls)}<span>${esc(c.rar)}</span></div></span><span class="check ${on ? 'on' : ''}">${on ? '✓' : '+'}</span></button></li>`;
      }).join('');
    };
    const close = () => { bg.remove(); viewTeam(); };
    input.addEventListener('input', draw);
    bg.addEventListener('click', e => {
      if (e.target === bg || e.target.closest('[data-done]')) { close(); return; }
      const b = e.target.closest('[data-tog]');
      if (!b) return;
      const n = b.dataset.tog;
      if (roster.includes(n)) roster = roster.filter(x => x !== n);
      else if (roster.length < ROSTER_MAX) roster.push(n);
      else { b.classList.add('shake'); return; }
      saveRoster(); draw();
    });
    document.body.appendChild(bg);
    draw();
    setTimeout(() => input.focus(), 50);
  }

  const teamSeg = () => `<div class="seg">${[['team', 'Equipo'], ['top', 'Top 30'], ['roster', `Roster (${roster.length})`], ['rec', 'Recomendar']].map(([k, l]) =>
    `<button data-tseg="${k}" class="${(S.teamSeg || 'team') === k ? 'on' : ''}">${l}</button>`).join('')}</div>`;

  // ---------- Top 30 para misiones (sin prestigio) ----------
  // Nota base = tier general de ataque (consenso de youtubers + Guide Blog, 0-100). En misiones lo que importa es
  // atacar, así que no se mira la defensa ni el prestigio (rango/firma). Se suma un poco por cada inmunidad a los
  // efectos de nodo más habituales, porque permite pasar más caminos sin preparar el combate.
  const TIER_CENTER = { 'S+': 98, 'S': 91, 'A': 79, 'B': 66, 'C': 53, 'D': 38, 'E': 25, 'F': 20 };
  function questScore(c) {
    let base = c.genScore;
    if (base == null) base = TIER_CENTER[c.gen] ?? TIER_CENTER[c.atk] ?? TIER_CENTER[c.yt] ?? null;
    if (base == null) return null; // sin valorar por nadie
    const perm = KEY_IMM_LIST.filter(e => c.immP.some(x => norm(x) === norm(e)));
    const cond = KEY_IMM_LIST.filter(e => !perm.includes(e) && c.immC.some(x => norm(x.e) === norm(e)));
    // bonus pequeño (máx. 3): ordena dentro del mismo nivel, pero no hace que un S adelante a un S+
    const bonus = Math.min(3, perm.length * 0.75 + cond.length * 0.25);
    return { c, base, bonus, total: base + bonus, perm, cond };
  }
  function viewTop() {
    const only7 = !!S.top7;
    const owned = new Set(roster);
    const list = champs.filter(c => !only7 || c.rar === '7★').map(questScore).filter(Boolean)
      .sort((a, b) => b.total - a.total || a.c.name.localeCompare(b.c.name)).slice(0, 30);
    view.innerHTML = `${teamSeg()}
      <div class="row"><label>Rareza</label>
        <button class="toggle ${only7 ? '' : 'on'}" data-top7="0">Todas</button>
        <button class="toggle ${only7 ? 'on' : ''}" data-top7="1">Solo 7★</button></div>
      <p class="hint">Los 30 campeones más útiles para <b>misiones</b> (historia, eventos y Batallas), <b>sin tener en cuenta el prestigio</b>: se ordenan por su tier general de ataque (consenso de Vega, Lagacy, MetalSonicDude, Seatin y MCOC Guide Blog) y, a igualdad de nivel, van antes los que tienen más inmunidades a efectos de nodo habituales (${esc(KEY_IMM_LIST.slice(0, 5).join(', '))}…). Es orientativo.</p>
      <ol class="list top30">${list.map((x, i) => {
        const c = x.c;
        const why = [`Tier general ${none(c.gen) ? '–' : c.gen}${c.genPos ? ` (puesto ${c.genPos} del ranking)` : ''}`];
        if (x.perm.length) why.push('inmune a ' + x.perm.join(', '));
        if (x.cond.length) why.push('inmune a veces a ' + x.cond.join(', '));
        if (owned.has(c.name)) why.push('✓ en tu roster');
        return champItem(c, `<div class="why"><b>#${i + 1}</b> · ${esc(why.join(' · '))}</div>`);
      }).join('')}</ol>`;
  }

  function viewRoster() {
    const members = sortChamps(roster.map(n => byName.get(n)).filter(Boolean), 'yt');
    view.innerHTML = `${teamSeg()}
      <p class="hint">Añade tus ${ROSTER_MAX} mejores campeones. En «Recomendaciones» verás los mejores equipos que puedes montar con ellos y qué campeones de fuera los mejorarían.</p>
      <div class="roster">${members.map(c => `<div><a class="slot" href="${champHref(c)}" style="border:0">${avatar(c)}<button class="rm" data-rmroster="${esc(c.name)}" aria-label="Quitar">×</button></a><div class="slot-name">${esc(c.name)}</div></div>`).join('')}
        ${roster.length < ROSTER_MAX ? `<div><button class="slot" data-addroster aria-label="Añadir campeones">+</button><div class="slot-name">${roster.length}/${ROSTER_MAX}</div></div>` : ''}</div>
      <div class="btns"><button class="btn" data-addroster>${roster.length ? 'Editar roster' : '+ Añadir campeones'}</button>
        ${roster.length ? '<button class="btn sec" data-clearroster>Vaciar</button>' : ''}</div>`;
  }

  function viewRec() {
    const use = S.recUse || 'atk', size = S.recSize || 5;
    const pool = roster.map(n => byName.get(n)).filter(Boolean);
    const ctrl = `${teamSeg()}
      <div class="row"><label>Uso</label>
        <button class="toggle ${use === 'atk' ? 'on' : ''}" data-recuse="atk">Atacar</button>
        <button class="toggle ${use === 'def' ? 'on' : ''}" data-recuse="def">Defender</button></div>
      <div class="row"><label>Tamaño</label>
        <button class="toggle ${size === 5 ? 'on' : ''}" data-recsize="5">5 (Misiones)</button>
        <button class="toggle ${size === 3 ? 'on' : ''}" data-recsize="3">3 (Guerra)</button></div>`;
    if (pool.length < size) {
      view.innerHTML = `${ctrl}<p class="hint">Necesitas al menos ${size} campeones en tu roster (tienes ${pool.length}).</p>
        <div class="btns"><button class="btn" data-tseg="roster">Ir a Mi roster</button></div>`;
      return;
    }
    const inRoster = new Set(roster);
    const teams = bestTeams(pool, size, use, 3);
    view.innerHTML = `${ctrl}
      <p class="hint">Puntuación = nivel de los campeones (tiers) + sinergias entre ellos + inmunidades clave cubiertas + variedad de clases. Es orientativa: no conoce el rango ni la firma de tus campeones.</p>
      ${teams.map((x, i) => {
        const ups = upgradesFor(x.t, use, inRoster);
        return `<section class="card"><h3>Equipo ${i + 1} · ${x.s.total.toFixed(1)} puntos</h3>
          <div class="team" style="grid-template-columns:repeat(${size},1fr)">${x.t.map(c => `<div><a class="slot" href="${champHref(c)}" style="border:0">${avatar(c)}</a><div class="slot-name">${esc(c.name)}</div></div>`).join('')}</div>
          ${x.s.syn.length ? `<p><b>Sinergias:</b> ${x.s.syn.map(s => `${esc(s.a.name)} + ${esc(s.b.name)} <small class="empty">(${esc(s.e.join(', '))})</small>`).join(' · ')}</p>` : '<p class="empty">Sin sinergias conocidas entre ellos.</p>'}
          <p><b>Inmunes a:</b> ${x.s.imm.size ? esc([...x.s.imm].map(k => KEY_IMM.get(k)).join(', ')) : '—'}</p>
          <p><b>Clases:</b> ${[...x.s.classes].map(clsTag).join(' ')}</p>
          ${ups.length ? `<h4>Lo mejorarían (fuera de tu roster)</h4><ul class="list">${ups.map(u => `<li><a class="citem" href="${champHref(u.o)}">${avatar(u.o, 'sm')}<span class="meta">
            <div class="nm">${esc(u.o.name)} <small class="empty">por ${esc(u.out.name)}</small></div><div class="why">+${u.gain.toFixed(1)} · ${esc(u.why.join(' · '))}</div></span></a></li>`).join('')}</ul>` : ''}
          <div class="btns"><button class="btn sec" data-useteam="${esc(x.t.map(c => c.name).join('|'))}">Usar este equipo</button></div>
        </section>`;
      }).join('')}`;
  }

  function viewTeam() {
    setTitle('Equipo', false);
    if (S.teamSeg === 'roster') return viewRoster();
    if (S.teamSeg === 'rec') return viewRec();
    if (S.teamSeg === 'top') return viewTop();
    const members = team.map(n => byName.get(n)).filter(Boolean);
    const slots = [];
    for (let i = 0; i < 5; i++) {
      const c = members[i];
      slots.push(c
        ? `<div><a class="slot" href="${champHref(c)}" style="border:0">${avatar(c)}<button class="rm" data-rmteam="${esc(c.name)}" aria-label="Quitar">×</button></a><div class="slot-name">${esc(c.name)}</div></div>`
        : `<div><button class="slot" data-addteam aria-label="Añadir campeón">+</button><div class="slot-name">&nbsp;</div></div>`);
    }
    let body = '';
    if (!members.length) {
      body = `<p class="hint">Monta un equipo de hasta 5 campeones para ver juntos sus inmunidades, los efectos que provocan, las sinergias entre ellos y la reliquia de cada uno.</p>`;
    } else {
      // inmunidades
      const imm = new Map();
      const addImm = (e, c, kind) => { const k = norm(e); if (!imm.has(k)) imm.set(k, { e, who: [] }); imm.get(k).who.push({ c, kind }); };
      members.forEach(c => { c.immP.forEach(e => addImm(e, c, '')); c.immC.forEach(x => addImm(x.e, c, 'cond.')); c.immU.forEach(e => addImm(e, c, '¿?')); });
      const immRows = [...imm.values()].sort((a, b) => b.who.length - a.who.length || a.e.localeCompare(b.e));
      const gaps = KEY_IMM_LIST.filter(e => !imm.has(norm(e)));
      // efectos provocados
      const prov = new Map();
      members.forEach(c => c.prov.forEach(e => { const k = norm(e); if (!prov.has(k)) prov.set(k, { e, n: 0 }); prov.get(k).n++; }));
      // sinergias
      const inner = [], outer = new Map();
      members.forEach(c => c.syn.forEach(s => s.with.forEach(w => {
        const p = findChamp(w);
        if (p && team.includes(p.name)) inner.push(`<b>${esc(c.name)}</b> + <b>${esc(p.name)}</b>: ${esc(synLabel(s))}`);
        else if (p) { if (!outer.has(p.name)) outer.set(p.name, { p, items: [] }); outer.get(p.name).items.push(`${synLabel(s)} (con ${c.name})`); }
      })));
      // también sinergias que otros declaran hacia miembros del equipo
      members.forEach(c => champs.forEach(o => {
        if (team.includes(o.name)) return;
        o.syn.forEach(s => { if (s.with.some(w => findChamp(w) === c)) { if (!outer.has(o.name)) outer.set(o.name, { p: o, items: [] }); outer.get(o.name).items.push(`${synLabel(s)} (con ${c.name})`); } });
      }));
      const outerArr = [...outer.values()].sort((a, b) => b.items.length - a.items.length || rank(a.p.yt) - rank(b.p.yt)).slice(0, 12);
      const innerU = [...new Set(inner)];

      body = `
        <section class="card"><h3>Inmunidades del equipo</h3>
          ${immRows.length ? `<div class="covgrid">${immRows.map(r => `<div class="cov"><span class="e">${esc(r.e)}</span><span class="w">${r.who.map(w => `<b>${esc(w.c.name)}</b>${w.kind ? ` (${w.kind})` : ''}`).join(', ')}</span></div>`).join('')}</div>` : '<p class="empty">Nadie tiene inmunidades conocidas.</p>'}
          ${gaps.length ? `<h4>Nadie es inmune a</h4><div class="chips">${gaps.map(e => effChip(e, 'unk', 'imm')).join('')}</div><p class="empty">Toca uno para ver qué campeones sí lo son.</p>` : ''}
        </section>
        <section class="card"><h3>Efectos que provoca el equipo</h3>
          ${prov.size ? `<div class="chips">${[...prov.values()].sort((a, b) => b.n - a.n || a.e.localeCompare(b.e)).map(x => effChip(x.e, 'bad', 'prov', x.n > 1 ? ` <small>×${x.n}</small>` : '')).join('')}</div>` : '<p class="empty">—</p>'}
        </section>
        <section class="card"><h3>Sinergias dentro del equipo</h3>
          ${innerU.length ? innerU.map(s => `<div class="syn">${s}</div>`).join('') : '<p class="empty">Ninguna sinergia conocida entre estos campeones.</p>'}
          ${outerArr.length ? `<h4>Campeones que harían sinergia</h4><ul class="list">${outerArr.map(o => champItem(o.p, `<div class="why">${esc([...new Set(o.items)].join(' · '))}</div>`)).join('')}</ul>` : ''}
        </section>
        <section class="card"><h3>Reliquias</h3>
          ${members.map(c => `<div class="relic">${avatar(c, 'sm')}<span class="meta"><b>${esc(c.name)}</b>
            <small>${c.bc ? `Battlecast: ${esc(c.bc)}` : ''}${c.bc && scKey(c) ? ' · ' : ''}${scKey(c) ? `Statcast: ${esc(c.scN || c.scP)}` : ''}</small>
            ${c.adv ? `<small>${esc(c.adv)}</small>` : ''}</span></div>`).join('')}
        </section>`;
    }
    view.innerHTML = `${teamSeg()}<div class="team">${slots.join('')}</div>
      ${members.length ? `<div class="btns"><button class="btn sec" data-clearteam>Vaciar equipo</button></div>` : ''}${body}`;
  }
  view.addEventListener('click', e => {
    if (e.target.closest('[data-addteam]')) { pickChampion(n => { if (team.length < 5) { team.push(n); saveTeam(); } viewTeam(); }); return; }
    const rm = e.target.closest('[data-rmteam]');
    if (rm) { e.preventDefault(); team = team.filter(n => n !== rm.dataset.rmteam); saveTeam(); viewTeam(); return; }
    if (e.target.closest('[data-clearteam]')) { team = []; saveTeam(); viewTeam(); return; }
    const ts = e.target.closest('[data-tseg]');
    if (ts) { S.teamSeg = ts.dataset.tseg; saveState(); window.scrollTo(0, 0); viewTeam(); return; }
    if (e.target.closest('[data-addroster]')) { pickMany(); return; }
    const rr = e.target.closest('[data-rmroster]');
    if (rr) { e.preventDefault(); roster = roster.filter(n => n !== rr.dataset.rmroster); saveRoster(); viewTeam(); return; }
    if (e.target.closest('[data-clearroster]')) {
      const bg = document.createElement('div');
      bg.className = 'sheet-bg';
      bg.innerHTML = `<div class="sheet"><p>¿Vaciar tu roster de ${roster.length} campeones?</p><div class="btns"><button class="btn" data-yes>Vaciar</button><button class="btn sec">Cancelar</button></div></div>`;
      bg.addEventListener('click', ev => {
        if (ev.target.closest('[data-yes]')) { roster = []; saveRoster(); viewTeam(); }
        if (ev.target === bg || ev.target.closest('.btn')) bg.remove();
      });
      document.body.appendChild(bg);
      return;
    }
    const t7 = e.target.closest('[data-top7]');
    if (t7) { S.top7 = t7.dataset.top7 === '1'; saveState(); viewTeam(); return; }
    const ru = e.target.closest('[data-recuse]');
    if (ru) { S.recUse = ru.dataset.recuse; saveState(); viewTeam(); return; }
    const rs = e.target.closest('[data-recsize]');
    if (rs) { S.recSize = +rs.dataset.recsize; saveState(); viewTeam(); return; }
    const ut = e.target.closest('[data-useteam]');
    if (ut) { team = ut.dataset.useteam.split('|'); saveTeam(); S.teamSeg = 'team'; saveState(); window.scrollTo(0, 0); viewTeam(); }
  });

  // ---------- vista: reliquias ----------
  function relicChampList(arr, mode) {
    const sorted = sortChamps(arr.slice(), mode === 'def' ? 'def' : 'yt');
    return `<ul class="list">${sorted.map(c => champItem(c, c.adv ? `<div class="why">${esc(c.adv)}</div>` : '')).join('')}</ul>`;
  }

  function viewRelics() {
    setTitle('Reliquias', false);
    const seg = S.relSeg;
    let body = '';
    if (seg === 'bc') {
      const rows = battlecasts.map(b => {
        const users = champs.filter(c => c.bc && norm(bcBase(c.bc)) === norm(b.key));
        const best = sortChamps(users.slice(), 'yt').slice(0, 3);
        return { b, users, best };
      }).sort((x, y) => y.users.length - x.users.length);
      body = `<p class="hint">Se activan durante el combate e invocan un ataque. Lo mejor para <b>atacar</b>. Toca una para ver a quién ponérsela.</p>
        ${rows.map(({ b, users, best }) => `<a class="relic" href="#/rb/${encodeURIComponent(b.key)}"><span class="ico" style="${ccVar(b.cls)}">⚔</span>
          <span class="meta"><b>${esc(b.name)}</b><small>${esc(b.cls)} · ${users.length} campeones · mejores: ${esc(best.map(c => c.name).join(', ') || '—')}</small></span>›</a>`).join('')}`;
    } else if (seg === 'sc') {
      body = `<p class="hint">Dan estadísticas pasivas todo el combate y funcionan también <b>defendiendo</b>. Se llaman «[Prefijo] [Clase] Relic of [Sufijo]»; MCOC Scout suele recomendarlas por el retrato que aparece en la reliquia.</p>
        ${CLASSES.map(cls => {
          const sc = [...statcasts.values()].filter(s => s.cls === cls).sort((a, b) => b.champs.length - a.champs.length);
          const fam = scFamilies.filter(f => f.cls === cls);
          return `<details class="panel"><summary><span class="cls" style="${ccVar(cls)}">${esc(cls)}</span></summary><div class="body">
            <h4 class="hint">Recomendadas (por retrato o nombre)</h4>
            ${sc.map(s => `<a class="relic" href="#/rs/${encodeURIComponent(s.key)}"><span class="ico" style="${ccVar(cls)}">⛨</span><span class="meta"><b>${esc(s.exact ? s.key : 'Retrato ' + s.key)}</b><small>${s.champs.length} campeones</small></span>›</a>`).join('')}
            <h4 class="hint">Familias y variantes</h4>
            <table class="tbl">${fam.map(f => `<tr><th>${esc(f.prefix)}${f.portrait ? `<br><small>${esc(f.portrait)}</small>` : ''}</th><td>${esc(f.suffixes.join(', '))}</td></tr>`).join('')}</table>
          </div></details>`;
        }).join('')}`;
    } else {
      const use = S.bestUse;
      const arr = champs.filter(c => (!S.bestCls || c.cls === S.bestCls) && (!S.best7 || c.rar === '7★') && !none(use === 'atk' ? c.atk : c.def) && (use === 'atk' ? c.bc : scKey(c)));
      const sorted = sortChamps(arr, use).slice(0, 40);
      body = `<p class="hint">Las mejores parejas campeón + reliquia según el tier de MCOC Guide Blog. Para atacar se muestra el Battlecast; para defender, el Statcast.</p>
        <div class="row"><label>Uso</label>
          <button class="toggle ${use === 'atk' ? 'on' : ''}" data-use="atk">Atacar</button>
          <button class="toggle ${use === 'def' ? 'on' : ''}" data-use="def">Defender</button></div>
        <div class="row"><label>Clase</label>
          <select class="sel" id="bestCls"><option value="">Todas</option>${CLASSES.map(k => `<option ${S.bestCls === k ? 'selected' : ''}>${esc(k)}</option>`).join('')}</select>
          <button class="toggle ${S.best7 ? 'on' : ''}" data-best7>Solo 7★</button></div>
        <ul class="list">${sorted.map(c => {
          const relic = use === 'atk' ? `⚔ Battlecast: ${c.bc}` : `⛨ Statcast: ${c.scN || c.scP}`;
          return `<li><a class="citem" href="${champHref(c)}">${avatar(c)}<span class="meta"><div class="nm">${esc(c.name)}</div>
            <div class="sub">${clsTag(c.cls)}<span>${esc(c.rar)}</span></div><div class="why">${esc(relic)}</div></span>
            <span class="tiers">${tierHtml(use === 'atk' ? c.atk : c.def, use === 'atk' ? 'Atq' : 'Def')}</span></a></li>`;
        }).join('')}</ul>`;
    }
    view.innerHTML = `<div class="seg">${[['bc', 'Battlecasts'], ['sc', 'Statcasts'], ['best', 'Mejores combos']].map(([k, l]) => `<button data-seg="${k}" class="${seg === k ? 'on' : ''}">${l}</button>`).join('')}</div>${body}`;
    const bs = view.querySelector('#bestCls');
    if (bs) bs.addEventListener('change', () => { S.bestCls = bs.value; saveState(); viewRelics(); });
  }
  view.addEventListener('click', e => {
    const sg = e.target.closest('[data-seg]');
    if (sg) { S.relSeg = sg.dataset.seg; saveState(); viewRelics(); return; }
    const u = e.target.closest('[data-use]');
    if (u) { S.bestUse = u.dataset.use; saveState(); viewRelics(); return; }
    if (e.target.closest('[data-best7]')) { S.best7 = !S.best7; saveState(); viewRelics(); }
  });

  function viewBattlecast(key) {
    const b = bcOf(key);
    const name = b ? b.name : key;
    setTitle(`Battlecast ${b ? b.key : key}`, true);
    const users = champs.filter(c => c.bc && norm(bcBase(c.bc)) === norm(bcBase(key)));
    view.innerHTML = `
      <section class="card"><h3>Battlecast · ${esc(b ? b.cls : '')}</h3>
        <p><b>${esc(name)}</b></p>
        <p>${b && b.eff ? esc(b.eff) : 'Efectos detallados pendientes de completar: consulta la carta de la reliquia en el juego.'}</p>
        <p class="empty">Las Battlecasts se activan cuando se cumple una condición en combate y aparece un campeón invitado que ataca («Striker»). Sus runas se personalizan; el duplicado desbloquea la habilidad innata.</p>
      </section>
      <h3 class="hint">Recomendada para ${users.length} campeones (de mejor a peor)</h3>
      ${relicChampList(users, 'atk')}`;
  }

  function viewStatcast(key) {
    const s = statcasts.get(key);
    setTitle(s && !s.exact ? `Statcast ${key}` : 'Statcast', true);
    if (!s) { view.innerHTML = '<p class="empty">Reliquia no encontrada.</p>'; return; }
    const fam = scFamilies.filter(f => f.cls === s.cls);
    const confirmed = fam.find(f => f.portrait && norm(f.portrait).includes(norm(key)));
    view.innerHTML = `
      <section class="card"><h3>Statcast · ${esc(s.cls)}</h3>
        <p><b>${esc(s.exact ? key : 'Reliquia con el retrato de ' + key)}</b></p>
        ${confirmed ? `<p>Corresponde a la familia <b>${esc(confirmed.pattern)}</b> (${esc(confirmed.suffixes.join(', '))}).</p>` : !s.exact ? `<p class="empty">Aún no está confirmado a qué prefijo corresponde este retrato. Familias de ${esc(s.cls)}: ${esc(fam.map(f => f.prefix).join(', '))}.</p>` : ''}
        <p class="empty">Da estadísticas pasivas durante todo el combate y también funciona cuando el campeón defiende (Guerra de Alianzas, Battlegrounds).</p>
      </section>
      <h3 class="hint">Recomendada para ${s.champs.length} campeones (de mejor a peor)</h3>
      ${relicChampList(s.champs, 'def')}`;
  }

  // ---------- vista: ranking por clase ----------
  const RK_BY = [['gen', 'General'], ['yt', 'Youtubers'], ['atk', 'Ataque'], ['def', 'Defensa']];
  const RK_SRC = {
    gen: 'Consenso de hasta 5 opiniones de ataque: Vega, Lagacy, MetalSonicDude y Seatin (vía mcoc.app) y MCOC Guide Blog. Cada una se pasa a una nota 0-100 y se hace la media. ⚠ = las fuentes discrepan mucho.',
    yt: 'Media de Vega, Lagacy, MetalSonicDude y Seatin (mcoc.app).',
    atk: 'MCOC Guide Blog, ataque (S+ … E). No valora Deathless ni Stellar-Forged.',
    def: 'MCOC Guide Blog, defensa (S … D). No valora Deathless ni Stellar-Forged.'
  };
  const ALL = 'Todas';
  function viewRanking() {
    const by = RK_BY.some(([k]) => k === S.rkBy) ? S.rkBy : 'gen';
    const cls = S.rkCls === ALL || CLASSES.includes(S.rkCls) ? S.rkCls : ALL;
    setTitle(cls === ALL ? 'Tier list general' : 'Ranking por clase', false);
    const tierOf = c => (by === 'gen' ? c.gen : by === 'atk' ? c.atk : by === 'def' ? c.def : c.yt);
    const arr = sortChamps(champs.filter(c => (cls === ALL || c.cls === cls) && (!S.rk7 || c.rar === '7★')), by);    // agrupar por tier
    const groupsArr = [];
    for (const c of arr) {
      const t = none(tierOf(c)) ? '' : tierOf(c);
      let g = groupsArr[groupsArr.length - 1];
      if (!g || g.t !== t) { g = { t, items: [] }; groupsArr.push(g); }
      g.items.push(c);
    }
    let pos = 0;
    const valOf = (k, c) => (k === 'gen' ? c.gen : k === 'atk' ? c.atk : k === 'def' ? c.def : c.yt);
    const subOf = c => {
      const parts = [];
      if (cls === ALL) parts.push(clsTag(c.cls));
      parts.push(`<span>${esc(c.rar)}</span>`);
      if (by === 'gen') {
        parts.push(`<span>${c.genScore != null ? `${c.genScore}/100 · ${c.nOps} op.` : 'sin datos'}</span>`);
        if (c.agree === 'Bajo') parts.push('<span class="disc" title="Las fuentes discrepan mucho">⚠ discrepan</span>');
      }
      RK_BY.filter(([k]) => k !== by && (by !== 'gen' || k !== 'yt')).forEach(([k, l]) => parts.push(`<span>${l} ${esc(none(valOf(k, c)) ? '–' : valOf(k, c))}</span>`));
      return parts.join('');
    };
    view.innerHTML = `
      <div class="row rk-cls"><button class="toggle ${cls === ALL ? 'on' : ''}" data-rkcls="${ALL}" style="--cc:#e8e8ee">Todas</button>${CLASSES.map(k => `<button class="toggle ${k === cls ? 'on' : ''}" data-rkcls="${esc(k)}" style="${ccVar(k)}">${esc(k)}</button>`).join('')}</div>
      <div class="seg">${RK_BY.map(([k, l]) => `<button data-rkby="${k}" class="${by === k ? 'on' : ''}">${l}</button>`).join('')}</div>
      <div class="count"><span>${arr.length} campeones${cls === ALL ? '' : ' ' + esc(cls)}</span><button class="toggle ${S.rk7 ? 'on' : ''}" data-rk7>Solo 7★</button></div>
      <p class="hint">${esc(RK_SRC[by])}</p>
      ${groupsArr.map(g => `<h3 class="tier-head">${g.t ? tierHtml(g.t) : '<span class="tier">–</span>'}<span>${g.t ? `Tier ${esc(g.t)}` : 'Sin valorar'}</span><small>${g.items.length}</small></h3>
        <ul class="list">${g.items.map(c => { pos++; return `<li><a class="citem" href="${champHref(c)}"><span class="pos">${g.t ? pos : ''}</span>${avatar(c)}<span class="meta"><div class="nm">${esc(c.name)}</div>
          <div class="sub">${subOf(c)}</div></span></a></li>`; }).join('')}</ul>`).join('')}`;
  }
  view.addEventListener('click', e => {
    const a = e.target.closest('[data-rkcls]');
    if (a) { S.rkCls = a.dataset.rkcls; saveState(); viewRanking(); return; }
    const b = e.target.closest('[data-rkby]');
    if (b) { S.rkBy = b.dataset.rkby; saveState(); viewRanking(); return; }
    if (e.target.closest('[data-rk7]')) { S.rk7 = !S.rk7; saveState(); viewRanking(); }
  });

  // ---------- vista: info ----------
  function viewInfo() {
    setTitle('Info', false);
    const tbl = name => `<table class="tbl">${RAW.sheets[name].rows.slice(1).filter(r => r[0] && r[1]).map(r => `<tr><th>${esc(r[0])}</th><td>${esc(r[1])}</td></tr>`).join('')}</table>`;
    const n7 = champs.filter(c => c.rar === '7★').length;
    view.innerHTML = `
      <section class="card"><h3>Sobre los datos</h3>
        <p>${champs.length} campeones (${n7} con 7★ y ${champs.length - n7} solo con 6★), ${battlecasts.length} Battlecasts y ${scFamilies.length} familias de Statcast. Datos de ${esc(RAW.generated)} a partir de <i>${esc(RAW.source)}</i>.</p>
        <p>Funciona sin conexión: todos los datos e imágenes van dentro de la app.</p>
      </section>
      ${isStandaloneNow() ? '' : `<section class="card"><h3>Instalar en el móvil</h3>
        <p>Añade la app a la pantalla de inicio para abrirla como una app más y usarla sin conexión.</p>
        <button class="install-btn" style="font-size:16px;padding:12px 20px;width:100%" onclick="document.getElementById('installBtn').click()">⤓ Instalar en la pantalla de inicio</button>
      </section>`}
      <details class="panel"><summary>Cómo instalarla en el móvil</summary><div class="body">
        <p><b>Android (Chrome):</b> abre la web y pulsa ⋮ → «Instalar aplicación» o «Añadir a pantalla de inicio».</p>
        <p><b>iPhone (Safari):</b> pulsa Compartir → «Añadir a pantalla de inicio».</p>
        <p>Tras abrirla una vez con conexión, queda guardada y funciona sin internet.</p>
      </div></details>
      <details class="panel"><summary>Guía rápida de reliquias</summary><div class="body">${tbl('Guía de reliquias')}</div></details>
      <details class="panel"><summary>Leyenda y fuentes</summary><div class="body">${tbl('Leyenda')}</div></details>
      <p class="legal"><b>Proyecto de fans, no afiliado a Kabam ni a Marvel.</b> MARVEL Batalla de Superhéroes, los nombres de los campeones y sus imágenes son marcas y propiedad de Kabam Games y Marvel. Tiers: opiniones de sus autores (Vega, Lagacy, MetalSonicDude y Seatin, recopiladas por <a href="https://mcoc.app" target="_blank" rel="noopener">mcoc.app</a>; <a href="https://www.mcocguideblog.com/p/mcoc-offense-and-defense-tier-list.html" target="_blank" rel="noopener">MCOC Guide Blog</a> — CaptainRaghu; <a href="https://mcocscout.com" target="_blank" rel="noopener">MCOC Scout</a> para los que el Guide Blog no valora). El tier general es una media calculada por esta app. Kits, inmunidades y retratos: <a href="https://mcochub.insaneskull.com" target="_blank" rel="noopener">MCOCHUB</a>. Reliquias recomendadas: <a href="https://mcocscout.com/relics" target="_blank" rel="noopener">MCOC Scout</a>.</p>`;
  }

  // ---------- enrutado ----------
  const scrollMem = {};
  let lastHash = '';
  function route() {
    const h = location.hash || '#/buscar';
    if (lastHash) scrollMem[lastHash] = window.scrollY;
    const [, sec, arg] = h.match(/^#\/([^/]+)\/?(.*)$/) || [, 'buscar', ''];
    const a = decodeURIComponent(arg || '');
    const tab = { buscar: 'buscar', c: 'buscar', ranking: 'ranking', equipo: 'equipo', reliquias: 'reliquias', rb: 'reliquias', rs: 'reliquias', info: 'info' }[sec] || 'buscar';
    document.querySelectorAll('.tabbar a').forEach(x => x.classList.toggle('on', x.dataset.tab === tab));
    if (sec === 'c') viewChamp(a);
    else if (sec === 'ranking') viewRanking();
    else if (sec === 'equipo') viewTeam();
    else if (sec === 'reliquias') viewRelics();
    else if (sec === 'rb') viewBattlecast(a);
    else if (sec === 'rs') viewStatcast(a);
    else if (sec === 'info') viewInfo();
    else viewSearch();
    window.scrollTo(0, scrollMem[h] || 0);
    lastHash = h;
  }
  let navCount = 0;
  window.addEventListener('hashchange', () => { navCount++; route(); });
  backBtn.addEventListener('click', () => {
    if (navCount > 0) history.back();
    else location.hash = (location.hash.startsWith('#/r') ? '#/reliquias' : '#/buscar');
  });
  route();

  // ---------- botón «Instalar» ----------
  const installBtn = document.getElementById('installBtn');
  const isStandalone = () => window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
  const isIOS = /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  let installPrompt = null;
  const isAndroid = /android/i.test(navigator.userAgent);
  const isInApp = /FBAN|FBAV|Instagram|WhatsApp|Line\/|Telegram|; wv\)/i.test(navigator.userAgent);
  function updateInstallBtn() {
    // Siempre visible salvo cuando ya se abre como app instalada. Si el navegador no ofrece la instalación
    // directa (iPhone, Firefox, Samsung Internet, navegadores dentro de otras apps…), el botón muestra instrucciones.
    installBtn.hidden = isStandalone();
  }
  function installHelp() {
    if (isInApp) return 'Estás dentro de otra app (WhatsApp, Instagram…). Abre el enlace en Chrome (Android) o Safari (iPhone): menú ⋮ o ··· → «Abrir en el navegador». Después vuelve a pulsar «Instalar».';
    if (isIOS) return 'En iPhone/iPad: abre esta página en Safari, pulsa Compartir (el cuadrado con la flecha hacia arriba) y elige «Añadir a pantalla de inicio».';
    if (isAndroid) return 'En Android: abre el menú del navegador (⋮ arriba a la derecha en Chrome, ≡ abajo en Samsung Internet) y elige «Instalar aplicación» o «Añadir a pantalla de inicio». Si no aparece, recarga la página y espera unos segundos. Si ya la tienes instalada, búscala entre tus aplicaciones.';
    return 'En el ordenador: en Chrome o Edge, pulsa el icono de instalar de la barra de direcciones (una pantalla con una flecha) o menú ⋮ → «Guardar y compartir» → «Instalar página como aplicación». Firefox y Safari de escritorio no permiten instalarla.';
  }
  window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); installPrompt = e; updateInstallBtn(); });
  window.addEventListener('appinstalled', () => { installPrompt = null; installBtn.hidden = true; });
  window.matchMedia('(display-mode: standalone)').addEventListener('change', updateInstallBtn);
  installBtn.addEventListener('click', async () => {
    if (installPrompt) {
      installPrompt.prompt();
      const { outcome } = await installPrompt.userChoice;
      if (outcome === 'accepted') { installPrompt = null; installBtn.hidden = true; }
    } else {
      alertSheet(installHelp());
    }
  });
  updateInstallBtn();

  // ---------- modo offline (PWA) ----------
  if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
    navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).catch(() => { /* sin SW: la app sigue funcionando */ });
  }
})();
