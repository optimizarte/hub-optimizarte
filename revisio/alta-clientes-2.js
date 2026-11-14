
// ─── STATE ───────────────────────────────────────────────────────
var tipo = 'par';
var vtosPerSi = false;
var autEmpSi = false;
var colabRecoge = 'M441819E';
var colabAsigna = 'M441819E';
var lastData = {};
var ramoData = {};
var currentRamoKey = null;
var currentRamoType = null;
var currentPillEl = null;
var ramoCounts = {};
var COLABS = { 'M441819E':'Dany Hernandez', 'M354046Y':'Fadoua Khachach', 'MA48168T':'Silvia Famoso' };
// ─── NUMBER FORMATTING ────────────────────────────────────────────
function numRaw(s){ return (s||'').toString().replace(/\./g,'').replace(',','.'); }
function numFmt(s){
  var raw=(s||'').toString().replace(/[^0-9]/g,'');
  if(!raw) return '';
  return raw.replace(/\B(?=(\d{3})+(?!\d))/g,'.');
}
function numFmtInput(el){
  var pos=el.selectionStart;var len=el.value.length;
  el.value=numFmt(el.value);
  var newLen=el.value.length;
  var newPos=pos+(newLen-len);if(newPos<0)newPos=0;
  try{el.setSelectionRange(newPos,newPos);}catch(e){}
}


var PILL_STATE_LABELS = {
  '': '\u2014 Estado \u2014',
  'vto-inmediato': 'VTO Inmediato / Tarificar',
  'no-necesita': 'No necesita',
  'necesita': 'Necesita pero no tiene',
  'facilita': 'Tiene en competencia',
  'no-cambia-vto': 'No cambia \u00b7 Da VTOS',
  'no-cambia': 'No quiere cambiar',
  'en-vigor': 'Contratado / En Vigor'
};

// ─── RAMO CONFIG ─────────────────────────────────────────────────
var RAMO_CONFIG = {
  'hogar':         { name: 'Hogar',              emoji: '\uD83C\uDFE0', hasModal: true  },
  'auto':          { name: 'Auto',               emoji: '\uD83D\uDE97', hasModal: true  },
  'moto':          { name: 'Moto',               emoji: '\uD83C\uDFCD\uFE0F', hasModal: true  },
  'vida':          { name: 'Vida',               emoji: '\u2764\uFE0F', hasModal: true  },
  'salud':         { name: 'Salud',              emoji: '\uD83C\uDFE5', hasModal: true  },
  'decesos':       { name: 'Decesos',            emoji: '\uD83D\uDD4A\uFE0F', hasModal: true  },
  'ahorro-g':      { name: 'Ahorro Garantizado', emoji: '\uD83D\uDCB0', hasModal: true  },
  'ahorro-i':      { name: 'Ahorro Inversion',   emoji: '\uD83D\uDCC8', hasModal: true  },
  'embarcaciones': { name: 'Embarcaciones',      emoji: '\u26F5', hasModal: true  },
  'comunidades':   { name: 'Comunidades',        emoji: '\uD83C\uDFD8\uFE0F', hasModal: true  },
  'movilidad':     { name: 'Movilidad Personal', emoji: '\uD83D\uDEF4', hasModal: false },
  'mascotas':      { name: 'Mascotas',           emoji: '\uD83D\uDC3E', hasModal: false },
  'bienes-consumo':{ name: 'Bienes Consumo',     emoji: '\uD83D\uDED2', hasModal: false },
  'plan-pensiones':{ name: 'Plan Pensiones',     emoji: '\uD83C\uDFE6', hasModal: false },
  'accidentes':    { name: 'Accidentes',         emoji: '\uD83E\uDE79', hasModal: false },
  'comercio-pyme': { name: 'Comercio / Pyme',    emoji: '\uD83C\uDFEA', hasModal: false },
  'rc':            { name: 'Resp. Civil',        emoji: '\u2696\uFE0F', hasModal: false },
  'acc-convenio':  { name: 'Acc. Convenio',      emoji: '\uD83D\uDCCB', hasModal: false },
  'acc-autonomo':  { name: 'Acc. Autonomo',      emoji: '\uD83D\uDCBC', hasModal: false },
  'salud-col':     { name: 'Salud Colectivo',    emoji: '\uD83D\uDC65', hasModal: false },
  'ahorro-col':    { name: 'Ahorro Colectivo',   emoji: '\uD83D\uDC8E', hasModal: false },
  'transportes':   { name: 'Transportes',        emoji: '\uD83D\uDE9A', hasModal: false },
  'baja-ilt-aut':  { name: 'Baja ILT Aut\u00f3nomo', emoji: '\uD83E\uDE7A', hasModal: false },
  'subsidio':      { name: 'Subsidio',           emoji: '\uD83D\uDCB6', hasModal: false },
  'otro':          { name: 'Otro',               emoji: '\uD83D\uDCDD', hasModal: true  }
};

var GRIDS = {
  'grid-estrategicos': ['hogar','auto','moto','vida','salud','decesos','ahorro-g','ahorro-i','otro'],
  'grid-ofertables': ['embarcaciones','comunidades','movilidad','mascotas','bienes-consumo','plan-pensiones','accidentes','otro'],
  'grid-neg-actividad': ['comercio-pyme','rc','transportes','otro'],
  'grid-neg-empleados': ['acc-convenio','salud-col','ahorro-col','otro'],
  'grid-neg-autonomo': ['baja-ilt-aut','salud','subsidio','otro']
};

// ─── DATA BADGE ──────────────────────────────────────────────────
function isModalComplete(type, d) {
  if (!d) return false;
  if (type==='hogar') return !!(d.tipoProp && d.personas > 0);
  if (type==='auto'||type==='moto') return !!(d.matricula && d.mesVcto);
  if (type==='salud') return !!(d.personas > 0 && d.modalidad);
  if (type==='decesos') return !!(d.personas > 0);
  if (type==='vida') return !!(d.personas > 0 || d.capital);
  if (type==='ahorro-g'||type==='ahorro-i') return !!(d.tieneAhorros || d.quiereAhorrar);
  if (type==='comunidades') return !!(d.nombreCom);
  if (type==='embarcaciones') return !!(d.eslora && d.nmotores);
  if (type==='otro') return !!(d.notas && d.notas.trim());
  return false;
}
function hasAnyData(type, d) {
  if (!d) return false;
  if (d._saved) return true;
  var keys = Object.keys(d).filter(function(k){return k!=='estado'&&k!=='_saved';});
  return keys.some(function(k){
    var v=d[k];
    if(typeof v==='boolean') return v;
    if(Array.isArray(v)) return v.length>0;
    return v!==''&&v!==0&&v!==null&&v!==undefined;
  });
}
function buildDataBadge(type, key) {
  var cfg = RAMO_CONFIG[type];
  if (!cfg || !cfg.hasModal) return '';
  var d = ramoData[key];
  if (!d || !hasAnyData(type, d)) return '';
  if (d.estado === 'no-necesita') return ''; // explicitly not needed
  if (d.estado === 'en-vigor') return ''; // contratado, no badge needed
  var ok = isModalComplete(type, d);
  return '<span class="'+(ok?'data-badge-ok':'data-badge-warn')+'">'+(ok?'\u2713':'\u0021')+'</span>';
}
function refreshPillBadge(key) {
  var pills = document.querySelectorAll('[data-key="'+key+'"]');
  if (!pills.length) return;
  pills.forEach(function(pill) {
    var type = pill.getAttribute('data-type');
    var d = ramoData[key] || {};
    // Gold pulse: vto-inmediato AND (no modal needed OR modal complete)
    var typeBase = type.replace(/-\d+$/,'');
    var hasModal = RAMO_CONFIG[typeBase] && RAMO_CONFIG[typeBase].hasModal;
    if(d.estado==='vto-inmediato' && (!hasModal || isModalComplete(typeBase,d)) && d.estado!=='en-vigor'){
      pill.classList.add('pill-gold-pulse');
    } else {
      pill.classList.remove('pill-gold-pulse');
    }
    pill.querySelectorAll('.data-badge-ok,.data-badge-warn').forEach(function(b){b.remove();});
    var badgeHtml = buildDataBadge(type, key);
    if (badgeHtml) {
      var tmp = document.createElement('div'); tmp.innerHTML = badgeHtml;
      var badgeEl = tmp.firstChild;
      pill.appendChild(badgeEl);
      if (badgeEl.classList.contains('data-badge-warn')) {
        pill.classList.add('pill-badge-danger');
      } else {
        pill.classList.remove('pill-badge-danger');
      }
    } else {
      pill.classList.remove('pill-badge-danger');
    }
  });
}

// ─── BUILD PILL ──────────────────────────────────────────────────
function buildPill(type, key, isDup) {
  var cfg = RAMO_CONFIG[type];
  if (!cfg) return null;
  var pill = document.createElement('div');
  pill.className = 'ramo-pill';
  pill.setAttribute('data-type', type);
  pill.setAttribute('data-key', key);
  if (isDup) pill.setAttribute('data-dup', '1');

  var badge = isDup ? '<span class="ramo-badge">#' + key.split('-').pop() + '</span>' : '';
  var dataBadge = buildDataBadge(type, key);
  var delBtn = isDup
    ? '<button class="rpa-btn rpa-del" type="button" onclick="delPill(this);event.stopPropagation()" title="Eliminar">\u2715</button>'
    : '';

  pill.innerHTML = badge + dataBadge +
    '<div class="rp-icon-col"><span class="rp-icon">' + cfg.emoji + '</span></div>' +
    '<div class="rp-name">' + cfg.name + '</div>' +
    '<div class="rp-acts-col">' +
      '<button class="rpa-btn rpa-dup" type="button" title="A\u00f1adir otro">+</button>' +
      delBtn +
    '</div>' +
    '<input type="checkbox" name="seg" value="' + cfg.name + '" style="display:none">';

  // "+" button duplicates without opening popup
  var dupBtn = pill.querySelector('.rpa-dup');
  if (dupBtn) dupBtn.addEventListener('click', function(e) { e.stopPropagation(); dupPill(dupBtn); });
  var delBtnEl = pill.querySelector('.rpa-del');
  if (delBtnEl) delBtnEl.addEventListener('click', function(e) { e.stopPropagation(); delPill(delBtnEl); });

  // Clicking icon or name opens state popup
  pill.addEventListener('click', function(e) {
    if (e.target.closest('.rp-acts-col')) return;
    openPillPopup(pill, e);
  });

  var d = ramoData[key] || {};
  if (d.estado) applyPillState(pill, d.estado, false);
  return pill;
}

function renderGrid(gridId, types) {
  var grid = document.getElementById(gridId);
  if (!grid) return;
  grid.innerHTML = '';
  var seen = {};
  types.forEach(function(t) {
    // 'otro' pill gets unique key per grid
    var key = (t === 'otro') ? (gridId + '-otro-1') : (t + '-1');
    var p = buildPill(t, key, false);
    if (p) grid.appendChild(p);
  });
}

function initGrids() {
  for (var gId in GRIDS) renderGrid(gId, GRIDS[gId]);
}

function updateNegocioGroups() {
  var isEmp = tipo === 'emp';
  var isAut = tipo === 'aut';
  // actividad + autonomo always visible for both emp and aut
  var grA = document.getElementById('neg-grupo-actividad');
  var grE = document.getElementById('neg-grupo-empleados');
  var grO = document.getElementById('neg-grupo-autonomo');
  if (grA) grA.style.display = (isEmp || isAut) ? '' : 'none';
  if (grO) grO.style.display = (isEmp || isAut) ? '' : 'none';
  // empleados: always for emp, only if autEmpSi for aut
  if (grE) grE.style.display = (isEmp || (isAut && autEmpSi)) ? '' : 'none';
  // Re-render actividad grid: AUT includes Acc.Convenio, EMP does not
  if (isEmp) renderGrid('grid-neg-actividad', ['comercio-pyme','rc','transportes','otro']);
  else if (isAut) renderGrid('grid-neg-actividad', ['comercio-pyme','rc','acc-convenio','transportes','otro']);
}

// ─── PILL POPUP ──────────────────────────────────────────────────
function openPillPopup(pill, ev) {
  if (ev) { ev.stopPropagation(); ev.preventDefault(); }
  var popup = document.getElementById('pillPopup');
  if (!popup) return;
  var key = pill.getAttribute('data-key');
  var d = ramoData[key] || {};
  var current = d.estado || '';
  popup.querySelectorAll('.pp-opt').forEach(function(opt) {
    opt.classList.toggle('active', opt.getAttribute('data-val') === current);
  });
  currentPillEl = pill;
  // Show first so we can measure
  popup.style.cssText = 'display:block!important;visibility:hidden;z-index:99999;position:fixed';
  var pw = popup.offsetWidth || 250;
  var ph = popup.offsetHeight || 200;
  popup.style.cssText = '';
  popup.style.display = 'block';
  popup.style.zIndex = '99999';
  var rect = pill.getBoundingClientRect();
  var spaceBelow = window.innerHeight - rect.bottom;
  var left = Math.min(rect.left, window.innerWidth - pw - 8);
  var top = spaceBelow > ph + 8 ? rect.bottom + 4 : rect.top - ph - 4;
  popup.style.left = Math.max(4, left) + 'px';
  popup.style.top = Math.max(4, top) + 'px';
}

function closePillPopup() {
  document.getElementById('pillPopup').style.display = 'none';
  currentPillEl = null;
}

document.addEventListener('click', function(e) {
  if (!e.target.closest('#pillPopup') && !e.target.closest('.ramo-pill')) {
    closePillPopup();
  }
});

function selectPillOption(opt, ev) {
  if (ev) ev.stopPropagation();
  var val = opt.getAttribute('data-val');
  if (currentPillEl) {
    applyPillState(currentPillEl, val, true);
  }
  closePillPopup();
}

// ─── PILL STATE ──────────────────────────────────────────────────
function applyPillState(pill, val, openModal) {
  var type = pill.getAttribute('data-type');
  var key  = pill.getAttribute('data-key');
  var cb   = pill.querySelector('input[name="seg"]');

  pill.classList.remove('pill-gris','pill-naranja','pill-verde','pill-azul','pill-rojo','pill-purpura','pill-en-vigor','active');
  // Remove EN VIGOR stamp if present
  var oldStamp=pill.querySelector('.rp-stamp-envigor');if(oldStamp)oldStamp.remove();
  if (cb) cb.checked = false;

  var needsModal = openModal && RAMO_CONFIG[type] && RAMO_CONFIG[type].hasModal;
  if (val === 'vto-inmediato') {
    pill.classList.add('pill-purpura','active');
    if (cb) cb.checked = true;
    if (needsModal) setTimeout(function() { openRamoModal(type, key); }, 180);
  } else if (val === 'no-necesita') {
    pill.classList.add('pill-gris');
  } else if (val === 'necesita') {
    pill.classList.add('pill-naranja','active');
    if (cb) cb.checked = true;
    if (needsModal) setTimeout(function() { openRamoModal(type, key); }, 180);
  } else if (val === 'facilita') {
    pill.classList.add('pill-verde','active');
    if (cb) cb.checked = true;
    if (needsModal) setTimeout(function() { openRamoModal(type, key); }, 180);
  } else if (val === 'no-cambia-vto') {
    pill.classList.add('pill-azul','active');
    if (cb) cb.checked = true;
    if (needsModal) setTimeout(function() { openRamoModal(type, key); }, 180);
  } else if (val === 'no-cambia') {
    pill.classList.add('pill-rojo','active');
    if (cb) cb.checked = true;
    // No modal — solo registra el rechazo
  } else if (val === 'en-vigor') {
    pill.classList.add('pill-en-vigor');
    if (cb) cb.checked = true;
    // Add EN VIGOR stamp overlay
    var stamp=document.createElement('div');stamp.className='rp-stamp-envigor';stamp.textContent='EN VIGOR';
    pill.appendChild(stamp);
  }

  // Sync visual state to any other pills sharing the same key (e.g. Salud in two grids)
  document.querySelectorAll('[data-key="'+key+'"]').forEach(function(p){
    if(p===pill) return;
    p.classList.remove('pill-gris','pill-naranja','pill-verde','pill-azul','pill-rojo','pill-purpura','pill-en-vigor','active');
    var os=p.querySelector('.rp-stamp-envigor');if(os)os.remove();
    var cb2=p.querySelector('input[name="seg"]');
    if(val==='vto-inmediato'){p.classList.add('pill-purpura','active');if(cb2)cb2.checked=true;}
    else if(val==='no-necesita'){p.classList.add('pill-gris');if(cb2)cb2.checked=false;}
    else if(val==='necesita'){p.classList.add('pill-naranja','active');if(cb2)cb2.checked=true;}
    else if(val==='facilita'){p.classList.add('pill-verde','active');if(cb2)cb2.checked=true;}
    else if(val==='no-cambia-vto'){p.classList.add('pill-azul','active');if(cb2)cb2.checked=true;}
    else if(val==='no-cambia'){p.classList.add('pill-rojo','active');if(cb2)cb2.checked=true;}
    else if(val==='en-vigor'){p.classList.add('pill-en-vigor');if(cb2)cb2.checked=true;var s2=document.createElement('div');s2.className='rp-stamp-envigor';s2.textContent='EN VIGOR';p.appendChild(s2);}
  });

  var d = ramoData[key] || {};
  d.estado = val;
  ramoData[key] = d;
  refreshPillBadge(key);
  updProg();
}

function dupPill(btn) {
  var pill = btn.closest('.ramo-pill');
  var type = pill.getAttribute('data-type');
  // For 'otro' pills use the existing key as base to avoid collisions across grids
  var baseKey = pill.getAttribute('data-key');
  var countKey = type === 'otro' ? baseKey : type;
  var n = (ramoCounts[countKey] || 1) + 1;
  ramoCounts[countKey] = n;
  var key = (type === 'otro') ? (baseKey.replace(/-\d+$/, '') + '-' + n) : (type + '-' + n);
  var newPill = buildPill(type, key, true);
  var grid = pill.parentElement;
  var all = grid.querySelectorAll('[data-type="' + type + '"]');
  var last = all[all.length - 1];
  last.parentNode.insertBefore(newPill, last.nextSibling);
}

function delPill(btn) {
  var pill = btn.closest('.ramo-pill');
  if (pill) {
    var key = pill.getAttribute('data-key');
    if (key) delete ramoData[key];
    pill.remove();
    updProg();
  }
}

// ─── TIPO CLIENTE ────────────────────────────────────────────────
function setTipo(t, el) {
  var oldTipo = tipo; // Guardar tipo anterior
  tipo = t;
  
  // TASCA 1: Copiar dades entre par ↔ aut si canvia entre aquests dos
  if ((oldTipo === 'par' && t === 'aut') || (oldTipo === 'aut' && t === 'par')) {
    copiarDatosEntreParAut(oldTipo, t);
  }
  
  document.querySelectorAll('.tipo-opt').forEach(function(o) {
    o.classList.remove('active'); o.querySelector('input').checked = false;
  });
  el.classList.add('active'); el.querySelector('input').checked = true;
  vtosPerSi = false;
  var vsi = document.getElementById('vper-si'); var vno = document.getElementById('vper-no');
  if (vsi) { vsi.classList.remove('si','no'); } if (vno) { vno.classList.remove('si','no'); }

  // Si estem en mode 'initial' (només card 0 visible) i l'usuari clica un tipus
  // → mostrem TOT el formulari (mode 'nou-client')
  if (typeof _currentFormMode !== 'undefined' && _currentFormMode === 'initial' && typeof setFormMode === 'function') {
    setFormMode('nou-client');
  } else {
    updateVisibility();
  }
  updProg();
}

// TASCA 1: Funció per copiar dades personals entre particular i autònom
function copiarDatosEntreParAut(desde, hacia) {
  // Camps a copiar (només els que tenen prefix par-/aut-)
  var campos = [
    'nombre', 'ap1', 'ap2', 'nif', 'fnac', 
    'nie-caducidad', 'nacionalidad',
    'carnet1-fecha', 'carnet1-tipo',
    'carnet2-fecha', 'carnet2-tipo',
    'sexo', 'estcivil', 'hijos'
  ];
  
  campos.forEach(function(campo) {
    var idDesde = desde + '-' + campo;
    var idHacia = hacia + '-' + campo;
    var elDesde = document.getElementById(idDesde);
    var elHacia = document.getElementById(idHacia);
    
    // Només copiar si ambdós camps existeixen i destí està buit
    if (elDesde && elHacia && elDesde.value && !elHacia.value) {
      elHacia.value = elDesde.value;
    }
  });
}

function updateVisibility() {
  var isPar = tipo==='par', isEmp = tipo==='emp', isAut = tipo==='aut';
  var isNegocio = isEmp || isAut;
  document.getElementById('block-par').style.display = isPar ? 'block' : 'none';
  document.getElementById('block-emp').style.display = isEmp ? 'block' : 'none';
  document.getElementById('block-aut').style.display = isAut ? 'block' : 'none';
  document.getElementById('card-negocio').style.display = isNegocio ? 'block' : 'none';
  document.getElementById('card-vtos-per-ask').style.display = isNegocio ? 'block' : 'none';
  var showPar = isPar || (isNegocio && vtosPerSi);
  document.getElementById('card-par-segs').style.display = showPar ? 'block' : 'none';
  updateNegocioGroups();
}

function setVtosPer(val) {
  vtosPerSi = val === 'si';
  var si = document.getElementById('vper-si'); var no = document.getElementById('vper-no');
  si.classList.remove('si','no'); no.classList.remove('si','no');
  if (vtosPerSi) si.classList.add('si'); else no.classList.add('no');
  updateVisibility();
}

function setAutEmp(val) {
  autEmpSi = val === 'si';
  var si = document.getElementById('aut-emp-si'); var no = document.getElementById('aut-emp-no');
  si.classList.remove('si','no'); no.classList.remove('si','no');
  if (autEmpSi) si.classList.add('si'); else no.classList.add('no');
  var wrap = document.getElementById('aut-emp-select-wrap');
  if (wrap) wrap.style.display = autEmpSi ? 'block' : 'none';
  updateNegocioGroups();
}

function setSexo(val, btn, fid) {
  var sel = btn.closest('.sexo-selector');
  sel.querySelectorAll('.sexo-btn').forEach(function(b) { b.classList.remove('active'); });
  btn.classList.add('active');
  var f = document.getElementById(fid); if (f) f.value = val;
}

function fillEmpleados(id) {
  var s = document.getElementById(id); if (!s) return;
  var h = '<option value="">\u2014 N\u00ba empleados \u2014</option>';
  for (var i=1; i<=100; i++) h += '<option value="' + i + '">' + i + '</option>';
  s.innerHTML = h;
}

function setColab(fila, c, el) {
  if (fila==='recoge') colabRecoge = c; else colabAsigna = c;
  var row = el.closest('.colab-row');
  row.querySelectorAll('.colab-opt').forEach(function(o) { o.classList.remove('active'); });
  el.classList.add('active'); updProg();
}

function updProg() {
  var fields = [];
  if (tipo==='par') { fields.push(document.getElementById('par-nombre')); fields.push(document.getElementById('par-nif')); }
  else if (tipo==='emp') { fields.push(document.getElementById('emp-razon')); fields.push(document.getElementById('emp-cif')); }
  else if (tipo==='aut') { fields.push(document.getElementById('aut-nombre')); fields.push(document.getElementById('aut-nif')); }
  ['tel1','email1','dir','cp','muni'].forEach(function(id) { fields.push(document.getElementById(id)); });
  var filled = fields.filter(function(f) { return f && f.value.trim(); }).length;
  var total = Math.max(fields.length, 1) + 2;
  var pct = Math.min(Math.round(((filled + 2) / total) * 100), 100);
  document.getElementById('progFill').style.width = pct + '%';
}

document.querySelectorAll('input, select, textarea').forEach(function(el) {
  el.addEventListener('input', updProg); el.addEventListener('change', updProg);
});

// ─── VALIDATION / COLLECT ────────────────────────────────────────
function v(id) {
  var el = document.getElementById(id);
  if (!el || !el.value.trim()) { if(el){el.style.borderColor='#DC0028';el.focus();} return false; }
  el.style.borderColor=''; return true;
}
function validate() {
  var ok = true;
  if (tipo==='par') { if(!v('par-nombre'))ok=false; if(!v('par-nif'))ok=false; }
  else if (tipo==='emp') { if(!v('emp-razon'))ok=false; if(!v('emp-cif'))ok=false; }
  else if (tipo==='aut') { if(!v('aut-nombre'))ok=false; if(!v('aut-nif'))ok=false; }
  if(!v('tel1'))ok=false; if(!v('email1'))ok=false;
  if(!v('dir'))ok=false; if(!v('cp'))ok=false; if(!v('muni'))ok=false;
  return ok;
}
function collectData() {
  var g = function(id){var el=document.getElementById(id);return el?el.value.trim():'';};
  var seguros = Array.from(document.querySelectorAll('.ramo-pill.active input[name="seg"]')).map(function(c){return c.value;});
  var origenSw = Array.from(document.querySelectorAll('input[name="origen-sw"]:checked')).map(function(c){return c.value;});
  var tipoLabel = {par:'Particular',emp:'Empresa',aut:'Aut\u00f3nomo'}[tipo]||tipo;
  var nombreCompleto = tipo==='emp' ? g('emp-razon') :
    tipo==='aut' ? [g('aut-nombre'),g('aut-ap1'),g('aut-ap2')].filter(Boolean).join(' ') :
    [g('par-nombre'),g('par-ap1'),g('par-ap2')].filter(Boolean).join(' ');
  return {
    tipo:tipo, tipo_label:tipoLabel, nombre_completo:nombreCompleto,
    nif_cif:tipo==='emp'?g('emp-cif'):(tipo==='aut'?g('aut-nif'):g('par-nif')),
    fecha_nacimiento:tipo==='aut'?g('aut-fnac'):g('par-fnac'),
    sexo:tipo==='aut'?g('aut-sexo'):g('par-sexo'),
    estado_civil:g('par-estcivil'), hijos:g('par-hijos'),
    actividad:tipo==='emp'?g('emp-actividad'):g('aut-actividad'),
    antiguedad:tipo==='emp'?g('emp-antiguedad'):g('aut-antiguedad'),
    empleados:tipo==='emp'?g('empleados-emp'):g('empleados-aut'),
    tel1:g('tel1'), whatsapp:g('wapp'), redes:g('redes'),
    email1:g('email1'), email2:g('email2'),
    direccion:g('dir'), cp:g('cp'), localidad:g('muni'),
    seguros:seguros, ramo_data:JSON.stringify(ramoData),
    colab_recoge:COLABS[colabRecoge]||colabRecoge, colab_recoge_racf:colabRecoge,
    colab_asigna:COLABS[colabAsigna]||colabAsigna, colab_asigna_racf:colabAsigna,
    motivo_contacto:g('obs'), primera_accion:g('primera-accion'), fecha_accion:g('fecha-accion'),
    origen_sw:origenSw.join(', '), origen_detalle:g('origen-detalle'),
    timestamp:new Date().toLocaleString('es-ES'),
    // Individual fields for reliable restoration
    par_nombre:g('par-nombre'), par_ap1:g('par-ap1'), par_ap2:g('par-ap2'),
    par_nif:g('par-nif'), par_fnac:g('par-fnac'), par_sexo:g('par-sexo'),
    par_estcivil:g('par-estcivil'), par_hijos:g('par-hijos'),
    emp_razon:g('emp-razon'), emp_cif:g('emp-cif'), emp_actividad:g('emp-actividad'),
    emp_antiguedad:g('emp-antiguedad'), emp_empleados:g('empleados-emp'),
    aut_nombre:g('aut-nombre'), aut_ap1:g('aut-ap1'), aut_ap2:g('aut-ap2'),
    aut_nif:g('aut-nif'), aut_fnac:g('aut-fnac'), aut_sexo:g('aut-sexo'),
    aut_actividad:g('aut-actividad'), aut_antiguedad:g('aut-antiguedad'),
    aut_empleados:g('empleados-aut'),
    vtos_per:vtosPerSi, aut_emp:autEmpSi
  };
}

// ─── SUBMIT / CLEAR / SAVE ────────────────────────────────────────
function submitForm() {
  if (!validate()) return;
  lastData = collectData();
  lastData.id = Date.now();
  // Save to localStorage (fallback)
  try{var lista=JSON.parse(localStorage.getItem('optimizarte_clientes')||'[]');lista.push(lastData);localStorage.setItem('optimizarte_clientes',JSON.stringify(lista));}catch(e){}
  // Save to file if folder connected
  if (clientesDir) {
    saveClientToFile(lastData).then(function(fname) {
      if (fname) showToast('\ud83d\udcbe Registrado: '+fname,'');
      else showToast('Error al guardar el archivo','error');
    });
  } else {
    showToast('\u26a0\ufe0f Conecta una carpeta para guardar el archivo. Guardado solo en localStorage.','warn');
  }
  document.getElementById('modal-data').innerHTML = buildSummary(lastData);
  document.getElementById('modal').classList.add('show');
}
function buildSummary(d) {
  return ['<strong>'+d.nombre_completo+'</strong> ('+d.tipo_label+')',
    'NIF/CIF: '+d.nif_cif,'Tel: '+d.tel1+(d.email1?' \u00b7 '+d.email1:''),
    'Dir: '+d.direccion+', '+d.cp+' '+d.localidad,
    'Ramos: '+(d.seguros.join(', ')||'\u2014'),
    'Recoge: '+d.colab_recoge+' \u00b7 Asignado: '+d.colab_asigna,
    d.primera_accion?'Acci\u00f3n: '+d.primera_accion+(d.fecha_accion?' \u00b7 '+d.fecha_accion:''):'',
    d.motivo_contacto?'Motivo: '+d.motivo_contacto:''
  ].filter(Boolean).join('<br>');
}
function copyData() {
  var lines=[lastData.nombre_completo+' ('+lastData.tipo_label+')',
    'NIF/CIF: '+lastData.nif_cif,'Tel: '+lastData.tel1+(lastData.email1?' | '+lastData.email1:''),
    'Dir: '+lastData.direccion+', '+lastData.cp+' '+lastData.localidad,
    'Ramos: '+(lastData.seguros.join(', ')||'\u2014'),
    'Recoge: '+lastData.colab_recoge+' | Asignado: '+lastData.colab_asigna,
    lastData.motivo_contacto?'Motivo: '+lastData.motivo_contacto:''
  ].filter(Boolean).join('\n');
  navigator.clipboard.writeText(lines).then(function(){
    var btn=event.target;var orig=btn.textContent;btn.textContent='\u2713 Copiado!';
    setTimeout(function(){btn.textContent=orig;},2000);
  });
}
function newAlt(){document.getElementById('modal').classList.remove('show');clearForm();}
function confirmClear(){
  showToast('\u26a0\ufe0f \u00bfBorrar todos los datos? Confirma pulsando Aceptar.','warn');
  setTimeout(function(){
    if(window.confirm('\u00bfEst\u00e1s seguro/a?\nSe borrar\u00e1n todos los datos del formulario.')){
      clearForm();
    }
  },100);
}
function clearForm(){
  setFormDates(null, null);
  document.getElementById('altaForm').reset();
  tipo='par';vtosPerSi=false;autEmpSi=false;
  document.querySelectorAll('.tipo-opt').forEach(function(o){o.classList.remove('active');o.querySelector('input').checked=false;});
  document.getElementById('tipo-par').classList.add('active');
  document.getElementById('tipo-par').querySelector('input').checked=true;
  document.querySelectorAll('.sexo-btn').forEach(function(b){b.classList.remove('active');});
  document.getElementById('aut-emp-select-wrap').style.display='none';
  document.getElementById('aut-emp-si').classList.remove('si','no');
  document.getElementById('aut-emp-no').classList.remove('si','no');
  colabRecoge='M441819E';colabAsigna='M441819E';
  document.querySelectorAll('.colab-opt').forEach(function(o,i){o.classList.toggle('active',i===0||i===3);});
  ramoData={};ramoCounts={};
  initGrids();updateVisibility();updProg();
  window.scrollTo({top:0,behavior:'smooth'});
}
function saveDraft(){
  var d=collectData();
  try{localStorage.setItem('optimizarte_draft',JSON.stringify(d));}catch(e){}
  var btn=event.target; var orig=btn.textContent;
  if(clientesDir){
    saveClientToFile(d).then(function(fname){
      if(fname){
        showToast('\ud83d\udcbe Guardado: '+fname,'');
        btn.textContent='\u2713 Guardado';btn.style.borderColor='#10B981';btn.style.color='#10B981';
        setTimeout(function(){btn.textContent=orig;btn.style.borderColor='';btn.style.color='';},2500);
      } else {
        showToast('Error al guardar el archivo','error');
      }
    });
  } else {
    showToast('\u26a0\ufe0f Conecta una carpeta para guardar el archivo','warn');
    btn.textContent='\u2713 Local';btn.style.borderColor='#F59E0B';btn.style.color='#F59E0B';
    setTimeout(function(){btn.textContent=orig;btn.style.borderColor='';btn.style.color='';},2500);
  }
}

// ─── RAMO MODAL ──────────────────────────────────────────────────
function openRamoModal(type, key) {
  if (!RAMO_CONFIG[type] || !RAMO_CONFIG[type].hasModal) return;
  currentRamoKey = key;
  currentRamoType = type;
  var d = ramoData[key] || {};
  var num = key.split('-').pop();
  document.getElementById('rmIcon').textContent = RAMO_CONFIG[type].emoji;
  document.getElementById('rmTitle').textContent = RAMO_CONFIG[type].name + (parseInt(num)>1?' #'+num:'');
  document.getElementById('rmBody').innerHTML = buildModalBody(type, d);
  document.getElementById('ramoOverlay').classList.add('show');
}
function getMissingFieldLabels(type){
  var missing=[];
  var d={};
  // Read current modal DOM state (without saving)
  if(type==='hogar'){
    var sw=document.querySelectorAll('#ramoOverlay .prop-switcher');
    if(!sw[0]||!sw[0].querySelector('.prop-btn.active')) missing.push('Tipo de inmueble');
    var pi=document.querySelector('.p-icons[data-field="hogar-p"]');
    if(!pi||!parseInt(pi.getAttribute('data-value'))) missing.push('Nº personas');
  } else if(type==='auto'||type==='moto'){
    var mat=document.getElementById('rm-matricula');if(!mat||!mat.value.trim()) missing.push('Matrícula');
    var mes=document.getElementById('rm-mes-vcto');if(!mes||!mes.value) missing.push('Mes vencimiento');
  } else if(type==='salud'){
    var modBtn=document.querySelector('#ramoOverlay .prop-btn.active');
    if(!modBtn) missing.push('Modalidad');
    var pi=document.querySelector('.p-icons[data-field="salud-p"]');
    if(!pi||!parseInt(pi.getAttribute('data-value'))) missing.push('Nº personas');
  } else if(type==='decesos'){
    var pi=document.querySelector('.p-icons[data-field="decesos-p"]');
    if(!pi||!parseInt(pi.getAttribute('data-value'))) missing.push('Nº personas');
  } else if(type==='vida'){
    var pi=document.querySelector('.p-icons[data-field="vida-p"]');
    var cap=document.getElementById('rm-vida-capital');
    if((!pi||!parseInt(pi.getAttribute('data-value')))&&(!cap||!cap.value.trim())) missing.push('Personas o Capital asegurado');
  } else if(type==='ahorro-g'||type==='ahorro-i'){
    var tGrp=document.querySelector('#ramoOverlay .yn-group[data-target="rm-tiene-ahorro-wrap"]');
    var qGrp=document.querySelector('#ramoOverlay .yn-group[data-target="rm-quiere-ahorrar-wrap"]');
    var tAnswered=tGrp&&(tGrp.querySelector('.yn-btn.yes')||tGrp.querySelector('.yn-btn.no'));
    var qAnswered=qGrp&&(qGrp.querySelector('.yn-btn.yes')||qGrp.querySelector('.yn-btn.no'));
    if(!tAnswered) missing.push('¿Tiene ahorros?');
    if(!qAnswered) missing.push('¿Quiere ahorrar?');
  } else if(type==='comunidades'){
    var nc=document.getElementById('rm-nombre-com');if(!nc||!nc.value.trim()) missing.push('Nombre de la comunidad');
  } else if(type==='embarcaciones'){
    var esl=document.getElementById('rm-eslora');var nm=document.getElementById('rm-nmotores');
    if(!esl||!esl.value.trim()) missing.push('Eslora');
    if(!nm||!nm.value) missing.push('Nº motores');
  } else if(type==='otro'){
    var nt=document.getElementById('rm-notas');if(!nt||!nt.value.trim()) missing.push('Anotaciones');
  }
  return missing;
}

function closeRamoModalWithCheck(){
  if(!currentRamoType) { closeRamoModal(); return; }
  var typeBase=currentRamoType.replace(/-\d+$/,'');
  if(!RAMO_CONFIG[typeBase]||!RAMO_CONFIG[typeBase].hasModal){ closeRamoModal(); return; }
  var missing=getMissingFieldLabels(typeBase);
  if(!missing.length){ closeRamoModal(); return; }
  var msg='<strong>Campos sin completar:</strong><ul style="margin:8px 0 0;padding-left:18px">';
  missing.forEach(function(m){msg+='<li>'+m+'</li>';});
  msg+='</ul>';
  document.getElementById('ramoWarnMsg').innerHTML=msg;
  var wd=document.getElementById('ramoCloseWarn');
  wd.style.display='flex';
}
function ramoWarnContinue(){
  document.getElementById('ramoCloseWarn').style.display='none';
  closeRamoModal();
}
function ramoWarnReview(){
  document.getElementById('ramoCloseWarn').style.display='none';
  // Focus first missing-field element
  var typeBase=currentRamoType?currentRamoType.replace(/-\d+$/,''):'';
  var firstFocus=null;
  if(typeBase==='hogar'){firstFocus=document.querySelector('#ramoOverlay .prop-switcher .prop-btn')||null;}
  else if(typeBase==='auto'||typeBase==='moto'){
    var mat=document.getElementById('rm-matricula');
    firstFocus=(!mat||!mat.value.trim())?mat:(document.getElementById('rm-mes-vcto'));
  } else if(typeBase==='salud'){
    firstFocus=document.querySelector('#ramoOverlay .prop-btn')||document.querySelector('.p-icons[data-field="salud-p"] .p-icon');
  } else if(typeBase==='decesos'){
    firstFocus=document.querySelector('.p-icons[data-field="decesos-p"] .p-icon');
  } else if(typeBase==='vida'){
    firstFocus=document.querySelector('.p-icons[data-field="vida-p"] .p-icon');
  } else if(typeBase==='ahorro-g'||typeBase==='ahorro-i'){
    firstFocus=document.querySelector('#ramoOverlay .yn-btn');
  } else if(typeBase==='comunidades'){
    firstFocus=document.getElementById('rm-nombre-com');
  } else if(typeBase==='embarcaciones'){
    var esl=document.getElementById('rm-eslora');firstFocus=(!esl||!esl.value.trim())?esl:document.getElementById('rm-nmotores');
  } else if(typeBase==='otro'){
    firstFocus=document.getElementById('rm-notas');
  }
  if(firstFocus&&firstFocus.focus) firstFocus.focus();
}
function closeRamoModal(){document.getElementById('ramoOverlay').classList.remove('show');currentRamoKey=null;currentRamoType=null;}
document.getElementById('ramoOverlay').addEventListener('click',function(e){if(e.target===this)closeRamoModalWithCheck();});

// ─── PERSON ICONS ────────────────────────────────────────────────
function buildPersonIcons(fid, n, maxIcons) {
  maxIcons = maxIcons || 6;
  var svg='<svg viewBox="0 0 28 32" fill="currentColor" xmlns="http://www.w3.org/2000/svg"><circle cx="14" cy="8.5" r="7"/><path d="M1 32c0-7.18 5.82-13 13-13s13 5.82 13 13"/></svg>';
  n = Math.min(n, maxIcons);
  var h='<div class="p-icons" data-field="'+fid+'" data-value="'+n+'">';
  for(var i=1;i<=maxIcons;i++) h+='<span class="p-icon'+(i<=n?' lit':'')+'" data-n="'+i+'" onclick="selectPersonCount(this)">'+svg+'</span>';
  return h+'</div>';
}
function selectPersonCount(icon) {
  var c=icon.closest('.p-icons');var n=parseInt(icon.getAttribute('data-n'));
  c.querySelectorAll('.p-icon').forEach(function(ic){ic.classList.toggle('lit',parseInt(ic.getAttribute('data-n'))<=n);});
  c.setAttribute('data-value',n);
  var fid=c.getAttribute('data-field');
  if(fid==='salud-p') updateSaludFechas(n);
  if(fid==='decesos-p') updateDecesosFechas(n);
  if(fid==='vida-p') updateVidaFechas(n);
}

function setPropType(btn){
  btn.closest('.prop-switcher').querySelectorAll('.prop-btn').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
}

// yn-group with optional data-target to show/hide a panel
function setYN(btn,val){
  var g=btn.closest('.yn-group');
  g.querySelectorAll('.yn-btn').forEach(function(b){b.classList.remove('yes','no');});
  if(val==='Si') btn.classList.add('yes'); else btn.classList.add('no');
  var tid=g.getAttribute('data-target');
  if(tid){var tw=document.getElementById(tid);if(tw)tw.style.display=val==='Si'?'':'none';}
}

function buildMesOptions(sel){
  var m=['','INMEDIATO','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  return m.map(function(x,i){var v=i===0?'':x;return '<option value="'+v+'"'+(v===(sel||'')?' selected':'')+'>'+(i===0?'\u2014 Mes \u2014':x)+'</option>';}).join('');
}

// ─── PER-PERSON FIELD BUILDERS ───────────────────────────────────
function buildPersonRowHtml(i, ex, extraFields) {
  var h='<div style="border:1px solid var(--border);border-radius:6px;padding:10px;margin-bottom:8px">';
  h+='<div style="font-size:10px;font-weight:700;color:var(--gray);margin-bottom:8px">Persona '+i+'</div>';
  h+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">';
  h+='<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">F. Nacimiento</span>';
  h+='<input class="rm-input" type="date" data-fnac="'+i+'" value="'+(ex.fnac||'')+'"></div>';
  h+='<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">Sexo</span>';
  h+='<select class="rm-input" data-sexo="'+i+'">';
  h+='<option value=""'+((!ex.sexo)?' selected':'')+'>&#8212;</option>';
  h+='<option value="H"'+(ex.sexo==='H'?' selected':'')+'>Hombre</option>';
  h+='<option value="M"'+(ex.sexo==='M'?' selected':'')+'>Mujer</option>';
  h+='</select></div>';
  if(extraFields){
    h+='<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">Profesi&#243;n</span>';
    h+='<input class="rm-input" type="text" data-prof="'+i+'" placeholder="Profesi&#243;n..." value="'+(ex.prof||'')+'"></div>';
    h+='<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">Deportes</span>';
    h+='<select class="rm-input" data-deporte="'+i+'" onchange="toggleDeporteCampo(this)">';
    h+='<option value=""'+((!ex.deporte)?' selected':'')+'>&#8212;</option>';
    h+='<option value="No"'+(ex.deporte==='No'?' selected':'')+'>No</option>';
    h+='<option value="Si"'+(ex.deporte==='Si'?' selected':'')+'>S&#237; (aficionado)</option>';
    h+='<option value="Competicion"'+(ex.deporte==='Competicion'?' selected':'')+'>Competici&#243;n</option>';
    h+='</select></div>';
    h+='<div id="rm-dep-campo-'+i+'" class="rm-field" style="margin:0;grid-column:span 2;'+(( ex.deporte==='Si'||ex.deporte==='Competicion')?'':'display:none')+'">';
    h+='<span class="rm-label" style="font-size:9px">Deporte practicado</span>';
    h+='<input class="rm-input" type="text" data-deportenombre="'+i+'" placeholder="Running, nataci&#243;n, ciclismo..." value="'+(ex.deporteNombre||'')+'"></div>';
  }
  h+='</div></div>';
  return h;
}

function toggleDeporteCampo(sel){
  var i=sel.getAttribute('data-deporte');
  var wrap=document.getElementById('rm-dep-campo-'+i);
  if(wrap)wrap.style.display=(sel.value==='Si'||sel.value==='Competicion')?'':'none';
}

function buildSaludFechasHtml(n, personasData){
  if(!n||n<1)return '';
  var h='';
  for(var i=1;i<=n;i++){var ex=(personasData&&personasData[i-1])||{};h+=buildPersonRowHtml(i,ex,true);}
  return h;
}
function buildDecesosFechasHtml(n, personasData){
  if(!n||n<1)return '';
  var h='';
  for(var i=1;i<=n;i++){var ex=(personasData&&personasData[i-1])||{};h+=buildPersonRowHtml(i,ex,false);}
  return h;
}

function updateSaludFechas(n){
  var wrap=document.getElementById('rm-fechas-wrap');if(!wrap)return;
  var existing=[];
  for(var i=1;i<=6;i++){
    var fnEl=wrap.querySelector('[data-fnac="'+i+'"]');
    var sxEl=wrap.querySelector('[data-sexo="'+i+'"]');
    var prEl=wrap.querySelector('[data-prof="'+i+'"]');
    var dpEl=wrap.querySelector('[data-deporte="'+i+'"]');
    if(fnEl)existing[i-1]={fnac:fnEl.value,sexo:sxEl?sxEl.value:'',prof:prEl?prEl.value:'',deporte:dpEl?dpEl.value:''};
  }
  wrap.innerHTML=buildSaludFechasHtml(n,existing);
}
function updateDecesosFechas(n){
  var wrap=document.getElementById('rm-decesos-fechas-wrap');if(!wrap)return;
  var existing=[];
  for(var i=1;i<=6;i++){
    var fnEl=wrap.querySelector('[data-fnac="'+i+'"]');
    var sxEl=wrap.querySelector('[data-sexo="'+i+'"]');
    if(fnEl)existing[i-1]={fnac:fnEl.value,sexo:sxEl?sxEl.value:''};
  }
  wrap.innerHTML=buildDecesosFechasHtml(n,existing);
}

function buildVidaFechasHtml(n, personasData){
  if(!n||n<1)return '';
  var h='';
  for(var i=1;i<=n;i++){var ex=(personasData&&personasData[i-1])||{};h+=buildPersonRowHtml(i,ex,false);}
  return h;
}
function updateVidaFechas(n){
  var wrap=document.getElementById('rm-vida-fechas-wrap');if(!wrap)return;
  var existing=[];
  for(var i=1;i<=4;i++){
    var fnEl=wrap.querySelector('[data-fnac="'+i+'"]');
    var sxEl=wrap.querySelector('[data-sexo="'+i+'"]');
    if(fnEl)existing[i-1]={fnac:fnEl.value,sexo:sxEl?sxEl.value:''};
  }
  wrap.innerHTML=buildVidaFechasHtml(n,existing);
}

// Persona block fields for autos (used in addAutoConductor)
function buildPersonBlockFields(prefix,carnetLabel,vals){
  var cl=carnetLabel||'F. Carnet B'; vals=vals||{};
  return '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:8px">'+
    '<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">Nombre</span><input class="rm-input" id="'+prefix+'-nombre" value="'+(vals.nombre||'')+'"></div>'+
    '<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">Primer apellido</span><input class="rm-input" id="'+prefix+'-ap1" value="'+(vals.ap1||'')+'"></div>'+
    '<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">Segundo apellido</span><input class="rm-input" id="'+prefix+'-ap2" value="'+(vals.ap2||'')+'"></div>'+
    '</div>'+
    '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px">'+
    '<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">DNI / NIE</span><input class="rm-input" id="'+prefix+'-dni" style="text-transform:uppercase" value="'+(vals.dni||'')+'"></div>'+
    '<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">F. Nacimiento</span><input class="rm-input" type="date" id="'+prefix+'-fnac" value="'+(vals.fnac||'')+'"></div>'+
    '<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">'+cl+'</span><input class="rm-input" type="date" id="'+prefix+'-fcarnet" value="'+(vals.fcarnet||'')+'"></div>'+
    '</div>';
}

// Auto section toggle (propietario / conductor)
function setAutoSection(which, btn, val){
  var sw=document.getElementById('rm-'+which+'-switcher');
  if(sw)sw.querySelectorAll('.prop-btn').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
  var fld=document.getElementById('rm-'+which+'-fields');
  if(fld)fld.style.display=(val==='otro')?'block':'none';
}

function addAutoConductor(){
  var c=document.getElementById('rm-otros-conductores');
  var n=c.querySelectorAll('.rm-conductor-block').length+1;
  var div=document.createElement('div');
  div.className='rm-conductor-block rm-block';
  div.innerHTML='<div class="rm-block-lbl">Conductor '+n+
    '<button class="rpa-btn rpa-del" type="button" style="margin-left:auto" onclick="this.closest(\'.rm-conductor-block\').remove()">\u2715</button></div>'+
    buildPersonBlockFields('rm-cond-'+n);
  c.appendChild(div);
}

function addSinco(){
  var c=document.getElementById('rm-sincos');
  var n=c.querySelectorAll('.rm-sinco-block').length+1;
  var div=document.createElement('div');
  div.className='rm-sinco-block rm-block';
  div.innerHTML='<div class="rm-block-lbl">SINCO '+n+
    (n>1?'<button class="rpa-btn rpa-del" type="button" style="margin-left:auto" onclick="this.closest(\'.rm-sinco-block\').remove()">\u2715</button>':'')+
    '</div>'+
    '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px">'+
    '<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">Compa\u00f1\u00eda</span><input class="rm-input rm-sinco-cia" placeholder="Nombre compa\u00f1\u00eda"></div>'+
    '<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">P\u00f3liza</span><input class="rm-input rm-sinco-poliza" placeholder="N\u00ba p\u00f3liza"></div>'+
    '<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">Matr\u00edcula</span><input class="rm-input rm-sinco-mat" placeholder="1234 ABC"></div>'+
    '</div>';
  c.appendChild(div);
}

function addMascotaDecesos(){
  var list=document.getElementById('rm-mascotas-list');
  var n=list.querySelectorAll('.rm-mascota-block').length+1;
  var div=document.createElement('div');
  div.className='rm-mascota-block rm-block';
  div.innerHTML='<div class="rm-block-lbl">Mascota '+n+
    '<button class="rpa-btn rpa-del" type="button" style="margin-left:auto" onclick="this.closest(\'.rm-mascota-block\').remove()">\u2715</button></div>'+
    '<div class="rm-field"><span class="rm-label" style="font-size:9px">Tipo</span>'+
    '<div class="prop-switcher rm-raza-sw">'+
    '<button class="prop-btn active" type="button" onclick="setMascotaTipo(this,\'mestizo\')">Mestizo</button>'+
    '<button class="prop-btn" type="button" onclick="setMascotaTipo(this,\'raza\')">Raza</button>'+
    '</div></div>'+
    '<div class="rm-raza-fields" style="display:none">'+
    '<div class="rm-field"><span class="rm-label" style="font-size:9px">\u00bfEs PPI?</span>'+
    '<div class="yn-group">'+
    '<button class="yn-btn" type="button" onclick="setMascotaPPI(this,\'Si\')">\u2713 S\u00ed</button>'+
    '<button class="yn-btn" type="button" onclick="setMascotaPPI(this,\'No\')">\u2717 No</button>'+
    '</div></div>'+
    '<div class="rm-pppi-fields" style="display:none">'+
    '<div class="rm-field"><span class="rm-label" style="font-size:9px">Raza PPPI</span>'+
    '<select class="rm-input rm-pppi-select">'+
    '<option value="">\u2014 Seleccionar raza \u2014</option>'+
    '<option>Pit Bull Terrier</option><option>Staffordshire Bull Terrier</option>'+
    '<option>American Staffordshire Terrier</option><option>Rottweiler</option>'+
    '<option>Dogo Argentino</option><option>Fila Brasileiro</option>'+
    '<option>Tosa Inu</option><option>Akita Inu</option>'+
    '</select></div>'+
    '<div class="rm-field"><span class="rm-label" style="font-size:9px">F. Nacimiento</span>'+
    '<input class="rm-input rm-masc-fnac" type="date"></div>'+
    '</div></div>';
  list.appendChild(div);
}

function setMascotaTipo(btn,tipo){
  var block=btn.closest('.rm-mascota-block');
  btn.closest('.rm-raza-sw').querySelectorAll('.prop-btn').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
  var rf=block.querySelector('.rm-raza-fields');if(rf)rf.style.display=tipo==='raza'?'block':'none';
}
function setMascotaPPI(btn,val){
  var g=btn.closest('.yn-group');
  g.querySelectorAll('.yn-btn').forEach(function(b){b.classList.remove('yes','no');});
  if(val==='Si')btn.classList.add('yes');else btn.classList.add('no');
  var block=btn.closest('.rm-mascota-block');
  var pf=block.querySelector('.rm-pppi-fields');if(pf)pf.style.display=val==='Si'?'block':'none';
}

// ─── BUILD MODAL BODY ────────────────────────────────────────────
function buildModalBody(type, d){
  d=d||{};
  var carnetOpts='<option value="">\u2014</option><option>AM/LCM</option><option>A1</option><option>A2</option><option>A</option><option>B</option><option>C</option>';

  switch(type){

    // ── HOGAR ──────────────────────────────────────────────────────
    case 'hogar':
      return '<div class="rm-grid2" style="margin-bottom:14px">'+
        '<div>'+
          '<div class="rm-field"><span class="rm-label">Tipo de uso</span>'+
            '<div class="prop-switcher">'+
              '<button class="prop-btn'+(d.tipoProp==='Propietario'?' active':'')+'" type="button" onclick="setPropType(this)">Propietario</button>'+
              '<button class="prop-btn'+(d.tipoProp==='Alquiler'?' active':'')+'" type="button" onclick="setPropType(this)">Alquiler</button>'+
              '<button class="prop-btn'+(d.tipoProp==='Inquilino'?' active':'')+'" type="button" onclick="setPropType(this)">Inquilino</button>'+
            '</div></div>'+
          '<div class="rm-field"><span class="rm-label">Tipo de vivienda</span>'+
            '<div class="prop-switcher" id="rm-tipo-viv">'+
              '<button class="prop-btn'+(d.tipoViv==='Principal'?' active':'')+'" type="button" onclick="setPropType(this)">Principal</button>'+
              '<button class="prop-btn'+(d.tipoViv==='Secundaria'?' active':'')+'" type="button" onclick="setPropType(this)">Secundaria</button>'+
              '<button class="prop-btn'+(d.tipoViv==='Tur\u00edstica'?' active':'')+'" type="button" onclick="setPropType(this)">Tur\u00edstica</button>'+
            '</div></div>'+
        '</div>'+
        '<div>'+
          '<div class="rm-field">'+buildPersonIcons('hogar-p',d.personas||0)+
            '<span class="rm-label" style="margin-top:6px">Personas que habitan</span></div>'+
          '<div class="rm-field"><span class="rm-label">Piscina</span>'+
            '<div class="yn-group">'+
              '<button class="yn-btn'+(d.piscina==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">\u2713 S\u00ed</button>'+
              '<button class="yn-btn'+(d.piscina==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">\u2717 No</button>'+
            '</div></div>'+
        '</div>'+
      '</div>'+
      '<div class="rm-field"><span class="rm-label">Hipoteca</span>'+
        '<div class="yn-group" data-target="rm-capital-wrap">'+
          '<button class="yn-btn'+(d.hipoteca==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">\u2713 S\u00ed</button>'+
          '<button class="yn-btn'+(d.hipoteca==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">\u2717 No</button>'+
        '</div></div>'+
      '<div class="rm-field" id="rm-capital-wrap" style="'+(d.hipoteca==='Si'?'':'display:none')+'">'+
        '<span class="rm-label">Capital pendiente (\u20ac)</span>'+
        '<input class="rm-input" type="text" inputmode="numeric" id="rm-capital" placeholder="150.000" oninput="numFmtInput(this)" value="'+(d.capital?numFmt(d.capital):'')+'">'+
      '</div>'+
      '<div class="divider" style="margin:12px 0"></div>'+
      '<div class="sec-tit" style="margin-bottom:8px">Direcci\u00f3n del inmueble <span style="font-weight:400;color:var(--gray2)">(cambiar si distinta)</span></div>'+
      '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:10px">'+
        '<div class="rm-field" style="grid-column:span 2;margin:0"><span class="rm-label" style="font-size:9px">Direcci\u00f3n</span><input class="rm-input" id="rm-dir-inmueble" placeholder="Calle, n\u00famero..." value="'+(d.dirInmueble||(document.getElementById('dir')?document.getElementById('dir').value:''))+'"></div>'+
        '<div class="rm-field" style="margin:0"><span class="rm-label" style="font-size:9px">CP</span><input class="rm-input" id="rm-cp-inmueble" placeholder="17600" value="'+(d.cpInmueble||(document.getElementById('cp')?document.getElementById('cp').value:''))+'"></div>'+
      '</div>'+
      '<div class="rm-field"><span class="rm-label" style="font-size:9px">Localidad</span><input class="rm-input" id="rm-loc-inmueble" placeholder="Localidad" value="'+(d.locInmueble||(document.getElementById('muni')?document.getElementById('muni').value:''))+'"></div>'+
        '<div class="rm-field" style="margin-bottom:8px"><span class="rm-label">Fecha de vencimiento</span><input class="rm-input" type="date" id="rm-vto" style="max-width:160px" value="'+(d.vto||'')+'"></div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="rm-field"><span class="rm-label">Observaciones</span>'+
          '<div class="yn-group" data-target="rm-hogar-obs-wrap">'+
            '<button class="yn-btn'+(d.observaciones==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">✓ Sí</button>'+
            '<button class="yn-btn'+(d.observaciones==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">✗ No</button>'+
          '</div>'+
          '<div id="rm-hogar-obs-wrap" style="'+(d.observaciones==='Si'?'':'display:none')+'">'+
            '<textarea class="rm-input rm-textarea" id="rm-hogar-obs" placeholder="Observaciones adicionales..." style="margin-top:6px;min-height:80px;resize:vertical">'+(d.observacionesNota||'')+'</textarea>'+
          '</div></div>';

    // ── AUTO / MOTO ────────────────────────────────────────────────
    case 'auto': case 'moto':
      var isAuto = type==='auto';
      var carnetLbl = isAuto ? 'F. Carnet B' : 'F. Carnet A1';
      var ownerVal = d.propietario||'tomador';
      var driverVal = d.conductor||'tomador';
      return '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px">'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Matr\u00edcula</span><input class="rm-input" id="rm-matricula" placeholder="1234 ABC" style="text-transform:uppercase" value="'+(d.matricula||'')+'"></div>'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Mes vencimiento</span><select class="rm-input" id="rm-mes-vcto">'+buildMesOptions(d.mesVcto)+'</select></div>'+
        '</div>'+
        '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:14px">'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Marca</span><input class="rm-input" id="rm-marca" placeholder="Volkswagen" value="'+(d.marca||'')+'"></div>'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Modelo</span><input class="rm-input" id="rm-modelo" placeholder="Golf" value="'+(d.modelo||'')+'"></div>'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Versi\u00f3n</span><input class="rm-input" id="rm-version" placeholder="1.5 TSI" value="'+(d.version||'')+'"></div>'+
        '</div>'+
        '<div class="divider" style="margin:10px 0"></div>'+
        '<div class="sec-tit" style="margin-bottom:8px">Propietario</div>'+
        '<div class="prop-switcher" id="rm-owner-switcher" style="margin-bottom:10px">'+
          '<button class="prop-btn'+(ownerVal==='tomador'?' active':'')+'" type="button" onclick="setAutoSection(\'owner\',this,\'tomador\')">Tomador</button>'+
          '<button class="prop-btn'+(ownerVal==='otro'?' active':'')+'" type="button" onclick="setAutoSection(\'owner\',this,\'otro\')">Otro</button>'+
        '</div>'+
        '<div id="rm-owner-fields" style="'+(ownerVal==='otro'?'':'display:none')+'">'+
          buildPersonBlockFields('rm-owner', carnetLbl, {nombre:d.ownerNombre,ap1:d.ownerAp1,ap2:d.ownerAp2,dni:d.ownerDni,fnac:d.ownerFnac,fcarnet:d.ownerFcarnet})+
        '</div>'+
        '<div class="divider" style="margin:10px 0"></div>'+
        '<div class="sec-tit" style="margin-bottom:8px">Conductor habitual</div>'+
        '<div class="prop-switcher" id="rm-driver-switcher" style="margin-bottom:10px">'+
          '<button class="prop-btn'+(driverVal==='tomador'?' active':'')+'" type="button" onclick="setAutoSection(\'driver\',this,\'tomador\')">Tomador</button>'+
          '<button class="prop-btn'+(driverVal==='propietario'?' active':'')+'" type="button" onclick="setAutoSection(\'driver\',this,\'propietario\')">Propietario</button>'+
          '<button class="prop-btn'+(driverVal==='otro'?' active':'')+'" type="button" onclick="setAutoSection(\'driver\',this,\'otro\')">Otro</button>'+
        '</div>'+
        '<div id="rm-driver-fields" style="'+(driverVal==='otro'?'':'display:none')+'">'+
          buildPersonBlockFields('rm-driver', carnetLbl, {nombre:d.driverNombre,ap1:d.driverAp1,ap2:d.driverAp2,dni:d.driverDni,fnac:d.driverFnac,fcarnet:d.driverFcarnet})+
        '</div>'+
        '<div class="divider" style="margin:10px 0"></div>'+
        '<div class="rm-section-hdr"><span class="sec-tit" style="margin:0">Otros conductores</span>'+
          '<button class="rpa-btn rpa-dup" type="button" onclick="addAutoConductor()">+ A\u00f1adir</button></div>'+
        '<div id="rm-otros-conductores" style="margin-top:8px"></div>'+
        '<div class="divider" style="margin:10px 0"></div>'+
        '<div class="rm-section-hdr"><span class="sec-tit" style="margin:0">SINCO (competencia)</span>'+
          '<button class="rpa-btn rpa-dup" type="button" onclick="addSinco()">+ A\u00f1adir</button></div>'+
        '<div id="rm-sincos" style="margin-top:8px"></div>'+
        '<script>setTimeout(function(){addSinco();},10);<\/script>'+
        '<div class="rm-field" style="margin-bottom:8px"><span class="rm-label">Fecha de vencimiento</span><input class="rm-input" type="date" id="rm-vto" style="max-width:160px" value="'+(d.vto||'')+'"></div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="rm-field"><span class="rm-label">Observaciones</span>'+
          '<div class="yn-group" data-target="rm-auto-obs-wrap">'+
            '<button class="yn-btn'+(d.observaciones==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">✓ Sí</button>'+
            '<button class="yn-btn'+(d.observaciones==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">✗ No</button>'+
          '</div>'+
          '<div id="rm-auto-obs-wrap" style="'+(d.observaciones==='Si'?'':'display:none')+'">'+
            '<textarea class="rm-input rm-textarea" id="rm-auto-obs" placeholder="Observaciones adicionales..." style="margin-top:6px;min-height:80px;resize:vertical">'+(d.observacionesNota||'')+'</textarea>'+
          '</div></div>';

    // ── SALUD ──────────────────────────────────────────────────────
    case 'salud':
      return '<div class="rm-field">'+
          '<span class="rm-label">Modalidad</span>'+
          '<div class="prop-switcher">'+
            '<button class="prop-btn'+(d.modalidad==='Cuadro m\u00e9dico'?' active':'')+'" type="button" onclick="setPropType(this)">Cuadro m\u00e9dico</button>'+
            '<button class="prop-btn'+(d.modalidad==='Reembolso'?' active':'')+'" type="button" onclick="setPropType(this)">Reembolso</button>'+
          '</div></div>'+
        '<div class="rm-field"><span class="rm-label">Copago</span>'+
          '<div class="yn-group">'+
            '<button class="yn-btn'+(d.copago==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">\u2713 S\u00ed</button>'+
            '<button class="yn-btn'+(d.copago==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">\u2717 No</button>'+
          '</div></div>'+
        '<div class="rm-field">'+buildPersonIcons('salud-p',d.personas||0)+
          '<span class="rm-label" style="margin-top:8px">Personas a asegurar</span></div>'+
        '<div id="rm-fechas-wrap">'+buildSaludFechasHtml(d.personas||0,d.personasData)+'</div>'+
        '<div class="rm-field" style="margin-bottom:8px"><span class="rm-label">Fecha de vencimiento</span><input class="rm-input" type="date" id="rm-vto" style="max-width:160px" value="'+(d.vto||'')+'"></div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="rm-field"><span class="rm-label">Observaciones</span>'+
          '<div class="yn-group" data-target="rm-salud-obs-wrap">'+
            '<button class="yn-btn'+(d.observaciones==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">\u2713 S\u00ed</button>'+
            '<button class="yn-btn'+(d.observaciones==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">\u2717 No</button>'+
          '</div>'+
          '<div id="rm-salud-obs-wrap" style="'+(d.observaciones==='Si'?'':'display:none')+'">'+
            '<textarea class="rm-input rm-textarea" id="rm-salud-obs" placeholder="Observaciones adicionales..." style="margin-top:6px;min-height:80px;resize:vertical">'+(d.observacionesNota||'')+'</textarea>'+
          '</div></div>';

    // ── DECESOS ────────────────────────────────────────────────────
    case 'decesos':
      return '<div class="rm-field">'+buildPersonIcons('decesos-p',d.personas||0)+
          '<span class="rm-label" style="margin-top:8px">Personas a asegurar</span></div>'+
        '<div id="rm-decesos-fechas-wrap">'+buildDecesosFechasHtml(d.personas||0,d.personasData)+'</div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="rm-section-hdr" style="margin-bottom:10px">'+
          '<span class="sec-tit" style="margin:0">Mascotas</span>'+
          '<button class="rpa-btn rpa-dup" type="button" onclick="addMascotaDecesos()">+ A\u00f1adir mascota</button>'+
        '</div>'+
        '<div id="rm-mascotas-list"></div>'+
        '<div class="rm-field" style="margin-bottom:8px"><span class="rm-label">Fecha de vencimiento</span><input class="rm-input" type="date" id="rm-vto" style="max-width:160px" value="'+(d.vto||'')+'"></div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="rm-field"><span class="rm-label">Observaciones</span>'+
          '<div class="yn-group" data-target="rm-decesos-obs-wrap">'+
            '<button class="yn-btn'+(d.observaciones==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">✓ Sí</button>'+
            '<button class="yn-btn'+(d.observaciones==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">✗ No</button>'+
          '</div>'+
          '<div id="rm-decesos-obs-wrap" style="'+(d.observaciones==='Si'?'':'display:none')+'">'+
            '<textarea class="rm-input rm-textarea" id="rm-decesos-obs" placeholder="Observaciones adicionales..." style="margin-top:6px;min-height:80px;resize:vertical">'+(d.observacionesNota||'')+'</textarea>'+
          '</div></div>';

    // ── VIDA ───────────────────────────────────────────────────────
    case 'vida':
      return '<div class="rm-field" style="margin-bottom:6px">'+buildPersonIcons('vida-p',d.personas||0,4)+
          '<span class="rm-label" style="margin-top:8px">Personas a asegurar (m\u00e1x. 4)</span></div>'+
        '<div id="rm-vida-fechas-wrap" style="margin-bottom:14px">'+buildVidaFechasHtml(d.personas||0,d.personasData)+'</div>'+
        '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px">'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Profesi\u00f3n</span>'+
            '<input class="rm-input" id="rm-vida-profesion" placeholder="M\u00e9dico, fontanero..." value="'+(d.profesion||'')+'"></div>'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Deportes (aficionado)</span>'+
            '<div class="yn-group" data-target="rm-vida-dep-campo">'+
              '<button class="yn-btn'+(d.deportes==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">\u2713 S\u00ed</button>'+
              '<button class="yn-btn'+(d.deportes==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">\u2717 No</button>'+
            '</div>'+
            '<div id="rm-vida-dep-campo" style="'+(d.deportes==='Si'?'':'display:none')+'">'+
              '<input class="rm-input" id="rm-vida-deporte-campo" placeholder="Running, nataci\u00f3n..." value="'+(d.deporteNombre||'')+'" style="margin-top:6px"></div>'+
          '</div>'+
        '</div>'+
        '<div class="rm-field"><span class="rm-label">Destino del capital</span>'+
          '<div style="display:flex;flex-direction:column;gap:8px;margin-top:4px">'+
            '<label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:12.5px">'+
              '<input type="checkbox" id="rm-destino-hipoteca"'+(d.destinoHipoteca?' checked':'')+' style="width:auto">'+
              'Hipotecas / Pr\u00e9stamos</label>'+
            '<label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:12.5px">'+
              '<input type="checkbox" id="rm-destino-sucesiones"'+(d.destinoSucesiones?' checked':'')+' style="width:auto">'+
              'Impuesto Sucesiones</label>'+
            '<label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:12.5px">'+
              '<input type="checkbox" id="rm-destino-familiar"'+(d.destinoFamiliar?' checked':'')+' style="width:auto">'+
              'Protecci\u00f3n Familiar</label>'+
          '</div></div>'+
        '<div class="rm-field"><span class="rm-label">Capital asegurado / Capital a contratar</span>'+
          '<input class="rm-input" type="text" inputmode="numeric" id="rm-vida-capital" placeholder="100.000" oninput="numFmtInput(this)" value="'+(d.capital?numFmt(d.capital):'')+'">'+'</div>'+
        '<div class="rm-field" style="margin-bottom:8px"><span class="rm-label">Fecha de vencimiento</span><input class="rm-input" type="date" id="rm-vto" style="max-width:160px" value="'+(d.vto||'')+'"></div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="rm-field"><span class="rm-label">Observaciones</span>'+
          '<div class="yn-group" data-target="rm-vida-obs-wrap">'+
            '<button class="yn-btn'+(d.observaciones==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">\u2713 S\u00ed</button>'+
            '<button class="yn-btn'+(d.observaciones==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">\u2717 No</button>'+
          '</div>'+
          '<div id="rm-vida-obs-wrap" style="'+(d.observaciones==='Si'?'':'display:none')+'">'+
            '<textarea class="rm-input rm-textarea" id="rm-vida-obs" placeholder="Observaciones adicionales..." style="margin-top:6px;min-height:80px;resize:vertical">'+(d.observacionesNota||'')+'</textarea>'+
          '</div></div>';

    // ── AHORRO ─────────────────────────────────────────────
    case 'ahorro-g': case 'ahorro-i':
      var isG = type==='ahorro-g';
      return '<div class="rm-field">'+
          '<div style="font-size:10px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;color:var(--gray2);margin-bottom:8px">'+(isG?'💰':'📈')+' Ahorro actual</div>'+
          '<div style="font-size:12px;color:var(--gray);margin-bottom:8px">¿El cliente tiene ahorros actualmente?</div>'+
          '<div class="yn-group" data-target="rm-tiene-ahorro-wrap">'+
            '<button class="yn-btn'+(d.tieneAhorros==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">✓ Sí</button>'+
            '<button class="yn-btn'+(d.tieneAhorros==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">✗ No</button>'+
          '</div>'+
          '<div id="rm-tiene-ahorro-wrap" style="'+(d.tieneAhorros==='Si'?'':'display:none')+'">'+
            '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px">'+
              '<div class="rm-field" style="margin:0"><span class="rm-label">Aportación mensual (€)</span><input class="rm-input" type="text" inputmode="numeric" id="rm-aportaciones" placeholder="200" oninput="numFmtInput(this)" value="'+(d.aportaciones?numFmt(d.aportaciones):'')+'"></div>'+
              '<div class="rm-field" style="margin:0"><span class="rm-label">Activo acumulado (€)</span><input class="rm-input" type="text" inputmode="numeric" id="rm-activo" placeholder="25.000" oninput="numFmtInput(this)" value="'+(d.activo?numFmt(d.activo):'')+'"></div>'+
            '</div>'+
            '<div class="rm-field" style="margin-top:10px"><span class="rm-label">Notas (ahorro actual)</span>'+
              '<textarea class="rm-input rm-textarea" id="rm-notas" placeholder="Tiene ahorro en banco, fondos, plazo fijo...">'+(d.notas||'')+'</textarea></div>'+
          '</div></div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="rm-field">'+
          '<div style="font-size:10px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;color:var(--gray2);margin-bottom:8px">🎯 Quiere ahorrar</div>'+
          '<div style="font-size:12px;color:var(--gray);margin-bottom:8px">¿El cliente quiere empezar a ahorrar?</div>'+
          '<div class="yn-group" data-target="rm-quiere-ahorrar-wrap">'+
            '<button class="yn-btn'+(d.quiereAhorrar==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">✓ Sí</button>'+
            '<button class="yn-btn'+(d.quiereAhorrar==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">✗ No</button>'+
          '</div>'+
          '<div id="rm-quiere-ahorrar-wrap" style="'+(d.quiereAhorrar==='Si'?'':'display:none')+'">'+
            '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px">'+
              '<div class="rm-field" style="margin:0"><span class="rm-label">Aportación mensual (€)</span><input class="rm-input" type="text" inputmode="numeric" id="rm-aportaciones-quiere" placeholder="200" oninput="numFmtInput(this)" value="'+(d.aportacionesQuiere?numFmt(d.aportacionesQuiere):'')+'"></div>'+
              '<div class="rm-field" style="margin:0"><span class="rm-label">Aportación inicial/única (€)</span><input class="rm-input" type="text" inputmode="numeric" id="rm-aportacion-inicial" placeholder="5.000" oninput="numFmtInput(this)" value="'+(d.aportacionInicial?numFmt(d.aportacionInicial):'')+'"></div>'+
            '</div>'+
            '<div class="rm-field" style="margin-top:10px"><span class="rm-label">Notas (objetivo de ahorro)</span>'+
              '<textarea class="rm-input rm-textarea" id="rm-notas-quiere" placeholder="Objetivo, plazo, perfil de riesgo...">'+(d.notasQuiere||'')+'</textarea></div>'+
          '</div></div>'+
        '<div class="rm-field" style="margin-bottom:8px"><span class="rm-label">Fecha de vencimiento</span><input class="rm-input" type="date" id="rm-vto" style="max-width:160px" value="'+(d.vto||'')+'"></div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="rm-field"><span class="rm-label">Observaciones</span>'+
          '<div class="yn-group" data-target="rm-ahorro-obs-wrap">'+
            '<button class="yn-btn'+(d.observaciones==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">✓ Sí</button>'+
            '<button class="yn-btn'+(d.observaciones==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">✗ No</button>'+
          '</div>'+
          '<div id="rm-ahorro-obs-wrap" style="'+(d.observaciones==='Si'?'':'display:none')+'">'+
            '<textarea class="rm-input rm-textarea" id="rm-ahorro-obs" placeholder="Observaciones adicionales..." style="margin-top:6px;min-height:80px;resize:vertical">'+(d.observacionesNota||'')+'</textarea>'+
          '</div></div>';

    // ── COMUNIDADES ────────────────────────────────────────────────
    case 'comunidades':
      return '<div class="rm-field"><span class="rm-label">Nombre de la comunidad</span>'+
          '<input class="rm-input" id="rm-nombre-com" placeholder="Comunidad Calle Mayor 15" value="'+(d.nombreCom||'')+'"></div>'+
        '<div class="rm-field"><span class="rm-label">Administrador</span>'+
          '<input class="rm-input" id="rm-admin" placeholder="Nombre del administrador de fincas" value="'+(d.admin||'')+'"></div>'+
        '<div class="rm-field" style="margin-bottom:8px"><span class="rm-label">Fecha de vencimiento</span><input class="rm-input" type="date" id="rm-vto" style="max-width:160px" value="'+(d.vto||'')+'"></div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="rm-field"><span class="rm-label">Observaciones</span>'+
          '<div class="yn-group" data-target="rm-comunidades-obs-wrap">'+
            '<button class="yn-btn'+(d.observaciones==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">✓ Sí</button>'+
            '<button class="yn-btn'+(d.observaciones==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">✗ No</button>'+
          '</div>'+
          '<div id="rm-comunidades-obs-wrap" style="'+(d.observaciones==='Si'?'':'display:none')+'">'+
            '<textarea class="rm-input rm-textarea" id="rm-comunidades-obs" placeholder="Observaciones adicionales..." style="margin-top:6px;min-height:80px;resize:vertical">'+(d.observacionesNota||'')+'</textarea>'+
          '</div></div>';

    // ── EMBARCACIONES ──────────────────────────────────────────────
    case 'embarcaciones':
      return '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px">'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Bandera / Pa\u00eds</span>'+
            '<input class="rm-input" id="rm-bandera" placeholder="Espa\u00f1a" value="'+(d.bandera||'')+'"></div>'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Zona de navegaci\u00f3n</span>'+
            '<select class="rm-input" id="rm-zona">'+
              '<option value=""'+((!d.zona)?' selected':'')+'>&#8212; Seleccionar &#8212;</option>'+
              '<option value="6a"'+(d.zona==='6a'?' selected':'')+'>6\u00aa Lista &#8212; Zona 1 (0&#8211;12 nm)</option>'+
              '<option value="7a"'+(d.zona==='7a'?' selected':'')+'>7\u00aa Lista &#8212; Zona 2 (alta mar)</option>'+
            '</select></div>'+
        '</div>'+
        '<div class="rm-field"><span class="rm-label">Material (casco)</span>'+
          '<div class="prop-switcher" id="rm-mat-casco">'+
            '<button class="prop-btn'+(d.materialCasco==='Fibra'?' active':'')+'" type="button" onclick="setPropType(this)">Fibra</button>'+
            '<button class="prop-btn'+(d.materialCasco==='Madera'?' active':'')+'" type="button" onclick="setPropType(this)">Madera</button>'+
            '<button class="prop-btn'+(d.materialCasco==='Metal'?' active':'')+'" type="button" onclick="setPropType(this)">Metal</button>'+
            '<button class="prop-btn'+(d.materialCasco==='Goma'?' active':'')+'" type="button" onclick="setPropType(this)">Goma</button>'+
          '</div></div>'+
        '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:12px">'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Eslora (m)</span>'+
            '<input class="rm-input" type="number" id="rm-eslora" placeholder="8.5" step="0.1" value="'+(d.eslora||'')+'"></div>'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">N\u00ba motores</span>'+
            '<select class="rm-input" id="rm-nmotores">'+
              '<option value=""'+((!d.nmotores)?' selected':'')+'>&#8212;</option>'+
              '<option value="1"'+(d.nmotores==='1'?' selected':'')+'>1</option>'+
              '<option value="2"'+(d.nmotores==='2'?' selected':'')+'>2</option>'+
              '<option value="3"'+(d.nmotores==='3'?' selected':'')+'>3</option>'+
            '</select></div>'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Potencia / motor (CV)</span>'+
            '<input class="rm-input" type="number" id="rm-potencia" placeholder="150" value="'+(d.potencia||'')+'"></div>'+
        '</div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="sec-tit" style="margin-bottom:10px">Da\u00f1os propios</div>'+
        '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px">'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Valor casco + motor (\u20ac)</span>'+
            '<input class="rm-input" type="text" inputmode="numeric" id="rm-valor-casco" placeholder="25.000" oninput="numFmtInput(this)" value="'+(d.valorCasco?numFmt(d.valorCasco):'')+'"></div>'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Valor accesorios (\u20ac)</span>'+
            '<input class="rm-input" type="text" inputmode="numeric" id="rm-valor-acc" placeholder="2.000" oninput="numFmtInput(this)" value="'+(d.valorAcc?numFmt(d.valorAcc):'')+'"></div>'+
        '</div>'+
        '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">A\u00f1o construcci\u00f3n</span>'+
            '<input class="rm-input" type="number" id="rm-anio-const" placeholder="2018" value="'+(d.anioConst||'')+'"></div>'+
          '<div class="rm-field" style="margin:0"><span class="rm-label">Material construcci\u00f3n</span>'+
            '<div class="prop-switcher" id="rm-mat-const">'+
              '<button class="prop-btn'+(d.materialConst==='Fibra'?' active':'')+'" type="button" onclick="setPropType(this)">Fibra</button>'+
              '<button class="prop-btn'+(d.materialConst==='Madera'?' active':'')+'" type="button" onclick="setPropType(this)">Madera</button>'+
              '<button class="prop-btn'+(d.materialConst==='Metal'?' active':'')+'" type="button" onclick="setPropType(this)">Metal</button>'+
              '<button class="prop-btn'+(d.materialConst==='Goma'?' active':'')+'" type="button" onclick="setPropType(this)">Goma</button>'+
            '</div></div>'+
        '</div>'+
        '<div class="rm-field" style="margin-bottom:8px"><span class="rm-label">Fecha de vencimiento</span><input class="rm-input" type="date" id="rm-vto" style="max-width:160px" value="'+(d.vto||'')+'"></div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="rm-field"><span class="rm-label">Observaciones</span>'+
          '<div class="yn-group" data-target="rm-embarcaciones-obs-wrap">'+
            '<button class="yn-btn'+(d.observaciones==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">✓ Sí</button>'+
            '<button class="yn-btn'+(d.observaciones==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">✗ No</button>'+
          '</div>'+
          '<div id="rm-embarcaciones-obs-wrap" style="'+(d.observaciones==='Si'?'':'display:none')+'">'+
            '<textarea class="rm-input rm-textarea" id="rm-embarcaciones-obs" placeholder="Observaciones adicionales..." style="margin-top:6px;min-height:80px;resize:vertical">'+(d.observacionesNota||'')+'</textarea>'+
          '</div></div>';

    // ── OTRO ───────────────────────────────────────────────────────
    case 'otro':
      return '<div class="rm-field"><span class="rm-label">Ramo / Producto</span>'+
          '<input class="rm-input" id="rm-otro-nombre" placeholder="Nombre del ramo o producto..." value="'+(d.nombre||'')+'"></div>'+
        '<div class="rm-field"><span class="rm-label">Anotaciones</span>'+
          '<textarea class="rm-input rm-textarea" id="rm-notas" placeholder="Notas libres sobre este seguro, vencimiento, compañía..." style="min-height:100px">'+(d.notas||'')+'</textarea></div>'+
        '<div class="rm-field" style="margin-bottom:8px"><span class="rm-label">Fecha de vencimiento</span><input class="rm-input" type="date" id="rm-vto" style="max-width:160px" value="'+(d.vto||'')+'"></div>'+
        '<div class="divider" style="margin:12px 0"></div>'+
        '<div class="rm-field"><span class="rm-label">Observaciones</span>'+
          '<div class="yn-group" data-target="rm-otro-obs-wrap">'+
            '<button class="yn-btn'+(d.observaciones==='Si'?' yes':'')+'" type="button" onclick="setYN(this,\'Si\')">✓ Sí</button>'+
            '<button class="yn-btn'+(d.observaciones==='No'?' no':'')+'" type="button" onclick="setYN(this,\'No\')">✗ No</button>'+
          '</div>'+
          '<div id="rm-otro-obs-wrap" style="'+(d.observaciones==='Si'?'':'display:none')+'">'+
            '<textarea class="rm-input rm-textarea" id="rm-otro-obs" placeholder="Observaciones adicionales..." style="margin-top:6px;min-height:80px;resize:vertical">'+(d.observacionesNota||'')+'</textarea>'+
          '</div></div>';

    default:
      return '<p style="color:var(--gray2);font-size:13px">Sin datos adicionales para este ramo.</p>';
  }
}

// ─── SAVE RAMO DATA ───────────────────────────────────────────────
function saveRamoData(){
  if(!currentRamoKey)return;
  // Use stored type (avoids key-parsing failures for pills with grid-prefixed keys like 'otro')
  var type = currentRamoType || (function(){var p=currentRamoKey.split('-');p.pop();return p.join('-');})();
  var g=function(id){var el=document.getElementById(id);return el?el.value.trim():'';};
  var gn=function(id){var el=document.getElementById(id);return el?el.value.trim().replace(/\./g,''):'';};  // strips thousands dots
  var d={};

  if(type==='hogar'){
    var switchers=document.querySelectorAll('#ramoOverlay .prop-switcher');
    d.tipoProp=switchers[0]?(switchers[0].querySelector('.prop-btn.active')||{textContent:''}).textContent.trim():'';
    d.tipoViv=document.getElementById('rm-tipo-viv')?(document.querySelector('#rm-tipo-viv .prop-btn.active')||{textContent:''}).textContent.trim():'';
    var piscinaYes=document.querySelector('#ramoOverlay .yn-group:not([data-target]) .yn-btn.yes');
    d.piscina=piscinaYes?'Si':(document.querySelector('#ramoOverlay .yn-group:not([data-target]) .yn-btn.no')?'No':'');
    var hipoGroup=document.querySelector('#ramoOverlay .yn-group[data-target="rm-capital-wrap"]');
    d.hipoteca=hipoGroup?(hipoGroup.querySelector('.yn-btn.yes')?'Si':(hipoGroup.querySelector('.yn-btn.no')?'No':'')):'';
    d.capital=gn('rm-capital');
    var pi=document.querySelector('.p-icons[data-field="hogar-p"]');d.personas=pi?parseInt(pi.getAttribute('data-value'))||0:0;
    d.dirInmueble=g('rm-dir-inmueble');d.cpInmueble=g('rm-cp-inmueble');d.locInmueble=g('rm-loc-inmueble');
    var obsGrp_hogar=document.querySelector('#ramoOverlay .yn-group[data-target="rm-hogar-obs-wrap"]');
    d.observaciones=obsGrp_hogar?(obsGrp_hogar.querySelector('.yn-btn.yes')?'Si':(obsGrp_hogar.querySelector('.yn-btn.no')?'No':'')):''; 
    d.observacionesNota=g('rm-hogar-obs');
    d.vto=g('rm-vto');

  }else if(type==='auto'||type==='moto'){
    d.matricula=g('rm-matricula');d.mesVcto=g('rm-mes-vcto');
    d.marca=g('rm-marca');d.modelo=g('rm-modelo');d.version=g('rm-version');
    var owSw=document.getElementById('rm-owner-switcher');
    d.propietario=owSw?(owSw.querySelector('.prop-btn.active')||{textContent:'tomador'}).textContent.trim().toLowerCase():'tomador';
    if(d.propietario==='otro'){
      d.ownerNombre=g('rm-owner-nombre');d.ownerAp1=g('rm-owner-ap1');d.ownerAp2=g('rm-owner-ap2');
      d.ownerDni=g('rm-owner-dni');d.ownerFnac=g('rm-owner-fnac');d.ownerFcarnet=g('rm-owner-fcarnet');
    }
    var drSw=document.getElementById('rm-driver-switcher');
    d.conductor=drSw?(drSw.querySelector('.prop-btn.active')||{textContent:'tomador'}).textContent.trim().toLowerCase():'tomador';
    if(d.conductor==='otro'){
      d.driverNombre=g('rm-driver-nombre');d.driverAp1=g('rm-driver-ap1');d.driverAp2=g('rm-driver-ap2');
      d.driverDni=g('rm-driver-dni');d.driverFnac=g('rm-driver-fnac');d.driverFcarnet=g('rm-driver-fcarnet');
    }
    d.sincos=[];
    document.querySelectorAll('#rm-sincos .rm-sinco-block').forEach(function(blk){
      var cia=blk.querySelector('.rm-sinco-cia');var pol=blk.querySelector('.rm-sinco-poliza');var mat=blk.querySelector('.rm-sinco-mat');
      d.sincos.push({cia:cia?cia.value:'',poliza:pol?pol.value:'',matricula:mat?mat.value:''});
    });
    var obsGrp_auto=document.querySelector('#ramoOverlay .yn-group[data-target="rm-auto-obs-wrap"]');
    d.observaciones=obsGrp_auto?(obsGrp_auto.querySelector('.yn-btn.yes')?'Si':(obsGrp_auto.querySelector('.yn-btn.no')?'No':'')):''; 
    d.observacionesNota=g('rm-auto-obs');
    d.vto=g('rm-vto');

  }else if(type==='salud'){
    var modBtn=document.querySelector('#ramoOverlay .prop-btn.active');d.modalidad=modBtn?modBtn.textContent.trim():'';
    var copGrp=document.querySelector('#ramoOverlay .yn-group');
    d.copago=copGrp?(copGrp.querySelector('.yn-btn.yes')?'Si':(copGrp.querySelector('.yn-btn.no')?'No':'')):'';
    var pi=document.querySelector('.p-icons[data-field="salud-p"]');d.personas=pi?parseInt(pi.getAttribute('data-value'))||0:0;
    d.personasData=[];
    var wrap=document.getElementById('rm-fechas-wrap');
    if(wrap){for(var i=1;i<=d.personas;i++){
      var fnEl=wrap.querySelector('[data-fnac="'+i+'"]');var sxEl=wrap.querySelector('[data-sexo="'+i+'"]');
      var prEl=wrap.querySelector('[data-prof="'+i+'"]');var dpEl=wrap.querySelector('[data-deporte="'+i+'"]');
      var dnEl=wrap.querySelector('[data-deportenombre="'+i+'"');d.personasData.push({fnac:fnEl?fnEl.value:'',sexo:sxEl?sxEl.value:'',prof:prEl?prEl.value:'',deporte:dpEl?dpEl.value:'',deporteNombre:dnEl?dnEl.value:''});
    }}
    var saludObs=document.querySelector('#ramoOverlay .yn-group[data-target="rm-salud-obs-wrap"]');
    d.observaciones=saludObs?(saludObs.querySelector('.yn-btn.yes')?'Si':(saludObs.querySelector('.yn-btn.no')?'No':'')):''; 
    d.observacionesNota=g('rm-salud-obs');
    d.vto=g('rm-vto');

  }else if(type==='decesos'){
    var pi=document.querySelector('.p-icons[data-field="decesos-p"]');d.personas=pi?parseInt(pi.getAttribute('data-value'))||0:0;
    d.personasData=[];
    var wrap=document.getElementById('rm-decesos-fechas-wrap');
    if(wrap){for(var i=1;i<=d.personas;i++){
      var fnEl=wrap.querySelector('[data-fnac="'+i+'"]');var sxEl=wrap.querySelector('[data-sexo="'+i+'"]');
      d.personasData.push({fnac:fnEl?fnEl.value:'',sexo:sxEl?sxEl.value:''});
    }}
    d.mascotas=[];
    document.querySelectorAll('#rm-mascotas-list .rm-mascota-block').forEach(function(blk){
      var razaBtn=blk.querySelector('.rm-raza-sw .prop-btn.active');
      var tipoMasc=razaBtn?razaBtn.textContent.trim():'Mestizo';
      var ppiYes=blk.querySelector('.yn-btn.yes');var ppiNo=blk.querySelector('.yn-btn.no');
      var ppi=ppiYes?'Si':(ppiNo?'No':'');
      var pppiSel=blk.querySelector('.rm-pppi-select');var pppiRaza=pppiSel?pppiSel.value:'';
      var fnacEl=blk.querySelector('.rm-masc-fnac');var fnac=fnacEl?fnacEl.value:'';
      d.mascotas.push({tipo:tipoMasc,ppi:ppi,pppiRaza:pppiRaza,fnac:fnac});
    });
    var obsGrp_decesos=document.querySelector('#ramoOverlay .yn-group[data-target="rm-decesos-obs-wrap"]');
    d.observaciones=obsGrp_decesos?(obsGrp_decesos.querySelector('.yn-btn.yes')?'Si':(obsGrp_decesos.querySelector('.yn-btn.no')?'No':'')):''; 
    d.observacionesNota=g('rm-decesos-obs');
    d.vto=g('rm-vto');

  }else if(type==='vida'){
    var pi=document.querySelector('.p-icons[data-field="vida-p"]');d.personas=pi?parseInt(pi.getAttribute('data-value'))||0:0;
    d.profesion=g('rm-vida-profesion');d.capital=gn('rm-vida-capital');
    var depGrp=document.querySelector('#ramoOverlay .yn-group[data-target="rm-vida-dep-campo"]');
    d.deportes=depGrp?(depGrp.querySelector('.yn-btn.yes')?'Si':(depGrp.querySelector('.yn-btn.no')?'No':'')):'';
    d.deporteNombre=g('rm-vida-deporte-campo');
    d.personasData=[];
    var vidaWrap=document.getElementById('rm-vida-fechas-wrap');
    if(vidaWrap){for(var i=1;i<=Math.min(d.personas,4);i++){
      var fnEl=vidaWrap.querySelector('[data-fnac="'+i+'"]');var sxEl=vidaWrap.querySelector('[data-sexo="'+i+'"]');
      d.personasData.push({fnac:fnEl?fnEl.value:'',sexo:sxEl?sxEl.value:''});
    }}
    var ckH=document.getElementById('rm-destino-hipoteca');d.destinoHipoteca=ckH?ckH.checked:false;
    var ckS=document.getElementById('rm-destino-sucesiones');d.destinoSucesiones=ckS?ckS.checked:false;
    var ckF=document.getElementById('rm-destino-familiar');d.destinoFamiliar=ckF?ckF.checked:false;
    var obsGrp=document.querySelector('#ramoOverlay .yn-group[data-target="rm-vida-obs-wrap"]');
    d.observaciones=obsGrp?(obsGrp.querySelector('.yn-btn.yes')?'Si':(obsGrp.querySelector('.yn-btn.no')?'No':'')):'';
    d.observacionesNota=g('rm-vida-obs');
    d.vto=g('rm-vto');

  }else if(type==='ahorro-g'||type==='ahorro-i'){
    var tieneGrp=document.querySelector('#ramoOverlay .yn-group[data-target="rm-tiene-ahorro-wrap"]');
    d.tieneAhorros=tieneGrp?(tieneGrp.querySelector('.yn-btn.yes')?'Si':(tieneGrp.querySelector('.yn-btn.no')?'No':'')):''; 
    d.aportaciones=gn('rm-aportaciones');d.activo=gn('rm-activo');d.notas=g('rm-notas');
    var quiereGrp=document.querySelector('#ramoOverlay .yn-group[data-target="rm-quiere-ahorrar-wrap"]');
    d.quiereAhorrar=quiereGrp?(quiereGrp.querySelector('.yn-btn.yes')?'Si':(quiereGrp.querySelector('.yn-btn.no')?'No':'')):''; 
    d.aportacionesQuiere=gn('rm-aportaciones-quiere');d.aportacionInicial=gn('rm-aportacion-inicial');d.notasQuiere=g('rm-notas-quiere');
    var obsGrp_ahorro=document.querySelector('#ramoOverlay .yn-group[data-target="rm-ahorro-obs-wrap"]');
    d.observaciones=obsGrp_ahorro?(obsGrp_ahorro.querySelector('.yn-btn.yes')?'Si':(obsGrp_ahorro.querySelector('.yn-btn.no')?'No':'')):''; 
    d.observacionesNota=g('rm-ahorro-obs');
    d.vto=g('rm-vto');

  }else if(type==='comunidades'){
    d.nombreCom=g('rm-nombre-com');d.admin=g('rm-admin');
    var obsGrp_comunidades=document.querySelector('#ramoOverlay .yn-group[data-target="rm-comunidades-obs-wrap"]');
    d.observaciones=obsGrp_comunidades?(obsGrp_comunidades.querySelector('.yn-btn.yes')?'Si':(obsGrp_comunidades.querySelector('.yn-btn.no')?'No':'')):''; 
    d.observacionesNota=g('rm-comunidades-obs');
    d.vto=g('rm-vto');

  }else if(type==='otro'){
    d.nombre=g('rm-otro-nombre');d.notas=g('rm-notas');
    var obsGrp_otro=document.querySelector('#ramoOverlay .yn-group[data-target="rm-otro-obs-wrap"]');
    d.observaciones=obsGrp_otro?(obsGrp_otro.querySelector('.yn-btn.yes')?'Si':(obsGrp_otro.querySelector('.yn-btn.no')?'No':'')):''; 
    d.observacionesNota=g('rm-otro-obs');
    d.vto=g('rm-vto');

  }else if(type==='embarcaciones'){
    d.bandera=g('rm-bandera');d.zona=g('rm-zona');
    var mcBtn=document.querySelector('#rm-mat-casco .prop-btn.active');d.materialCasco=mcBtn?mcBtn.textContent.trim():'';
    d.eslora=g('rm-eslora');d.nmotores=g('rm-nmotores');d.potencia=g('rm-potencia');
    d.valorCasco=gn('rm-valor-casco');d.valorAcc=gn('rm-valor-acc');d.anioConst=g('rm-anio-const');
    var mkBtn=document.querySelector('#rm-mat-const .prop-btn.active');d.materialConst=mkBtn?mkBtn.textContent.trim():'';
    var obsGrp_embarcaciones=document.querySelector('#ramoOverlay .yn-group[data-target="rm-embarcaciones-obs-wrap"]');
    d.observaciones=obsGrp_embarcaciones?(obsGrp_embarcaciones.querySelector('.yn-btn.yes')?'Si':(obsGrp_embarcaciones.querySelector('.yn-btn.no')?'No':'')):''; 
    d.observacionesNota=g('rm-embarcaciones-obs');
    d.vto=g('rm-vto');

  }

  // Merge preserving pill state
  var prev=ramoData[currentRamoKey]||{};
  if(prev.estado)d.estado=prev.estado;
  d._saved=true;
  ramoData[currentRamoKey]=d;
  var pill=document.querySelector('[data-key="'+currentRamoKey+'"]');
  // Update Otro pill label with custom ramo name
  if(type==='otro' && d.nombre){
    var nameSpan=pill?pill.querySelector('.rp-name'):null;
    if(nameSpan) nameSpan.textContent=d.nombre;
  }
  if(pill){var cb=pill.querySelector('input[name="seg"]');if(cb)cb.checked=true;}
  refreshPillBadge(currentRamoKey);
  closeRamoModal();
}

// ─── FILE SYSTEM ACCESS + INDEXEDDB ──────────────────────────────
var clientesDir = null;
var allClients  = [];
var searchDebounce = null;

function openIDB() {
  return new Promise(function(resolve, reject) {
    var req = indexedDB.open('optimizarte-alta-v1', 1);
    req.onupgradeneeded = function(e) { e.target.result.createObjectStore('handles'); };
    req.onsuccess = function(e) { resolve(e.target.result); };
    req.onerror   = function(e) { reject(e.target.error); };
  });
}
function idbPut(key, val) {
  return openIDB().then(function(db) {
    return new Promise(function(resolve, reject) {
      var tx = db.transaction('handles', 'readwrite');
      tx.objectStore('handles').put(val, key);
      tx.oncomplete = resolve; tx.onerror = function(e){reject(e.target.error);};
    });
  });
}
function idbGet(key) {
  return openIDB().then(function(db) {
    return new Promise(function(resolve, reject) {
      var tx  = db.transaction('handles', 'readonly');
      var req = tx.objectStore('handles').get(key);
      req.onsuccess = function(e) { resolve(e.target.result); };
      req.onerror   = function(e) { reject(e.target.error); };
    });
  });
}

// ─── REPORT SAVE DIRECTORY ───────────────────────────────────────
var reportesDirHandle = null; // cached in memory once acquired

async function getReportesDir() {
  // Return cached handle if still valid
  if (reportesDirHandle) {
    try { var p = await reportesDirHandle.requestPermission({mode:'readwrite'}); if(p==='granted') return reportesDirHandle; } catch(e){}
    reportesDirHandle = null;
  }
  // Try IndexedDB
  try {
    var h = await idbGet('reportesDir');
    if (h) {
      var perm = await h.requestPermission({mode:'readwrite'});
      if (perm === 'granted') { reportesDirHandle = h; return h; }
    }
  } catch(e) {}
  // Ask user to pick folder — inform them of the expected path
  if (!window.showDirectoryPicker) return null;
  try {
    showToast('📁 Selecciona la carpeta AltaClientesLocal donde se guardarán los informes','warn');
    var picked = await window.showDirectoryPicker({ id:'reportes-dir', mode:'readwrite' });
    await idbPut('reportesDir', picked);
    reportesDirHandle = picked;
    setReportesDirBtn(picked.name);
    return picked;
  } catch(e) {
    if (e.name !== 'AbortError') console.error('reportesDir pick error:', e);
    return null;
  }
}

function setReportesDirBtn(name) {
  var lbl = document.getElementById('reportesDirStatus');
  if (lbl) lbl.textContent = name ? ('📁 ' + name) : '';
}

async function saveReportToDir(blob, filename) {
  var dirHandle = await getReportesDir();
  if (!dirHandle) {
    // Fallback: browser blob download
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a'); a.href=url; a.download=filename; a.style.display='none';
    document.body.appendChild(a); a.click();
    setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(url);},1500);
    return null; // path unknown (Downloads)
  }
  try {
    var fh = await dirHandle.getFileHandle(filename, {create:true});
    var wr = await fh.createWritable();
    await wr.write(blob);
    await wr.close();
    return dirHandle.name + '/' + filename;
  } catch(e) {
    console.error('saveReportToDir error:', e);
    // Fallback to blob download
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a'); a.href=url; a.download=filename; a.style.display='none';
    document.body.appendChild(a); a.click();
    setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(url);},1500);
    return null;
  }
}

async function initReportesDirFromIDB() {
  if (!window.showDirectoryPicker) return;
  try {
    var h = await idbGet('reportesDir');
    if (!h) return;
    var perm = await h.queryPermission({mode:'readwrite'});
    if (perm === 'granted') { reportesDirHandle = h; setReportesDirBtn(h.name); }
  } catch(e) {}
}

async function selectClientesDir() {
  if (!window.showDirectoryPicker) {
    alert('Tu navegador no soporta acceso a carpetas locales.\nUsa Google Chrome o Microsoft Edge.');
    return;
  }
  try {
    clientesDir = await window.showDirectoryPicker({ id: 'clientes-dir', mode: 'readwrite' });
    await idbPut('clientesDir', clientesDir);
    setDirBtn(true, clientesDir.name);
    await refreshAllClients();
  } catch(e) { if (e.name !== 'AbortError') console.error(e); }
}

function setDirBtn(ok, name) {
  var btn = document.getElementById('dirBtn');
  var lbl = document.getElementById('dirStatus');
  if (!btn) return;
  btn.className = ok ? 'dir-btn connected' : 'dir-btn';
  if (lbl) lbl.textContent = ok ? (name || 'Conectado') : '';
}

async function initDirFromIDB() {
  if (!window.showDirectoryPicker) return;
  try {
    var h = await idbGet('clientesDir');
    if (!h) return;
    var perm = await h.requestPermission({ mode: 'readwrite' });
    if (perm === 'granted') {
      clientesDir = h;
      setDirBtn(true, h.name);
      await refreshAllClients();
    }
  } catch(e) { /* silent fail */ }
}

function formatDatetimeDisplay(iso) {
  if (!iso) return '—';
  var d = new Date(iso);
  var day = String(d.getDate()).padStart(2,'0');
  var mon = String(d.getMonth()+1).padStart(2,'0');
  var yr  = d.getFullYear();
  var hh  = String(d.getHours()).padStart(2,'0');
  var mm  = String(d.getMinutes()).padStart(2,'0');
  return day+'/'+mon+'/'+yr+' '+hh+':'+mm;
}
function setFormDates(creacion, modificacion) {
  var fc = document.getElementById('fecha-creacion');
  var fma = document.getElementById('fecha-modificacion-asigna');
  if (fc) fc.textContent = formatDatetimeDisplay(creacion);
  if (fma) fma.textContent = formatDatetimeDisplay(modificacion);
}
async function saveClientToFile(data) {
  if (!clientesDir) return null;
  try {
    var now = new Date();
    var pad = function(n){return n<10?'0'+n:''+n;};
    var nif = (data.nif_cif||data.par_nif||data.aut_nif||data.emp_cif||'').trim().toUpperCase();
    var fname = null;
    // Check for existing file with same DNI/NIE/CIF
    if (nif) {
      for await (var [name, handle] of clientesDir.entries()) {
        if (name.endsWith('.json') && handle.kind === 'file') {
          try {
            var existing = JSON.parse(await (await handle.getFile()).text());
            var existNif = (existing.nif_cif||existing.par_nif||existing.aut_nif||existing.emp_cif||'').trim().toUpperCase();
            if (existNif && existNif === nif) {
              fname = name; // overwrite this file
              showToast('\u267b\ufe0f Actualizando cliente existente: '+name,'warn');
              break;
            }
          } catch(e) {}
        }
      }
    }
    if (!fname) {
      var d = now.getFullYear()+''+pad(now.getMonth()+1)+pad(now.getDate());
      var t = pad(now.getHours())+pad(now.getMinutes())+pad(now.getSeconds());
      var nom = (data.nombre_completo||'cliente').replace(/[^a-zA-Z\u00C0-\u024F0-9 ]/g,'').replace(/ +/g,'_').slice(0,30);
      fname = d+'_'+t+'_'+nom+'.json';
    }
    var fh = await clientesDir.getFileHandle(fname, {create:true});
    var now = new Date().toISOString();
    if (!data.fechaCreacion) data.fechaCreacion = now;
    data.fechaModificacion = now;
    setFormDates(data.fechaCreacion, data.fechaModificacion);
    var wr = await fh.createWritable();
    await wr.write(JSON.stringify(data, null, 2));
    await wr.close();
    await refreshAllClients();
    return fname;
  } catch(e) { console.error('saveClientToFile:', e); return null; }
}

async function refreshAllClients() {
  if (!clientesDir) return;
  try {
    var list = [];
    for await (var [name, handle] of clientesDir.entries()) {
      if (name.endsWith('.json') && handle.kind === 'file') {
        try {
          var file = await handle.getFile();
          var d    = JSON.parse(await file.text());
          d._filename = name;
          d._mtime    = file.lastModified;
          list.push(d);
        } catch(e) {}
      }
    }
    allClients = list.sort(function(a,b){return (b._mtime||0)-(a._mtime||0);});
  } catch(e) {}
}

// ─── SEARCH ───────────────────────────────────────────────────────
// IMPORTANT: NO mostrar suggerencies només per fer focus al cercador.
// Les suggerencies només apareixen quan l'usuari ha escrit 3+ caràcters.
function onSearchFocus() {
  var inp = document.getElementById('searchInput');
  if (!inp) return;
  var val = (inp.value || '').trim();
  if (val.length >= 3) {
    // Si ja hi havia text vàlid (cas tornada al cercador), re-executem el filtre
    onSearchInput(inp.value);
  }
  // Si no, no fem res — esperem 3+ caràcters
}

function onSearchInput(val) {
  clearTimeout(searchDebounce);
  searchDebounce = setTimeout(function() {
    // Si esborra el text, reset complet
    if (!val.trim()) {
      hideSearchResults();
      resetFormulariNouClient();
      // Restaurar card 0 si tenia fitxa
      var cardTipo = document.getElementById('card-tipo-cliente');
      if (cardTipo && cardTipo._originalHTML) {
        cardTipo.innerHTML = cardTipo._originalHTML;
        cardTipo._originalHTML = null;
      }
      window._pendingClient = null;
      setFormMode('initial');
      return;
    }

    // Si l'usuari escriu nou text i hi havia fitxa al card 0 o estàvem en mode carregat → reset card 0
    var cardTipo = document.getElementById('card-tipo-cliente');
    if (cardTipo && cardTipo._originalHTML) {
      cardTipo.innerHTML = cardTipo._originalHTML;
      cardTipo._originalHTML = null;
      window._pendingClient = null;
    }
    // Si estàvem en mode carregat (directe/fitxa) i l'usuari torna a cercar → tornar a 'initial' visualment
    if (_currentFormMode === 'client-loaded-directe' || _currentFormMode === 'client-loaded-fitxa') {
      resetFormulariNouClient();
      setFormMode('initial');
    }

    // MÍNIM 3 CARÀCTERS per cercar
    if (val.trim().length < 3) {
      hideSearchResults();
      return;
    }

    var q = val.trim().toLowerCase();
    var res = allClients.filter(function(c) { return matchClient(c, q); }).slice(0,10);
    showSearchResults(res);
  }, 180);
}

// TASCA 2: Funció per resetar formulari a estat "nou client"
function resetFormulariNouClient() {
  // Netejar formulari (però mantenir col·laboradors)
  var form = document.getElementById('altaForm');
  if (form) {
    // Guardar col·laboradors actuals
    var colabRecogePrev = colabRecoge;
    var colabAsignaPrev = colabAsigna;
    
    form.reset();
    
    // Restaurar col·laboradors
    colabRecoge = colabRecogePrev;
    colabAsigna = colabAsignaPrev;
  }
  
  // Restaurar tipo a 'par' per defecte
  tipo = 'par';
  var parOpt = document.getElementById('tipo-par');
  if (parOpt) {
    document.querySelectorAll('.tipo-opt').forEach(function(o) {
      o.classList.remove('active');
      var inp = o.querySelector('input');
      if (inp) inp.checked = false;
    });
    parOpt.classList.add('active');
    var parInp = parOpt.querySelector('input');
    if (parInp) parInp.checked = true;
  }
  
  // Reset vencimientos
  vtosPerSi = false;
  var vsi = document.getElementById('vper-si');
  var vno = document.getElementById('vper-no');
  if (vsi) vsi.classList.remove('si','no');
  if (vno) vno.classList.remove('si','no');
  
  // Amagar banner "Cliente cargado"
  var banner = document.getElementById('loadBanner');
  if (banner) banner.style.display = 'none';
  
  // Actualitzar visibilitat i progress
  updateVisibility();
  updProg();
}

function matchClient(c, q) {
  var fields = [
    c.nombre_completo, c.nif_cif, c.tel1, c.email1, c.email2,
    c.localidad, c.cp, c.tipo_label, c.motivo_contacto,
    c.par_nombre, c.par_ap1, c.par_ap2, c.par_nif,
    c.emp_razon, c.emp_cif, c.aut_nombre, c.aut_ap1
  ];
  return fields.some(function(f) { return f && (f+'').toLowerCase().indexOf(q) >= 0; });
}

function showSearchResults(results) {
  var el = document.getElementById('searchSuggestions');
  var cardTipo = document.getElementById('card-tipo-cliente');
  
  if (!el) return;
  
  if (!results.length) {
    el.innerHTML = '<div style="background:#fff;border:1px solid #ddd;border-radius:8px;padding:12px;text-align:center;color:#999;font-size:12px">No s\'han trobat resultats amb aquest criteri</div>';
    el.style.display = 'block';
    if (cardTipo) cardTipo.style.display = 'none'; // Col·lapsar card tipo
    return;
  }
  
  // Col·lapsar card tipo cliente quan hi ha resultats
  if (cardTipo) cardTipo.style.display = 'none';
  
  var html = '<div style="background:#fff;border:1px solid #ddd;border-radius:8px;overflow:hidden">';
  html += '<div style="padding:10px 14px;background:#f5f5f5;border-bottom:1px solid #ddd;font-size:12px;font-weight:600;color:#666">' + 
    results.length + ' client' + (results.length>1?'s':'') + ' trobat' + (results.length>1?'s':'') + '</div>';
  
  html += results.map(function(c) {
    var fn = c._filename || '';
    var fecha = fn.length >= 8 ? fn.slice(6,8)+'/'+fn.slice(4,6)+'/'+fn.slice(0,4) : '';
    var nif = c.nif_cif ? ' · '+c.nif_cif : '';
    var tel = c.tel1 ? ' · '+c.tel1 : '';
    var meta = (c.tipo_label||'') + nif + tel + (fecha?' · '+fecha:'');
    
    // EMOJIS segons origen
    var origen = '';
    var hasOD = c._source && (c._source === 'od' || c._source === 'both');
    var hasCRM = c._source && (c._source === 'crm' || c._source === 'both');
    
    if (hasOD && hasCRM) {
      origen = '<span style="margin-left:6px" title="OneDrive + CRM">📁💼</span>';
    } else if (hasOD) {
      origen = '<span style="margin-left:6px" title="OneDrive">📁</span>';
    } else if (hasCRM) {
      origen = '<span style="margin-left:6px" title="CRM">💼</span>';
    }
    
    var safeFn = fn.replace(/'/g,"&#39;").replace(/"/g,"&quot;");
    return '<div onclick="loadClientAndClose(\'' + safeFn + '\')" style="padding:12px 14px;border-bottom:1px solid #eee;cursor:pointer;transition:background .15s" onmouseover="this.style.background=\'#f8f8f8\'" onmouseout="this.style.background=\'#fff\'">' +
      '<div style="display:flex;align-items:center;justify-content:space-between">' +
      '<div style="flex:1"><div style="font-weight:600;font-size:14px;color:#333">' + (c.nombre_completo||'—') + origen + '</div>' +
      '<div style="font-size:12px;color:#666;margin-top:3px">' + meta + '</div></div>' +
      '<span style="font-size:11px;color:#999;background:#f0f0f0;padding:4px 10px;border-radius:5px;font-weight:600">' + (c.tipo_label||'—') + '</span></div></div>';
  }).join('');
  
  html += '</div>';
  
  el.innerHTML = html;
  el.style.display = 'block';
}

function hideSearchResults() {
  var el = document.getElementById('searchSuggestions');
  var cardTipo = document.getElementById('card-tipo-cliente');
  
  if (el) el.style.display = 'none';
  
  // Restaurar card tipo cliente
  if (cardTipo) cardTipo.style.display = 'block';
}

document.addEventListener('click', function(e) {
  if (!e.target.closest('.search-box')) hideSearchResults();
});

async function loadClientAndClose(filename) {
  hideSearchResults();
  var si = document.getElementById('searchInput');
  if(si) si.value = '';
  // Try to find from allClients cache first
  var client = allClients.find(function(cl) { return cl._filename === filename; });

  // Si el client no està al cache i ve d'OD index (filename comença amb 'od:'), no podem fer fallback al disc
  if (!client && filename && filename.indexOf('od:') === 0) {
    showToast('Client no disponible al cache.','error');
    return;
  }

  // Si tenim client al cache → enriquir si cal i respectar el toggle
  if (client) {
    // ── Enriquir amb dades completes si és client OD lleuger (només té nom/NIF/tel1/email1) ──
    if (client._odIndex && !client._enriched && client.nif_cif) {
      console.log('☁️ [ENRICH] Demanant dades completes per NIF ' + client.nif_cif);
      try {
        if (typeof showToast === 'function') showToast('Carregant dades del client...', 'info');
        var details = await _requestClientDetailsViaPostMessage(client.nif_cif, 5000);
        if (details && typeof details === 'object') {
          // ── DEBUG: mostrar TOTS els camps rebuts per facilitar mapping ──
          console.log('☁️ [ENRICH-DEBUG] Camps rebuts (' + Object.keys(details).length + '):', Object.keys(details));
          console.log('☁️ [ENRICH-DEBUG] Sample valors:', JSON.stringify(details).substring(0, 1500));
          // Merge: completes prevalen però conservem metadades de l'índex
          var enrichedClient = Object.assign({}, client, details);
          enrichedClient._enriched = true;
          enrichedClient._enrichedAt = Date.now();
          enrichedClient._filename = client._filename;
          enrichedClient._source = client._source;
          enrichedClient._refnumpers = client._refnumpers || details.refnumpers || null;
          var idx = allClients.findIndex(function(c) { return c._filename === filename; });
          if (idx >= 0) allClients[idx] = enrichedClient;
          client = enrichedClient;
          console.log('☁️ [ENRICH] Client enriquit amb ' + Object.keys(details).length + ' camps');
        } else {
          console.warn('☁️ [ENRICH] No s\'han rebut dades completes — usant dades lleugeres');
        }
      } catch(e) {
        console.warn('☁️ [ENRICH] Error obtenint detalls:', e && e.message);
      }
    }

    if (_showOdModal) {
      // Toggle ON → fitxa al card 0 + identificación/contacto collapsed
      showClientFitxaCard0(client);
      setFormMode('client-loaded-fitxa');
    } else {
      // Toggle OFF → carrega directe + identificación/contacto collapsed
      loadClientIntoForm(client);
      setFormMode('client-loaded-directe');
    }
    return;
  }

  // Fallback: read directly from file (només per clients de clientesDir)
  if (clientesDir) {
    try {
      var fh = await clientesDir.getFileHandle(filename);
      var file = await fh.getFile();
      var data = JSON.parse(await file.text());
      data._filename = filename;
      data._source = 'od';
      if (_showOdModal) {
        showClientFitxaCard0(data);
        setFormMode('client-loaded-fitxa');
      } else {
        loadClientIntoForm(data);
        setFormMode('client-loaded-directe');
      }
    } catch(e) {
      showToast('No se pudo cargar el cliente.','error');
    }
  }
}

// ─── LOAD CLIENT INTO FORM ────────────────────────────────────────
// Helper: troba el primer valor no buit d'una llista de noms de camp candidats.
// Permet acceptar JSONs amb diferents convencions (format formulari "par_nombre"
// o format extret del CRM "nombre", "name", etc.) sense haver de saber l'origen.
function _pickField(d) {
  for (var i = 1; i < arguments.length; i++) {
    var k = arguments[i];
    if (!k) continue;
    var v = d[k];
    if (v !== undefined && v !== null && v !== '') return v;
  }
  return '';
}

// ─── Helpers de parsing per format CRM Occident (cli_crm_<NIF>.json) ──
// Data DD.MM.YYYY o DD/MM/YYYY → YYYY-MM-DD (per input type="date")
function _parseCRMDate(s) {
  if (!s || typeof s !== 'string') return '';
  var m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
  if (m) {
    var dd = ('0' + m[1]).slice(-2);
    var mm = ('0' + m[2]).slice(-2);
    return m[3] + '-' + mm + '-' + dd;
  }
  return s; // possiblement ja en format ISO
}

// Adreça format CRM: "C/ Núria, 1, 2n, - 17600 - FIGUERES"
// Estratègia: localitzar CP (5 dígits) i separar la cadena
function _parseCRMAdreca(s) {
  if (!s || typeof s !== 'string') return { dir: '', cp: '', muni: '' };
  var cpMatch = s.match(/\b(\d{5})\b/);
  if (!cpMatch) return { dir: s.replace(/^[\s,-]+|[\s,-]+$/g, ''), cp: '', muni: '' };
  var cp = cpMatch[1];
  var idx = s.indexOf(cp);
  var before = s.substring(0, idx).replace(/[\s,-]+$/g, '').trim();
  var after = s.substring(idx + 5).replace(/^[\s,-]+/g, '').trim();
  return { dir: before, cp: cp, muni: after };
}

// Nom complet format ESPANYOL: "NOM AP1 AP2" (el que apareix als rebuts/sinistres)
// Variant amb coma: "AP1 AP2, NOM"
function _parseCRMNom(s) {
  if (!s || typeof s !== 'string') return { nombre: '', ap1: '', ap2: '' };
  s = s.trim().replace(/\s+/g, ' ');
  // Variant amb coma: "AP1 AP2, NOM" (format CRM intern)
  if (s.indexOf(',') >= 0) {
    var parts = s.split(',');
    var cognoms = parts[0].trim().split(/\s+/);
    return {
      nombre: parts.slice(1).join(',').trim(),
      ap1: cognoms[0] || '',
      ap2: cognoms.slice(1).join(' ') || ''
    };
  }
  // Sense coma: format espanyol natural "NOM AP1 AP2" o "NOM1 NOM2 AP1 AP2"
  var p = s.split(/\s+/);
  if (p.length === 1) return { nombre: p[0], ap1: '', ap2: '' };
  if (p.length === 2) return { nombre: p[0], ap1: p[1], ap2: '' };
  if (p.length === 3) return { nombre: p[0], ap1: p[1], ap2: p[2] };
  // 4+ paraules: heurística per nom compost a l'inici
  // Casos comuns: "Maria Carmen X Y", "Jose Antonio X Y", "Maria del Carmen X Y"
  var compounds = ['carmen','antonio','jose','jesus','luis','manuel','francisco',
                   'alfonso','angel','enrique','dolores','del','de','la','los'];
  var lower = p[1].toLowerCase();
  if (compounds.indexOf(lower) >= 0 && p.length >= 4) {
    // "Maria Carmen Faya Romero" → nom="Maria Carmen", ap1="Faya", ap2="Romero"
    // "Maria del Carmen Faya Romero" → nom="Maria del Carmen", ap1="Faya", ap2="Romero"
    var nomEnd = 2;
    while (nomEnd < p.length - 2 && compounds.indexOf(p[nomEnd].toLowerCase()) >= 0) nomEnd++;
    return {
      nombre: p.slice(0, nomEnd).join(' '),
      ap1: p[nomEnd] || '',
      ap2: p.slice(nomEnd + 1).join(' ')
    };
  }
  // Per defecte (4+ paraules sense partícula reconeguda): NOM=1a, AP1=2a, AP2=resta
  return { nombre: p[0], ap1: p[1], ap2: p.slice(2).join(' ') };
}

// Extreure nom del client des dels rebuts/sinistres/pòlisses
// Cerca patró "Asegurado: <NOM COMPLET> F. nacimiento:" o "Asegurada: ..."
function _extractNameFromReceipts(d) {
  var sources = [];
  if (Array.isArray(d.rebuts)) sources = sources.concat(d.rebuts);
  if (Array.isArray(d.sinistres)) sources = sources.concat(d.sinistres);
  if (Array.isArray(d.polisses)) sources = sources.concat(d.polisses);
  for (var i = 0; i < sources.length; i++) {
    var row = sources[i];
    if (!Array.isArray(row)) continue;
    for (var j = 0; j < row.length; j++) {
      var cell = row[j];
      if (typeof cell !== 'string') continue;
      // Patró estricte amb F. nacimiento
      var m = cell.match(/Asegurad[oa]:\s*(.+?)\s+F\.\s*nacimiento/i);
      if (m && m[1]) return m[1].trim();
      // Variant més laxa: "Asegurado: NOM" sense més
      var m2 = cell.match(/Asegurad[oa]:\s*([^,\n\r]+?)(?:\s*[-·]|\s*$)/i);
      if (m2 && m2[1]) {
        var v = m2[1].trim();
        if (v.length > 3 && v.length < 80) return v;
      }
    }
  }
  return '';
}

// Detecta si una cadena és nom d'empresa (conté S.L., S.A., etc.)
function _isCompanyName(s) {
  if (!s) return false;
  return /\b(S\.?L\.?U?\.?|S\.?A\.?U?\.?|S\.?C\.?|S\.?C\.?P\.?|C\.?B\.?|S\.?L\.?L\.?|S\.?COOP\.?)\b/i.test(s);
}

function loadClientIntoForm(d) {
  // Reset first
  clearForm();

  var sf = function(id, val) {
    if (val===undefined||val===null||val==='') return;
    var el = document.getElementById(id);
    if (el) { el.value = val; }
  };

  // ── Determinar tipus de client ──
  var t = d.tipo || d.tipoCliente || d.tipo_cliente;
  if (!t) {
    var nifCif = (d.nif_cif || d.nif || d.cif || '').toUpperCase();
    var hasCif = nifCif && /^[A-HJNPQRSUVW]/.test(nifCif);
    var nameSeemsCompany = _isCompanyName(d.nom || d.razon_social || d.emp_razon || '');
    if (hasCif || d.emp_cif || d.razon_social || d.razonSocial || nameSeemsCompany) t = 'emp';
    else if (d.aut_actividad || d.actividad_aut || d.es_autonomo) t = 'aut';
    else t = 'par';
  }
  var tipoEl = document.getElementById('tipo-' + t);
  if (tipoEl) setTipo(t, tipoEl);

  // Restore autEmpSi
  if ((d.aut_emp || d.aut_empresa) && t==='aut') {
    autEmpSi = true;
    var si=document.getElementById('aut-emp-si');var no=document.getElementById('aut-emp-no');
    if(si){si.classList.add('si');} if(no){no.classList.remove('no');}
    var w=document.getElementById('aut-emp-select-wrap');if(w)w.style.display='block';
    updateNegocioGroups();
  }
  if (d.vtos_per) {
    vtosPerSi = true;
    var si=document.getElementById('vper-si');var no=document.getElementById('vper-no');
    if(si)si.classList.add('si'); if(no)no.classList.remove('no');
    updateVisibility();
  }

  // ── PARSING ESPECIAL camps format CRM Occident català ──
  // El camp `nom` pot venir null si no s'extreu de la fitxa V360. En aquest cas,
  // fem fallback als rebuts/sinistres/pòlisses que contenen "Asegurado: NOM COMPLET F. nacimiento:"
  var nomFromCRM = _pickField(d, 'nom', 'nombre_crm');
  if (!nomFromCRM || nomFromCRM === null) {
    nomFromCRM = _extractNameFromReceipts(d);
    if (nomFromCRM) console.log('☁️ [ENRICH] Nom extret dels rebuts/sinistres:', nomFromCRM);
  }
  if (nomFromCRM && !d.par_nombre && !d.aut_nombre && !d.emp_razon && !d.razon_social) {
    if (t === 'emp') {
      sf('emp-razon', nomFromCRM); // Empresa: nom complet → razón social
    } else {
      var partsNom = _parseCRMNom(nomFromCRM);
      if (t === 'aut') {
        sf('aut-nombre', partsNom.nombre);
        sf('aut-ap1', partsNom.ap1);
        sf('aut-ap2', partsNom.ap2);
      } else {
        sf('par-nombre', partsNom.nombre);
        sf('par-ap1', partsNom.ap1);
        sf('par-ap2', partsNom.ap2);
      }
    }
  }

  // Adreça: parsejar si ve com a string únic (format CRM)
  var adrecaFromCRM = _pickField(d, 'adreca', 'direccion_completa', 'adressComplete');
  if (adrecaFromCRM && !d.direccion && !d.dir) {
    var addr = _parseCRMAdreca(adrecaFromCRM);
    sf('dir', addr.dir);
    sf('cp', addr.cp);
    sf('muni', addr.muni);
  }

  // Data naixement: convertir format si cal
  var fnacRaw = _pickField(d, 'fnaixement', 'fecha_nacimiento', 'fechaNacimiento', 'par_fnac', 'aut_fnac');
  if (fnacRaw) {
    var fnacISO = _parseCRMDate(fnacRaw);
    if (t === 'aut') sf('aut-fnac', fnacISO);
    else sf('par-fnac', fnacISO);
  }

  // ── Camps Particular (sinònims si encara no s'han omplert) ──
  if (!document.getElementById('par-nombre').value) sf('par-nombre', _pickField(d, 'par_nombre', 'par-nombre', 'nombre', 'name', 'first_name', 'firstName', 'NOMBRE'));
  if (!document.getElementById('par-ap1').value) sf('par-ap1', _pickField(d, 'par_ap1', 'par-ap1', 'apellido1', 'ap1', 'apellidoPaterno', 'AP1', 'apellido_1', 'primer_apellido'));
  if (!document.getElementById('par-ap2').value) sf('par-ap2', _pickField(d, 'par_ap2', 'par-ap2', 'apellido2', 'ap2', 'apellidoMaterno', 'AP2', 'apellido_2', 'segundo_apellido'));
  sf('par-nif',      _pickField(d, 'par_nif', 'par-nif', 'nif', 'dni', 'nie', 'nif_cif', 'documento', 'NIF'));
  sf('par-estcivil', _pickField(d, 'par_estcivil', 'par-estcivil', 'estado_civil', 'estadoCivil', 'estcivil', 'ESTADO_CIVIL'));
  sf('par-hijos',    _pickField(d, 'par_hijos', 'par-hijos', 'hijos', 'numero_hijos', 'num_hijos'));

  // ── Camps Empresa ──
  if (!document.getElementById('emp-razon').value) sf('emp-razon', _pickField(d, 'emp_razon', 'emp-razon', 'razon_social', 'razonSocial', 'denominacion', 'nombre_empresa', 'RAZON_SOCIAL', 'razon'));
  sf('emp-cif',       _pickField(d, 'emp_cif', 'emp-cif', 'cif', 'nif_cif', 'nif', 'NIF', 'CIF'));
  sf('emp-actividad', _pickField(d, 'emp_actividad', 'emp-actividad', 'actividad', 'professio', 'sector', 'cnae', 'actividad_economica'));
  sf('emp-antiguedad',_pickField(d, 'emp_antiguedad', 'emp-antiguedad', 'antiguedad', 'antiguitat', 'fecha_constitucion', 'fundacion'));
  sf('empleados-emp', _pickField(d, 'emp_empleados', 'empleados-emp', 'empleados', 'num_empleados', 'numero_empleados'));

  // ── Camps Autònom ──
  if (!document.getElementById('aut-nombre').value) sf('aut-nombre', _pickField(d, 'aut_nombre', 'aut-nombre', 'nombre', 'NOMBRE'));
  if (!document.getElementById('aut-ap1').value) sf('aut-ap1', _pickField(d, 'aut_ap1', 'aut-ap1', 'apellido1', 'ap1', 'primer_apellido', 'AP1'));
  if (!document.getElementById('aut-ap2').value) sf('aut-ap2', _pickField(d, 'aut_ap2', 'aut-ap2', 'apellido2', 'ap2', 'segundo_apellido', 'AP2'));
  sf('aut-nif',       _pickField(d, 'aut_nif', 'aut-nif', 'nif', 'dni', 'nif_cif', 'NIF'));
  sf('aut-actividad', _pickField(d, 'aut_actividad', 'aut-actividad', 'actividad', 'professio', 'sector', 'cnae'));
  sf('aut-antiguedad',_pickField(d, 'aut_antiguedad', 'aut-antiguedad', 'antiguedad', 'antiguitat', 'fecha_alta_autonomo'));
  sf('empleados-aut', _pickField(d, 'aut_empleados', 'empleados-aut', 'empleados', 'num_empleados'));

  // Sexo (visual)
  setSexoVal('par-sexo', _pickField(d, 'par_sexo', 'par-sexo', 'sexo', 'SEXO', 'genero'));
  setSexoVal('aut-sexo', _pickField(d, 'aut_sexo', 'aut-sexo', t==='aut' ? 'sexo' : '', t==='aut' ? 'SEXO' : '', t==='aut' ? 'genero' : ''));

  // ── Contacte ──
  sf('tel1',   _pickField(d, 'tel1', 'telefono1', 'telefono', 'telef', 'telefon', 'TELEFONO', 'TEL1', 'phone', 'phone1'));
  sf('wapp',   _pickField(d, 'whatsapp', 'wapp', 'tel_whatsapp', 'telefono_whatsapp'));
  sf('redes',  _pickField(d, 'redes', 'redes_sociales', 'social'));
  sf('email1', _pickField(d, 'email1', 'email', 'EMAIL', 'EMAIL1', 'correo', 'correo_electronico', 'mail'));
  sf('email2', _pickField(d, 'email2', 'EMAIL2', 'correo2', 'email_2', 'mail2'));
  if (!document.getElementById('dir').value) sf('dir', _pickField(d, 'direccion', 'dir', 'address', 'domicilio', 'DIRECCION', 'calle', 'via'));
  if (!document.getElementById('cp').value) sf('cp', _pickField(d, 'cp', 'codigo_postal', 'codigoPostal', 'CP', 'postal_code', 'zip'));
  if (!document.getElementById('muni').value) sf('muni', _pickField(d, 'localidad', 'muni', 'municipio', 'poblacion', 'LOCALIDAD', 'city', 'ciudad'));

  // ── Observacions ──
  sf('obs',            _pickField(d, 'motivo_contacto', 'obs', 'observaciones', 'notas', 'comentarios'));
  sf('primera-accion', _pickField(d, 'primera_accion', 'primera-accion', 'primeraAccion'));
  sf('fecha-accion',   _pickField(d, 'fecha_accion', 'fecha-accion', 'fechaAccion'));
  sf('origen-detalle', _pickField(d, 'origen_detalle', 'origen-detalle', 'origenDetalle'));

  // Origen sw (radio)
  if (d.origen_sw) {
    document.querySelectorAll('input[name="origen-sw"]').forEach(function(r) {
      r.checked = (r.value === d.origen_sw || d.origen_sw.indexOf(r.value) >= 0);
    });
  }

  // Colaboradores
  function restoreColab(radioName, racf, stateVar) {
    document.querySelectorAll('[name="'+radioName+'"]').forEach(function(r) {
      if (r.value === racf) {
        r.checked = true;
        var opt = r.closest('.colab-opt');
        if (opt) {
          opt.closest('.colab-row').querySelectorAll('.colab-opt').forEach(function(o){o.classList.remove('active');});
          opt.classList.add('active');
        }
      }
    });
  }
  if (d.colab_recoge_racf) { colabRecoge=d.colab_recoge_racf; restoreColab('colab-recoge', d.colab_recoge_racf); }
  if (d.colab_asigna_racf) { colabAsigna=d.colab_asigna_racf; restoreColab('colab-asigna', d.colab_asigna_racf); }

  // Ramo data + pill states
  try { ramoData = d.ramo_data ? JSON.parse(d.ramo_data) : {}; } catch(e) { ramoData={}; }
  setTimeout(function() {
    Object.keys(ramoData).forEach(function(key) {
      var st = ramoData[key] && ramoData[key].estado;
      if (st) {
        var pill = document.querySelector('[data-key="'+key+'"]');
        if (pill) applyPillState(pill, st, false);
      }
      refreshPillBadge(key);
    });
    updProg();
  }, 120);

  // Show load banner
  setFormDates(d.fechaCreacion||null, d.fechaModificacion||null);
  var banner = document.getElementById('loadBanner');
  if (banner) {
    document.getElementById('loadBannerText').textContent = '\uD83D\uDC64 Cargado: ' + (d.nombre_completo||d.nom||d.nombre||'cliente') + (d.timestamp ? '  \u00b7  ' + d.timestamp : '');
    banner.style.display = 'flex';
  }

  window.scrollTo({top:0, behavior:'smooth'});
}

function setSexoVal(hiddenId, val) {
  if (!val) return;
  var hid = document.getElementById(hiddenId);
  if (!hid) return;
  document.querySelectorAll('.sexo-selector').forEach(function(sel) {
    sel.querySelectorAll('.sexo-btn').forEach(function(b) {
      var oc = b.getAttribute('onclick')||'';
      if (oc.indexOf(hiddenId) >= 0 && oc.indexOf("'"+val+"'") >= 0) {
        sel.querySelectorAll('.sexo-btn').forEach(function(x){x.classList.remove('active');});
        b.classList.add('active');
        hid.value = val;
      }
    });
  });
}

// ─── INIT ────────────────────────────────────────────────────────
fillEmpleados('empleados-emp');
fillEmpleados('empleados-aut');
initGrids();
updateVisibility();
updateNegocioGroups();
updProg();
initDirFromIDB();
  initReportesDirFromIDB();

// ─── NIF / NIE VALIDATION ────────────────────────────────────
var _NIF_LETTERS = 'TRWAGMYFPDXBNJZSQVHLCKE';
function _calcNifLetter(n){ return _NIF_LETTERS[n%23]; }

function validateNIF(val) {
  val = (val||'').toUpperCase().replace(/\s/g,'');
  if(!val) return true;
  if(/^[0-9]{8}[A-Z]$/.test(val)) return val[8]===_calcNifLetter(parseInt(val.slice(0,8)));
  if(/^[XYZ][0-9]{7}[A-Z]$/.test(val)){
    var first={X:'0',Y:'1',Z:'2'}[val[0]];
    return val[8]===_calcNifLetter(parseInt(first+val.slice(1,8)));
  }
  return false;
}

function showToast(msg, type) {
  var t=document.getElementById('_globalToast');
  if(!t){t=document.createElement('div');t.id='_globalToast';document.body.appendChild(t);}
  t.className='toast-notif toast-'+(type||'error');
  t.textContent=msg;
  void t.offsetWidth;
  t.classList.add('show');
  clearTimeout(t._tid);
  t._tid=setTimeout(function(){t.classList.remove('show');},3800);
}

function validateNifField(input) {
  var val=(input.value||'').toUpperCase().replace(/\s/g,'');
  input.value=val;
  if(!val){input.style.borderColor='';return;}
  if(!validateNIF(val)){
    showToast('\u26A0\uFE0F NIF/NIE inv\u00e1lido: '+val+' \u2014 verifica el d\u00edgito de control','error');
    input.style.borderColor='#DC0028';
    setTimeout(function(){
      input.value='';
      input.style.borderColor='';
      checkNieNacionalidad(input);
      input.focus();
    },1800);
  } else {
    input.style.borderColor='#10B981';
    setTimeout(function(){input.style.borderColor='';},2200);
  }
}

function checkNieNacionalidad(input){
  var v=(input.value||'').trim().toUpperCase();
  var isNIE=/^[XYZ]/i.test(v);
  var prefix = (input.id||'').split('-')[0] || 'par';
  var cadWrap=document.getElementById(prefix+'-nie-caducidad-wrap');
  var nacWrap=document.getElementById(prefix+'-nacionalidad-wrap');
  if(cadWrap) cadWrap.style.display=isNIE?'':'none';
  if(nacWrap) nacWrap.style.display=isNIE?'flex':'none';
}

// ─── CARNET AGE VALIDATION ─────────────────────────────────
var _CARNET_AGES = {'AM/LCM':14,'A1':16,'A2':18,'A':18,'B':18,'C':18};

function checkCarnetAge(prefix, num) {
  var fnacEl  = document.getElementById(prefix+'-fnac');
  var tipoEl  = document.getElementById(prefix+'-carnet'+num+'-tipo');
  var fechaEl = document.getElementById(prefix+'-carnet'+num+'-fecha');
  if(!fnacEl||!tipoEl||!fechaEl) return;
  if(!fnacEl.value||!tipoEl.value||!fechaEl.value){
    if(tipoEl)tipoEl.style.borderColor=''; if(fechaEl)fechaEl.style.borderColor=''; return;
  }
  var fnacD=new Date(fnacEl.value), carnetD=new Date(fechaEl.value);
  var minAge=_CARNET_AGES[tipoEl.value]; if(!minAge) return;
  var age=carnetD.getFullYear()-fnacD.getFullYear();
  var m=carnetD.getMonth()-fnacD.getMonth();
  if(m<0||(m===0&&carnetD.getDate()<fnacD.getDate())) age--;
  if(age<minAge){
    tipoEl.style.borderColor='#DC0028'; fechaEl.style.borderColor='#DC0028';
    showToast('\u26A0\uFE0F Carnet '+tipoEl.value+': edad m\u00ednima '+minAge+' a\u00f1os (tiene '+age+')','warn');
  } else {
    tipoEl.style.borderColor='#10B981'; fechaEl.style.borderColor='#10B981';
    setTimeout(function(){tipoEl.style.borderColor='';fechaEl.style.borderColor='';},2200);
  }
}


// ─── REPORT GENERATOR ────────────────────────────────────
function fmtDate(iso){
  if(!iso) return '—';
  // try parse
  var d=new Date(iso); if(isNaN(d)) return iso;
  return d.toLocaleDateString('es-ES',{day:'2-digit',month:'2-digit',year:'numeric'});
}
function fmtMoney(v){
  if(!v&&v!==0) return '—';
  return parseFloat(v).toLocaleString('es-ES',{minimumFractionDigits:0,maximumFractionDigits:2})+' €';
}
function rptPillDetails(type, key) {
  var d = ramoData[key] || {};
  var lines = [];
  var typeBase = type.replace(/-\d+$/,'');
  if(typeBase==='hogar'){
    if(d.tipoProp) lines.push(['Tipo propiedad', d.tipoProp]);
    if(d.personas) lines.push(['Personas aseguradas', d.personas]);
    if(d.m2) lines.push(['M\u00b2', d.m2]);
    if(d.manobraContinente) lines.push(['Continente', fmtMoney(d.manobraContinente)]);
    if(d.manobraContenido) lines.push(['Contenido', fmtMoney(d.manobraContenido)]);
    if(d.capital) lines.push(['Capital', fmtMoney(d.capital)]);
    if(d.hipoteca) lines.push(['Hipoteca', d.hipoteca]);
    if(d.dirInmueble) lines.push(['Inmueble', [d.dirInmueble,d.cpInmueble,d.locInmueble].filter(Boolean).join(', ')]);
    if(d.primaActual) lines.push(['Prima actual', fmtMoney(d.primaActual)]);
    if(d.vto) lines.push(['Vencimiento', fmtDate(d.vto)]);
    if(d.cia) lines.push(['Compa\u00f1\u00eda actual', d.cia]);
  } else if(typeBase==='auto'||typeBase==='moto'){
    if(d.matricula) lines.push(['Matr\u00edcula', d.matricula]);
    if(d.marca||d.modelo||d.version) lines.push(['Veh\u00edculo', [d.marca,d.modelo,d.version].filter(Boolean).join(' ')]);
    if(d.mesVcto) lines.push(['Mes vencimiento', d.mesVcto]);
    if(d.propietario) lines.push(['Propietario', d.propietario==='otro'?
      [d.ownerNombre,d.ownerAp1,d.ownerAp2].filter(Boolean).join(' ')||'Otro (sin nombre)':
      d.propietario.charAt(0).toUpperCase()+d.propietario.slice(1)]);
    if(d.propietario==='otro' && d.ownerDni) lines.push(['DNI propietario', d.ownerDni]);
    if(d.conductor) lines.push(['Conductor habitual', d.conductor==='otro'?
      [d.driverNombre,d.driverAp1,d.driverAp2].filter(Boolean).join(' ')||'Otro (sin nombre)':
      d.conductor.charAt(0).toUpperCase()+d.conductor.slice(1)]);
    if(d.conductor==='otro' && d.driverDni) lines.push(['DNI conductor', d.driverDni]);
    if(d.sincos && d.sincos.some(function(s){return s.cia;})){
      var sincoStr = d.sincos.filter(function(s){return s.cia;}).map(function(s){
        return s.cia+(s.poliza?' / P\u00f3l: '+s.poliza:'')+(s.matricula?' / Mat: '+s.matricula:'');
      }).join(' | ');
      lines.push(['SINCO / Competencia', sincoStr]);
    }
    if(d.primaActual) lines.push(['Prima actual', fmtMoney(d.primaActual)]);
    if(d.vto) lines.push(['Vencimiento', fmtDate(d.vto)]);
    if(d.cia) lines.push(['Compa\u00f1\u00eda actual', d.cia]);
  } else if(typeBase==='salud'){
    if(d.modalidad) lines.push(['Modalidad', d.modalidad]);
    if(d.copago) lines.push(['Copago', d.copago]);
    if(d.personas) lines.push(['Personas', d.personas]);
    if(d.personasData && d.personasData.length){
      d.personasData.forEach(function(p,i){
        var parts=[];
        if(p.fnac)parts.push('Nac:'+fmtDate(p.fnac));
        if(p.sexo)parts.push('Sexo:'+p.sexo);
        if(parts.length)lines.push(['Asegurado '+(i+1), parts.join(' | ')]);
        if(p.prof)lines.push(['Profesi\u00f3n '+(i+1), p.prof]);
        if(p.deporte==='Si'&&p.deporteNombre)lines.push(['Deporte '+(i+1), p.deporteNombre]);
      });
    }
    if(d.observaciones)lines.push(['Observaciones', d.observaciones]);
    if(d.observacionesNota)lines.push(['Notas', d.observacionesNota]);
    if(d.primaActual) lines.push(['Prima actual', fmtMoney(d.primaActual)]);
    if(d.vto) lines.push(['Vencimiento', fmtDate(d.vto)]);
    if(d.cia) lines.push(['Compa\u00f1\u00eda actual', d.cia]);
  } else if(typeBase==='decesos'){
    if(d.personas) lines.push(['Personas', d.personas]);
    if(d.personasData && d.personasData.length){
      d.personasData.forEach(function(p,i){
        var parts=[];
        if(p.fnac)parts.push('Nac:'+fmtDate(p.fnac));
        if(p.sexo)parts.push('Sexo:'+p.sexo);
        if(parts.length)lines.push(['Asegurado '+(i+1), parts.join(' | ')]);
      });
    }
    if(d.mascotas && d.mascotas.length){
      d.mascotas.forEach(function(m,i){
        lines.push(['Mascota '+(i+1), m.tipo+(m.ppi?' | PPI:'+m.ppi:'')+(m.fnac?' | Nac:'+fmtDate(m.fnac):'')]);
      });
    }
    if(d.primaActual) lines.push(['Prima actual', fmtMoney(d.primaActual)]);
    if(d.vto) lines.push(['Vencimiento', fmtDate(d.vto)]);
    if(d.cia) lines.push(['Compa\u00f1\u00eda actual', d.cia]);
  } else if(typeBase==='vida'){
    if(d.personas) lines.push(['Personas', d.personas]);
    if(d.capital) lines.push(['Capital asegurado / a contratar', fmtMoney(d.capital)]);
    if(d.profesion) lines.push(['Profesi\u00f3n', d.profesion]);
    if(d.deportes) lines.push(['Pr\u00e1ctica deportes', d.deportes]);
    if(d.deporteNombre) lines.push(['Deporte', d.deporteNombre]);
    var destinos=[];
    if(d.destinoHipoteca)destinos.push('Hipoteca');
    if(d.destinoSucesiones)destinos.push('Sucesiones');
    if(d.destinoFamiliar)destinos.push('Protecci\u00f3n familiar');
    if(destinos.length) lines.push(['Destino', destinos.join(', ')]);
    if(d.personasData && d.personasData.length){
      d.personasData.forEach(function(p,i){
        var parts=[];
        if(p.fnac)parts.push('Nac:'+fmtDate(p.fnac));
        if(p.sexo)parts.push('Sexo:'+p.sexo);
        if(parts.length)lines.push(['Asegurado '+(i+1), parts.join(' | ')]);
      });
    }
    if(d.primaActual) lines.push(['Prima actual', fmtMoney(d.primaActual)]);
    if(d.vto) lines.push(['Vencimiento', fmtDate(d.vto)]);
    if(d.cia) lines.push(['Compa\u00f1\u00eda actual', d.cia]);
  } else if(typeBase==='ahorro-g'||typeBase==='ahorro-i'){
    if(d.tieneAhorros) lines.push(['Tiene ahorros actuales', d.tieneAhorros]);
    if(d.tieneAhorros==='Si'){
      if(d.aportaciones) lines.push(['Aportaci\u00f3n mensual', fmtMoney(d.aportaciones)]);
      if(d.activo) lines.push(['Activo acumulado', fmtMoney(d.activo)]);
      if(d.notas) lines.push(['Entidad actual', d.notas]);
    }
    if(d.quiereAhorrar) lines.push(['Quiere iniciar ahorro', d.quiereAhorrar]);
    if(d.quiereAhorrar==='Si'){
      if(d.aportacionesQuiere) lines.push(['Aportaci\u00f3n objetivo', fmtMoney(d.aportacionesQuiere)]);
      if(d.aportacionInicial) lines.push(['Aportaci\u00f3n inicial', fmtMoney(d.aportacionInicial)]);
      if(d.notasQuiere) lines.push(['Notas ahorro objetivo', d.notasQuiere]);
    }
    if(d.vto) lines.push(['Vencimiento', fmtDate(d.vto)]);
  } else if(typeBase==='comunidades'){
    if(d.nombreCom) lines.push(['Comunidad', d.nombreCom]);
    if(d.admin) lines.push(['Administrador', d.admin]);
    if(d.primaActual) lines.push(['Prima actual', fmtMoney(d.primaActual)]);
    if(d.vto) lines.push(['Vencimiento', fmtDate(d.vto)]);
    if(d.cia) lines.push(['Compa\u00f1\u00eda actual', d.cia]);
  } else if(typeBase==='embarcaciones'){
    if(d.bandera) lines.push(['Bandera / Zona', [d.bandera, d.zona].filter(Boolean).join(' | ')]);
    if(d.eslora) lines.push(['Eslora', d.eslora+'m']);
    if(d.nmotores) lines.push(['Motores', d.nmotores+(d.potencia?' × '+d.potencia+'CV':'')]);
    if(d.valorCasco) lines.push(['Valor casco', fmtMoney(d.valorCasco)]);
    if(d.valorAcc) lines.push(['Valor accesorios', fmtMoney(d.valorAcc)]);
    if(d.materialCasco) lines.push(['Material casco', d.materialCasco]);
    if(d.materialConst) lines.push(['Material construcci\u00f3n', d.materialConst]);
    if(d.anioConst) lines.push(['A\u00f1o construcci\u00f3n', d.anioConst]);
    if(d.primaActual) lines.push(['Prima actual', fmtMoney(d.primaActual)]);
    if(d.vto) lines.push(['Vencimiento', fmtDate(d.vto)]);
    if(d.cia) lines.push(['Compa\u00f1\u00eda actual', d.cia]);
  } else if(typeBase==='otro'){
    if(d.nombre) lines.push(['Ramo personalizado', d.nombre]);
    if(d.notas) lines.push(['Notas / Detalle', d.notas]);
    if(d.vto) lines.push(['Vencimiento', fmtDate(d.vto)]);
    if(d.cia) lines.push(['Compa\u00f1\u00eda actual', d.cia]);
  }
  // Common fields for all types
  if(d.observaciones) lines.push(['Observaciones', d.observaciones]);
  if(d.observacionesNota) lines.push(['Detalle observaciones', d.observacionesNota]);
  return lines;
}

// Render a pill as a card for the 3-column grid layout
function rptPillCard(type, key, accentColor, borderColor) {
  var cfg = RAMO_CONFIG[type.replace(/-\d+$/,'')] || {name:type, emoji:''};
  var rd = ramoData[key] || {};
  var details = rptPillDetails(type, key);
  var pillNameStr = rd.nombre && type.replace(/-\d+$/,'') === 'otro' ? rd.nombre : cfg.name;
  // Label-on-top, value-below layout (2-column grid)
  var detailRows = details.map(function(p){
    return '<div style="padding:5px 0;border-bottom:1px solid #f0f0f0">'+
      '<div style="color:#999;font-size:9px;font-family:Arial,sans-serif;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px">'+p[0]+'</div>'+
      '<div style="color:#202020;font-size:12px;font-family:Arial,sans-serif;font-weight:600;word-break:break-word">'+p[1]+'</div>'+
    '</div>';
  }).join('');
  var cardId = 'rptcard-'+key.replace(/[^a-zA-Z0-9]/g,'-');
  return '<div style="background:#fff;border:2px solid '+(borderColor||accentColor)+';border-radius:10px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,.08);display:flex;flex-direction:column">'+
    '<div onclick="var b=document.getElementById(\''+cardId+'\');var open=b.style.display!==\'none\';b.style.display=open?\'none\':\'block\';this.querySelector(\'.rpc-arrow\').textContent=open?\'\u25bc\':\'\u25b2\'" style="display:flex;align-items:center;gap:8px;padding:10px 14px;background:'+accentColor+';cursor:pointer;user-select:none">'+
      '<span style="font-size:16px;line-height:1">'+cfg.emoji+'</span>'+
      '<span style="font-family:Arial,sans-serif;font-size:12px;font-weight:700;color:#fff;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+pillNameStr+'</span>'+
      '<span class="rpc-arrow" style="font-size:10px;color:rgba(255,255,255,.8);flex-shrink:0">\u25b2</span>'+
    '</div>'+
    '<div id="'+cardId+'" data-rpcbody="1" style="padding:0 12px;display:grid;grid-template-columns:1fr 1fr;column-gap:12px">'+
      (detailRows || '<div style="padding:10px 0;color:#bbb;font-size:11px;font-style:italic;font-family:Arial,sans-serif;grid-column:1/-1">Sin datos adicionales</div>')+
    '</div>'+
  '</div>';
}

function rptPillCards(pills, accentColor, borderColor) {
  if(!pills || !pills.length) return '';
  var cards = pills.map(function(p){
    return '<div>'+rptPillCard(p.type,p.key,accentColor,borderColor)+'</div>';
  }).join('');
  return '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;padding:8px 0">'+cards+'</div>';
}

function rptPillRow(type, key, badgeBg, badgeColor, badgeBorder) {
  var cfg = RAMO_CONFIG[type.replace(/-\d+$/,'')] || {name:type, emoji:''};
  var rd = ramoData[key] || {};
  var details = rptPillDetails(type, key);
  var pillNameStr = rd.nombre && type.replace(/-\d+$/,'') === 'otro' ? rd.nombre : cfg.name;
  var detailHtml = '';
  if(details.length){
    detailHtml = details.map(function(p){
      return '<span style="display:inline-block;margin-right:14px"><strong style="color:#888;font-size:10px;text-transform:uppercase;letter-spacing:.4px;display:block">'+p[0]+'</strong><span style="color:#202020;font-size:12px">'+p[1]+'</span></span>';
    }).join('');
    detailHtml = '<div style="display:flex;flex-wrap:wrap;gap:4px 0;margin-top:6px">'+detailHtml+'</div>';
  }
  var badge = '<span class="rpt-pill-badge" style="background:'+badgeBg+';color:'+badgeColor+';border:1.5px solid '+badgeBorder+'">'+cfg.emoji+' '+pillNameStr+'</span>';
  return '<div class="rpt-pill-row">'+badge+'<div class="rpt-pill-notes">'+detailHtml+'</div></div>';
}

function rptSection(num, title, colorBar, content, emptyMsg) {
  var body = content || ('<p class="rpt-empty">'+(emptyMsg||'Sin datos')+'</p>');
  var secId = 'rptsec-'+num;
  var hasPillCards = content && content.indexOf('data-rpcbody=') !== -1;
  var toggleBtn = hasPillCards ? (
    '<button class="rpt-toggle-btn" data-open="1" onclick="'+
    '(function(b){'+
      'var s=b.closest(\'.rpt-section\');'+
      'var cs=s.querySelectorAll(\'[data-rpcbody]\');'+
      'var open=b.getAttribute(\'data-open\')==\'1\';'+
      'cs.forEach(function(c){c.style.display=open?\'none\':\'block\';});'+
      'b.setAttribute(\'data-open\',open?\'0\':\'1\');'+
      'b.innerHTML=open?\'&#9660; Mostrar\':\'&#9650; Ocultar\';'+
    '})(this)'+
    '">&#9650; Ocultar</button>'
  ) : '';
  return '<div class="rpt-section" id="'+secId+'">'+
    '<div class="rpt-section-hdr">'+
      '<span class="rpt-section-hdr-label" style="background:'+colorBar+'">'+title+'</span>'+
      toggleBtn+
    '</div>'+body+'</div>';
}

// Map a pill key to its category groups for report classification
function pillCategory(key) {
  var type = key.replace(/-\d+$/,'').replace(/^grid-[a-z]+-/,'');
  var estrateg = ['hogar','auto','moto','vida','salud','decesos','ahorro-g','ahorro-i'];
  var ofertable = ['embarcaciones','comunidades','movilidad','mascotas','bienes-consumo','plan-pensiones','accidentes'];
  var negAct = ['comercio-pyme','rc','acc-convenio','transportes'];
  var negEmp = ['acc-convenio','salud-col','ahorro-col'];
  var negAut = ['baja-ilt-aut','salud','subsidio'];
  // detect grid from key prefix
  if(key.indexOf('grid-estrategicos')===0||estrateg.indexOf(type)!==-1) return 'Particulares Estrat\u00e9gicos';
  if(key.indexOf('grid-ofertables')===0||ofertable.indexOf(type)!==-1) return 'Particulares Ofertables';
  if(key.indexOf('grid-neg-actividad')===0) return 'Negocio \u00b7 Actividad';
  if(key.indexOf('grid-neg-empleados')===0) return 'Negocio \u00b7 Empleados';
  if(key.indexOf('grid-neg-autonomo')===0) return 'Negocio \u00b7 Aut\u00f3nomo/Empresario';
  return 'Otros';
}

// Build pill chips grouped by category
function rptPillChipsGrouped(pills) {
  if(!pills||!pills.length) return '';
  var groups = {};
  var order = [];
  pills.forEach(function(p){
    var cat = pillCategory(p.key);
    if(!groups[cat]){groups[cat]=[];order.push(cat);}
    groups[cat].push(p);
  });
  return order.map(function(cat){
    var chips = groups[cat].map(function(p){
      var cfg = RAMO_CONFIG[p.type.replace(/-\d+$/,'')] || {name:p.type, emoji:''};
      var rd2 = ramoData[p.key] || {};
      var nm = rd2.nombre && p.type.replace(/-\d+$/,'') === 'otro' ? rd2.nombre : cfg.name;
      return '<span style="display:inline-block;background:#f0f0f0;border:1.5px dashed #bbb;border-radius:16px;padding:5px 14px;font-size:12px;font-family:Arial,sans-serif;color:#888;margin:3px 3px 3px 0">'+cfg.emoji+' '+nm+'</span>';
    }).join('');
    return '<div style="margin-bottom:8px"><div style="font-family:Arial,sans-serif;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:#bbb;margin-bottom:4px">'+cat+'</div>'+chips+'</div>';
  }).join('');
}

// Build un-offered pill chips grouped by category
function rptPillChipsUnoffered(pills) {
  if(!pills||!pills.length) return '';
  var groups = {};
  var order = [];
  pills.forEach(function(p){
    // Never show 'otro' in estratégicos no ofertados
    var typeBase = p.type.replace(/-\d+$/,'');
    if(typeBase==='otro' && p.gridId==='grid-estrategicos') return;
    var cat = pillCategory(p.key);
    if(!groups[cat]){groups[cat]=[];order.push(cat);}
    groups[cat].push(p);
  });
  return order.map(function(cat){
    var chips = groups[cat].map(function(p){
      var cfg = RAMO_CONFIG[p.type.replace(/-\d+$/,'')] || {name:p.type, emoji:''};
      var isEstrategico = (p.gridId === 'grid-estrategicos') || p.key.indexOf('grid-estrategicos') === 0;
      var chipStyle = isEstrategico
        ? 'display:inline-block;background:#FFF1F2;border:2px solid #DC0028;border-radius:16px;padding:4px 12px;font-size:11px;font-family:Arial,sans-serif;color:#DC0028;font-weight:600;margin:3px 3px 3px 0'
        : 'display:inline-block;background:#f8f8f8;border:1px solid #e8e8e8;border-radius:16px;padding:4px 12px;font-size:11px;font-family:Arial,sans-serif;color:#aaa;margin:3px 3px 3px 0';
      return '<span style="'+chipStyle+'">'+cfg.emoji+' '+cfg.name+'</span>';
    }).join('');
    return '<div style="margin-bottom:8px"><div style="font-family:Arial,sans-serif;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:#ccc;margin-bottom:4px">'+cat+'</div>'+chips+'</div>';
  }).join('');
}

function generateReport() {
  var d = collectData();
  var rd = ramoData;

  // Collect pills by state — filtering by client type, deduplicating by key
  var pills_vto=[], pills_facilita=[], pills_no_cambia_vto=[], pills_necesita=[], pills_no_cambia=[], pills_no_necesita=[], pills_unset=[], pills_en_vigor=[];
  var seenKeys = {};
  document.querySelectorAll('.ramo-pill[data-key]').forEach(function(pill){
    var key = pill.getAttribute('data-key');
    var type = pill.getAttribute('data-type');
    // Deduplicate: if same key already processed (e.g. salud-1 in two grids), skip
    if(seenKeys[key]) return;
    seenKeys[key] = true;
    // Filter by client type using parent grid ID (more reliable than key prefix)
    var gridId = pill.parentElement ? pill.parentElement.id : '';
    if(tipo==='par' && gridId.indexOf('grid-neg')===0) return;
    // For EMP, skip par/ofertable grids if vtosPerSi not set
    if(tipo==='emp' && !vtosPerSi && (gridId==='grid-estrategicos' || gridId==='grid-ofertables')) return;
    var estado = (rd[key]||{}).estado || '';
    var info = {type:type, key:key, gridId:gridId};
    if(estado==='vto-inmediato') pills_vto.push(info);
    else if(estado==='facilita') pills_facilita.push(info);
    else if(estado==='no-cambia-vto') pills_no_cambia_vto.push(info);
    else if(estado==='necesita') pills_necesita.push(info);
    else if(estado==='no-cambia') pills_no_cambia.push(info);
    else if(estado==='no-necesita') pills_no_necesita.push(info);
    else if(estado==='en-vigor') pills_en_vigor.push(info);
    else pills_unset.push(info);
  });

  // Format date of birth / age
  var fnac = d.fecha_nacimiento;
  var edad = '';
  if(fnac){ var bd=new Date(fnac); if(!isNaN(bd)){ edad=' \u00b7 '+(new Date().getFullYear()-bd.getFullYear())+' a\u00f1os'; } }

  // HEADER — logo (full) izquierda · fecha derecha
  var html = '<div class="rpt-header">';
  html += '<img src="imagenes/Bordado-OPTIMIZARTE copia.jpg" style="height:68px;width:auto;display:block;object-fit:contain" alt="OPTIMIZARTE">';
  html += '<div style="text-align:right"><div style="color:rgba(255,255,255,.7);font-size:11px;font-family:Arial,sans-serif">Informe generado</div><div style="color:#fff;font-size:13px;font-family:Arial,sans-serif;font-weight:700">'+new Date().toLocaleDateString('es-ES',{day:'2-digit',month:'long',year:'numeric'})+'</div><div style="color:rgba(255,255,255,.85);font-size:12px;font-family:Arial,sans-serif;font-weight:600;margin-top:4px">'+(d.nombre_completo||'')+'</div></div>';
  html += '</div>';

  // ── P0: PRODUCTOS CONTRATADOS / EN VIGOR ──────────────────────
  if(pills_en_vigor.length > 0){
    var s0 = rptPillCards(pills_en_vigor, '#16A34A', '#16A34A');
    html += rptSection(0,'PRODUCTOS CONTRATADOS / EN VIGOR','#16A34A', s0||'', '');
  }

  // ── P1: TARIFICACION INMEDIATA ──────────────────────────────────
  var s1 = rptPillCards(pills_vto, '#111', '#111');
  html += rptSection(1,'TARIFICACI\u00d3N INMEDIATA','#111', s1||'', 'Sin p\u00f3lizas con vencimiento inmediato.');

  // ── P2: CREACION DE NUEVAS OPORTUNIDADES ───────────────────────
  var s2 = rptPillCards(pills_facilita.concat(pills_no_cambia_vto), '#0563C1', '#2563EB');
  html += rptSection(2,'CREACI\u00d3N DE NUEVAS OPORTUNIDADES','#0563C1', s2||'', 'Sin p\u00f3lizas en competencia con VTO informado.');

  // ── P3: CANDI — Necesita pero no tiene (Creacion de Etiquetas CANDI) ──
  var s3cards = rptPillCards(pills_necesita, '#E57C00', '#F59E0B');
  // CANDI label line
  var candiNecLine = '';
  if(pills_necesita.length > 0){
    var candiNecNames = pills_necesita.map(function(p){
      var cfg = RAMO_CONFIG[p.type.replace(/-\d+$/,'')] || {name:p.type};
      var rd2 = ramoData[p.key]||{};
      var nm = (rd2.nombre && p.type.replace(/-\d+$/,'') === 'otro' ? rd2.nombre : cfg.name);
      return 'CANDI'+new Date().getFullYear()+(nm.toUpperCase().replace(/\s+/g,''));
    });
    var candiNecText = candiNecNames.join(', ');
    candiNecLine = '<div style="margin:12px 0 4px;background:#FEF3C7;border:2px solid #F59E0B;border-radius:8px;padding:10px 14px">'+
      '<div style="font-family:Arial,sans-serif;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#92400E;margin-bottom:6px">\ud83c\udff7\ufe0f Copiar directamente en GESTIONA</div>'+
      '<div style="font-family:Arial,sans-serif;font-size:13px;font-weight:700;color:#78350F;word-break:break-word">'+candiNecText+'</div>'+
    '</div>';
  }
  html += rptSection(3,'Creaci\u00f3n de Etiquetas CANDI \u2014 Necesita pero no tiene','#F59E0B', (s3cards+candiNecLine)||'', 'Sin productos identificados como necesarios.');

  // ── P4: CANDI — No quiere / Puede Cambiar ──────────────────────
  var s4cards = rptPillCards(pills_no_cambia, '#DC0028', '#DC0028');
  var candiNcLine = '';
  if(pills_no_cambia.length > 0){
    var candiNcNames = pills_no_cambia.map(function(p){
      var cfg = RAMO_CONFIG[p.type.replace(/-\d+$/,'')] || {name:p.type};
      var rd2 = ramoData[p.key]||{};
      var nm = (rd2.nombre && p.type.replace(/-\d+$/,'') === 'otro' ? rd2.nombre : cfg.name);
      return 'CANDI'+new Date().getFullYear()+(nm.toUpperCase().replace(/\s+/g,''));
    });
    var candiNcText = candiNcNames.join(', ');
    candiNcLine = '<div style="margin:12px 0 4px;background:#FEE2E2;border:2px solid #DC0028;border-radius:8px;padding:10px 14px">'+
      '<div style="font-family:Arial,sans-serif;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#991B1B;margin-bottom:6px">\ud83c\udff7\ufe0f Copiar directamente en GESTIONA</div>'+
      '<div style="font-family:Arial,sans-serif;font-size:13px;font-weight:700;color:#7F1D1D;word-break:break-word">'+candiNcText+'</div>'+
    '</div>';
  }
  html += rptSection(4,'Candi \u2014 No quiere / Puede Cambiar','#DC0028', (s4cards+candiNcLine)||'', 'Sin p\u00f3lizas en esta situaci\u00f3n.');

  // ── P5: PRODUCTOS NO OFERTADOS (before No Necesita per request) ──
  var s5unset = rptPillChipsUnoffered(pills_unset);
  html += rptSection(5,'Productos No Ofertados','#9CA3AF', s5unset?('<div style="padding:6px 0">'+s5unset+'</div>'):'', 'Todos los productos han sido valorados.');

  // ── P6: PRODUCTOS NO NECESITA ──────────────────────────────────
  var s6noNec = rptPillChipsGrouped(pills_no_necesita);
  html += rptSection(6,'Productos no necesita','#9CA3AF', s6noNec?('<div style="padding:6px 0">'+s6noNec+'</div>'):'', 'Ning\u00fan producto marcado como no necesita.');

  // ── P7: DATOS PERSONALES ───────────────────────────────────────
  var sexoMap = {'H':'Hombre','M':'Mujer'};
  var datosP = '';
  if(d.tipo==='par'||d.tipo==='aut'){
    var c1fecha = (document.getElementById('par-carnet1-fecha')||document.getElementById('aut-carnet1-fecha')||{value:''}).value;
    var c1tipo  = (document.getElementById('par-carnet1-tipo')||document.getElementById('aut-carnet1-tipo')||{value:''}).value;
    var c2fecha = (document.getElementById('par-carnet2-fecha')||document.getElementById('aut-carnet2-fecha')||{value:''}).value;
    var c2tipo  = (document.getElementById('par-carnet2-tipo')||document.getElementById('aut-carnet2-tipo')||{value:''}).value;
    datosP = '<div class="rpt-data-grid">';
    datosP += '<div class="rpt-data-item"><strong>Nombre completo</strong>'+d.nombre_completo+'</div>';
    datosP += '<div class="rpt-data-item"><strong>DNI / NIE</strong>'+(d.nif_cif||'\u2014')+'</div>';
    datosP += '<div class="rpt-data-item"><strong>Fecha nacimiento</strong>'+fmtDate(d.fecha_nacimiento)+'</div>';
    datosP += '<div class="rpt-data-item"><strong>Sexo</strong>'+(sexoMap[d.sexo]||d.sexo||'\u2014')+'</div>';
    datosP += '<div class="rpt-data-item"><strong>Estado civil</strong>'+(d.par_estcivil||'\u2014')+'</div>';
    datosP += '<div class="rpt-data-item"><strong>Hijos</strong>'+(d.par_hijos||'\u2014')+'</div>';
    if(c1tipo) datosP += '<div class="rpt-data-item"><strong>Carnet 1 ('+c1tipo+')</strong>'+fmtDate(c1fecha)+'</div>';
    if(c2tipo) datosP += '<div class="rpt-data-item"><strong>Carnet 2 ('+c2tipo+')</strong>'+fmtDate(c2fecha)+'</div>';
    datosP += '<div class="rpt-data-item"><strong>Tel\u00e9fono</strong>'+(d.tel1||'\u2014')+(d.whatsapp?' \u00b7 WA: '+d.whatsapp:'')+'</div>';
    datosP += '<div class="rpt-data-item"><strong>Email</strong>'+(d.email1||'\u2014')+'</div>';
    datosP += '<div class="rpt-data-item" style="grid-column:1/-1"><strong>Direcci\u00f3n</strong>'+(d.direccion?d.direccion+', '+d.cp+' '+d.localidad:'\u2014')+'</div>';
    datosP += '</div>';
  } else if(d.tipo==='emp'){
    var repNombre = (document.getElementById('emp-rep-nombre')||{value:''}).value;
    var repNif = (document.getElementById('emp-rep-nif')||{value:''}).value;
    datosP = '<div class="rpt-data-grid">';
    datosP += '<div class="rpt-data-item"><strong>Raz\u00f3n social</strong>'+d.nombre_completo+'</div>';
    datosP += '<div class="rpt-data-item"><strong>CIF</strong>'+(d.nif_cif||'\u2014')+'</div>';
    datosP += '<div class="rpt-data-item"><strong>Actividad</strong>'+(d.actividad||'\u2014')+'</div>';
    datosP += '<div class="rpt-data-item"><strong>Antig\u00fcedad</strong>'+(d.antiguedad||'\u2014')+'</div>';
    datosP += '<div class="rpt-data-item"><strong>Empleados</strong>'+(d.empleados||'\u2014')+'</div>';
    if(repNombre) datosP += '<div class="rpt-data-item"><strong>Representante</strong>'+repNombre+(repNif?' \u00b7 '+repNif:'')+'</div>';
    datosP += '<div class="rpt-data-item"><strong>Tel\u00e9fono</strong>'+(d.tel1||'\u2014')+'</div>';
    datosP += '<div class="rpt-data-item"><strong>Email</strong>'+(d.email1||'\u2014')+'</div>';
    datosP += '<div class="rpt-data-item" style="grid-column:1/-1"><strong>Direcci\u00f3n</strong>'+(d.direccion?d.direccion+', '+d.cp+' '+d.localidad:'\u2014')+'</div>';
    datosP += '</div>';
  }
  html += rptSection(7,'Datos Personales','#414141', datosP||'', '');

  // ── P8: DATOS PROFESIONALES (emp/aut only) ─────────────────────
  var sectionNum = 8;
  if(d.tipo==='emp'||d.tipo==='aut'){
    var activ = d.actividad||'\u2014';
    var empl  = d.empleados||'\u2014';
    var antig = d.antiguedad||'\u2014';
    var vtosPer = d.vtos_per?'S\u00ed':'No';
    var autEmp  = d.aut_emp?'S\u00ed':'No';
    var profHtml = '<div class="rpt-data-grid">';
    profHtml += '<div class="rpt-data-item"><strong>Tipo</strong>'+{par:'Particular',emp:'Empresa',aut:'Aut\u00f3nomo'}[d.tipo]+'</div>';
    profHtml += '<div class="rpt-data-item"><strong>Actividad</strong>'+activ+'</div>';
    profHtml += '<div class="rpt-data-item"><strong>Antig\u00fcedad</strong>'+antig+'</div>';
    profHtml += '<div class="rpt-data-item"><strong>Empleados</strong>'+empl+'</div>';
    if(d.tipo==='aut') profHtml += '<div class="rpt-data-item"><strong>Tiene empleados</strong>'+autEmp+'</div>';
    if(d.tipo==='emp'||d.tipo==='aut') profHtml += '<div class="rpt-data-item"><strong>Vtos personales recogidos</strong>'+vtosPer+'</div>';
    profHtml += '</div>';
    html += rptSection(sectionNum,'Datos Profesionales','#5B21B6', profHtml, '');
    sectionNum++;
  }

  // ── OBSERVACIONES ───────────────────────────────────────────────
  var obsHtml = '';
  var obs = d.motivo_contacto;
  var accion = d.primera_accion;
  var fechaAcc = d.fecha_accion;
  var origenSw = d.origen_sw;
  var origenDet = d.origen_detalle;
  if(obs||accion||origenSw){
    obsHtml = '<div class="rpt-data-grid">';
    if(origenSw) obsHtml += '<div class="rpt-data-item"><strong>Origen</strong>'+origenSw+(origenDet?' \u00b7 '+origenDet:'')+'</div>';
    if(accion) obsHtml += '<div class="rpt-data-item"><strong>Primera acci\u00f3n</strong>'+accion+(fechaAcc?' \u00b7 '+fmtDate(fechaAcc):'')+'</div>';
    if(obs) obsHtml += '<div class="rpt-data-item" style="grid-column:1/-1"><strong>Notas / Motivo</strong><div style="margin-top:4px;white-space:pre-wrap;color:#202020">'+obs+'</div></div>';
    obsHtml += '</div>';
  }
  html += rptSection(sectionNum,'Observaciones y Notas','#6B7280', obsHtml, 'Sin observaciones registradas.');
  sectionNum++;

  // ── USUARIOS ────────────────────────────────────────────────────
  var usrHtml = '<div class="rpt-data-grid">';
  usrHtml += '<div class="rpt-data-item"><strong>Recoge datos</strong>'+(d.colab_recoge||'\u2014')+(d.colab_recoge_racf?' ('+d.colab_recoge_racf+')':'')+'</div>';
  usrHtml += '<div class="rpt-data-item"><strong>Asignado a</strong>'+(d.colab_asigna||'\u2014')+(d.colab_asigna_racf?' ('+d.colab_asigna_racf+')':'')+'</div>';
  if(d.fechaCreacion) usrHtml += '<div class="rpt-data-item"><strong>Fecha creaci\u00f3n</strong>'+new Date(d.fechaCreacion).toLocaleString('es-ES')+'</div>';
  usrHtml += '</div>';
  html += rptSection(sectionNum,'Datos de Usuarios','#374151', usrHtml, '');

  document.getElementById('reportBody').innerHTML = html;
  document.getElementById('reportOverlay').style.display = 'block';

  // ── Download HTML report + send via Outlook ────────────────────────
  var COLAB_EMAILS = {
    'M441819E': 'dany@optimizarte.com',
    'M354046Y': 'fadoua@optimizarte.com',
    'MA48168T': 'silvia@optimizarte.com'
  };
  var assignedRacf = colabAsigna;
  var toEmail = COLAB_EMAILS[assignedRacf] || 'optimizarte@optimizarte.com';
  var ccEmail = 'optimizarte@optimizarte.com';
  var subjStr = 'Informe Cliente Occident \u2014 '+(d.nombre_completo||'')+(d.nif_cif?' ('+d.nif_cif+')':'');
  var pillName2=function(p){
    var cfg=RAMO_CONFIG[p.type.replace(/-\d+$/,'')] || {name:p.type};
    var rd2=ramoData[p.key]||{};
    return rd2.nombre && p.type.replace(/-\d+$/,'')==='otro'?rd2.nombre:cfg.name;
  };
  var pillDetail2=function(p){
    var rd2=ramoData[p.key]||{};var parts=[];
    if(rd2.cia)parts.push('Cia:'+rd2.cia);
    if(rd2.vto)parts.push('VTO:'+fmtDate(rd2.vto));
    if(rd2.primaActual)parts.push('Prima:'+fmtMoney(rd2.primaActual));
    if(rd2.capital)parts.push('Capital:'+fmtMoney(rd2.capital));
    if(rd2.matricula)parts.push('Matr.:'+rd2.matricula);
    if(rd2.observacionesNota)parts.push('Obs.:'+rd2.observacionesNota.substring(0,80));
    return parts.length?' ['+parts.join(' | ')+']':'';
  };
  var yr = new Date().getFullYear();
  var candiNec = pills_necesita.map(function(p){return 'CANDI'+yr+(pillName2(p).toUpperCase().replace(/\s+/g,''));}).join(', ');
  var candiNc = pills_no_cambia.map(function(p){return 'CANDI'+yr+(pillName2(p).toUpperCase().replace(/\s+/g,''));}).join(', ');
  var sexoMap2={'H':'Hombre','M':'Mujer'};

  // ── Build full HTML file ──────────────────────────────────────────
  var reportHtml = document.getElementById('reportBody').innerHTML;
  var emailStyles = '<style>'+
    'body{margin:0;padding:16px;font-family:Arial,sans-serif;background:#f5f5f5}'+
    '.rpt-section{background:#fff;margin:0 0 3px;padding:20px 26px}'+
    '.rpt-header{background:#DC0028;padding:14px 26px;display:flex;align-items:center;justify-content:space-between;border-radius:8px 8px 0 0}'+
    '.rpt-section-hdr{display:flex;align-items:center;gap:10px;margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid #f0f0f0}'+
    '.rpt-section-hdr-label{font-family:Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.8px;text-transform:uppercase;color:#fff;padding:3px 10px;border-radius:4px}'+
    '.rpt-toggle-btn{display:inline-flex;align-items:center;gap:6px;margin-left:auto;padding:5px 14px;border-radius:20px;border:2px solid currentColor;font-family:Arial,sans-serif;font-size:11px;font-weight:700;cursor:pointer;letter-spacing:.3px;flex-shrink:0}'+
    '.rpt-data-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px 20px}'+
    '.rpt-data-item{font-size:12px;color:#414141;padding:4px 0}'+
    '.rpt-data-item strong{color:#202020;font-size:10px;display:block;text-transform:uppercase;letter-spacing:.5px;margin-bottom:1px}'+
    '.rpt-pill-card-detail{display:grid;grid-template-columns:1fr;gap:4px;padding:10px 12px}'+
    '.rpt-pill-detail-row{display:flex;flex-direction:column;padding:4px 0;border-bottom:1px solid #f5f5f5}'+
    '.rpt-pill-detail-label{color:#888;font-size:9px;text-transform:uppercase;letter-spacing:.4px;font-family:Arial,sans-serif;margin-bottom:2px}'+
    '.rpt-pill-detail-value{color:#202020;font-size:12px;font-family:Arial,sans-serif;font-weight:600}'+
    '.rpt-empty{font-size:12px;color:#bbb;font-style:italic;font-family:Arial,sans-serif}'+
    '.rpt-client-name{font-size:20px;font-weight:700;color:#fff;font-family:Arial,sans-serif}'+
    '.rpt-client-dni{font-size:13px;color:rgba(255,255,255,.85);font-family:Arial,sans-serif}'+
    '</style>';
  var safeName = (d.nombre_completo||'informe').replace(/[^a-zA-Z\u00C0-\u017E0-9 ]/g,'_');
  var reportFileName = 'Informe_'+safeName+'_'+new Date().toISOString().slice(0,10)+'.html';
  var fullHtmlDoc = '<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><title>Informe '+
    (d.nombre_completo||'Cliente')+'</title>'+emailStyles+'</head><body>'+reportHtml+'</body></html>';

  // Send email SYNCHRONOUSLY (must be in user gesture context — before any async)
  var altaPath = 'file:///C:/Users/primo/OneDrive%20-%20OPTIMIZARTE%203.0%20-%20SCO/!!IA/!!!PortalOccident/hub-optimizarte/AltaClientesLocal/' + encodeURIComponent(reportFileName);
  var mailBody = encodeURIComponent('Descargar informe: ' + altaPath);
  var mailSubj = encodeURIComponent(subjStr);
  var mailtoLink = 'mailto:'+toEmail+'?cc='+encodeURIComponent(ccEmail)+'&subject='+mailSubj+'&body='+mailBody;
  var mailA = document.createElement('a');
  mailA.href = mailtoLink; mailA.style.display='none';
  document.body.appendChild(mailA); mailA.click();
  setTimeout(function(){ document.body.removeChild(mailA); }, 500);

  // Download HTML report — save to AltaClientesLocal folder
  var htmlBlob = new Blob([fullHtmlDoc], {type:'text/html;charset=utf-8'});
  saveReportToDir(htmlBlob, reportFileName).then(function(savedName) {
    showToast('\ud83d\udcbe Informe guardado \u2014 \ud83d\udce7 Abriendo Outlook para '+toEmail,'');
  });
}

function printReport() {
  var content = '<html><head><meta charset="utf-8"><title>Informe Cliente Occident</title>';
  content += '<style>';
  content += 'body{margin:0;padding:0;font-family:Arial,sans-serif;background:#f5f5f5}';
  content += '.rpt-section{background:#fff;margin:0 0 3px;padding:20px 26px;page-break-inside:avoid}';
  content += '.rpt-header{background:#DC0028;padding:18px 26px 14px;display:flex;align-items:center;justify-content:space-between}';
  content += '.rpt-section-hdr{display:flex;align-items:center;gap:10px;margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid #f0f0f0}';
  content += '.rpt-section-hdr-label{font-family:Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.8px;text-transform:uppercase;color:#fff;padding:3px 10px;border-radius:4px}';
  content += '.rpt-section-num{font-family:Arial,sans-serif;font-size:11px;font-weight:700;color:#888;margin-left:auto}';
  content += '.rpt-pill-row{display:flex;align-items:flex-start;gap:12px;padding:10px 0;border-bottom:1px solid #f5f5f5}';
  content += '.rpt-pill-row:last-child{border-bottom:none}';
  content += '.rpt-pill-badge{display:inline-flex;align-items:center;gap:6px;padding:5px 12px;border-radius:20px;font-size:12px;font-weight:700;white-space:nowrap;flex-shrink:0}';
  content += '.rpt-data-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px 20px}';
  content += '.rpt-data-item{font-size:12px;color:#414141}';
  content += '.rpt-data-item strong{color:#202020;font-size:10px;display:block;text-transform:uppercase;letter-spacing:.5px;margin-bottom:1px}';
  content += '.rpt-client-name{font-size:20px;font-weight:700;color:#fff}';
  content += '.rpt-client-dni{font-size:13px;color:rgba(255,255,255,.85)}';
  content += '.rpt-empty{font-size:12px;color:#bbb;font-style:italic}';
  content += '@page{margin:10mm}';
  content += '</style></head><body>';
  content += document.getElementById('reportBody').innerHTML;
  content += '</body></html>';
  var w = window.open('','_blank','width=860,height=900');
  w.document.write(content);
  w.document.close();
  w.focus();
  setTimeout(function(){ w.print(); },400);
}
function downloadReportHTML() {
  var d = collectData();
  var reportContent = document.getElementById('reportBody').innerHTML;
  var styles = Array.from(document.querySelectorAll('style')).map(function(s){return s.outerHTML;}).join('\n');
  var fullHtml = '<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">'+
    '<title>Informe Cliente Occident \u2014 '+(d.nombre_completo||'Cliente')+'</title>'+styles+
    '<style>body{margin:0;background:#f5f5f5;font-family:Arial,sans-serif}</style>'+
    '</head><body>'+reportContent+'</body></html>';
  var blob = new Blob([fullHtml], {type:'text/html;charset=utf-8'});
  var safe = (d.nombre_completo||'informe').replace(/[^a-zA-Z\u00C0-\u017E0-9 ]/g,'_');
  var filename = 'Informe_'+safe+'_'+new Date().toISOString().slice(0,10)+'.html';
  saveReportToDir(blob, filename).then(function(savedName){
    showToast('\ud83d\udcbe Informe guardado'+(savedName?' \u2013 '+savedName:''),'');
  });
}
// ─── END REPORT GENERATOR ─────────────────────────────────

// ─── INTEGRACIÓN GESTIONA ─────────────────────────────────
function registrarEnGestiona() {
  var d = lastData;
  if (!d) {
    showToast('⚠️ No hay datos de cliente. Registra el cliente primero.', 'warn');
    return;
  }

  var isEmp = (d.tipo === 'emp');
  var isAut = (d.tipo === 'aut');

  // Fecha nacimiento: nuestro formato yyyy-mm-dd → Gestiona dd/mm/aaaa
  var fnacRaw = isEmp ? '' : (isAut ? (d.aut_fnac || '') : (d.par_fnac || ''));
  var fnac = '';
  if (fnacRaw && fnacRaw.indexOf('-') !== -1) {
    var parts = fnacRaw.split('-');
    if (parts.length === 3) fnac = parts[2] + '/' + parts[1] + '/' + parts[0];
  } else {
    fnac = fnacRaw;
  }

  // Teléfono: quitar prefijo +34
  var tel = (d.tel1 || '').replace(/^\+34[\s-]?/, '').trim();

  var ap1 = isAut ? (d.aut_ap1 || '') : (d.par_ap1 || '');
  var ap2 = isAut ? (d.aut_ap2 || '') : (d.par_ap2 || '');

  var payload = {
    // "Tipo de persona" en Gestiona: Masculino | Femenino | Juridica
    tipo_persona: isEmp ? 'Juridica' : (d.sexo === 'H' ? 'Masculino' : 'Femenino'),
    nombre:       isEmp ? (d.emp_razon || '') : (isAut ? (d.aut_nombre || '') : (d.par_nombre || '')),
    // Gestiona usa un único campo "Apellidos" (no apellido1 + apellido2)
    apellidos:    isEmp ? '' : [ap1, ap2].filter(Boolean).join(' '),
    // Identificador: siempre "NIF (DNI, CIF, NIE)" — cubre NIF, NIE y CIF
    id_valor:     isEmp ? (d.emp_cif || '') : (isAut ? (d.aut_nif || '') : (d.par_nif || '')),
    fnac:         fnac,
    tel:          tel,
    email:        d.email1 || ''
  };

  var encoded = btoa(unescape(encodeURIComponent(JSON.stringify(payload))));
  var url = 'https://gestiona.gco.global/gestionadministrativa/gestiondeclientes/gestionclientesoccident#ac=' + encoded;

  // Abrir Gestiona en pestaña nueva — llamado desde onclick directo (no async),
  // por lo que el bloqueador de popups lo permite.
  var w = window.open(url, '_blank');
  if (w) {
    showToast('🏢 Abriendo Gestiona…', '');
  } else {
    // Popup bloqueado: ofrecemos copiar la URL como alternativa
    showToast('⚠️ El navegador bloqueó la ventana emergente. Permite popups para esta página.', 'warn');
    try { navigator.clipboard.writeText(url); } catch(e) {}
  }
}
// ─── END INTEGRACIÓN GESTIONA ──────────────────────────────

// ─── TASCA 9: DOBLE CLIC CARDS PER MOSTRAR/AMAGAR ──────────
var cardsOcultas = false; // Estado toggle

// Inicialitzar event listeners doble clic
(function() {
  // Afegir doble clic als 3 tipus (par, emp, aut)
  ['tipo-par', 'tipo-emp', 'tipo-aut'].forEach(function(id) {
    var el = document.getElementById(id);
    if (el) {
      el.addEventListener('dblclick', toggleCardsVisibility);
      // Afegir cursor help per indicar que es pot fer doble clic
      el.style.cursor = 'help';
      el.title = 'Doble clic per mostrar/amagar cards de dades personals';
    }
  });
})();

function toggleCardsVisibility() {
  cardsOcultas = !cardsOcultas;
  
  // Selectors de les 2 cards a amagar (només Identificación i Contacto)
  // Card 0 = Tipo Cliente (no tocar)
  // Card 1 = Identificación
  // Card 2 = Contacto
  // Card 3 = Dirección (SEMPRE VISIBLE)
  // Card 4 = Bancarios (SEMPRE VISIBLE)
  var allCards = document.querySelectorAll('.card');
  var cardsToToggle = [allCards[1], allCards[2]]; // Només 2 cards!
  
  cardsToToggle.forEach(function(card) {
    if (card) {
      card.style.display = cardsOcultas ? 'none' : 'block';
    }
  });
  
  // Feedback visual temporal
  mostrarToastToggle(cardsOcultas ? '🔒 Dades personals ocultades' : '👁️ Dades personals visibles');
  
  // Actualitzar progress bar
  updProg();
}

function mostrarToastToggle(msg) {
  // Crear toast temporal per feedback
  var toast = document.createElement('div');
  toast.textContent = msg;
  toast.style.cssText = 'position:fixed;bottom:20px;left:50%;transform:translateX(-50%);' +
    'background:#1F2937;color:#fff;padding:10px 18px;border-radius:8px;' +
    'font-size:13px;font-weight:600;z-index:9999;' +
    'box-shadow:0 10px 30px -10px rgba(0,0,0,.4);' +
    'animation:fadeIn 0.2s ease-out';
  document.body.appendChild(toast);
  setTimeout(function() { 
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s';
    setTimeout(function() { toast.remove(); }, 300);
  }, 2000);
}
// ─── END DOBLE CLIC CARDS ──────────────────────────────────

// ─── FUNCIÓ BOTÓ "NOU CLIENT" ─────────────────────────────
// ─── MODE DE VISIBILITAT DE CARDS ─────────────────────────
// 'initial'                → només card 0 visible
// 'nou-client'             → totes les cards visibles (updateVisibility per condicionals)
// 'client-loaded-directe'  → totes visibles excepte Identificación + Contacto (amagades, dades preservades)
// 'client-loaded-fitxa'    → fitxa al card 0; Identificación + Contacto amagades; resta visibles
//
// PRINCIPI: NO modifiquem mai innerHTML d'aquestes cards.
// Sempre fem display:none / display:'' per preservar valors dels inputs.
// L'usuari mostra/amaga Ident+Contacto fent DOBLE-CLIC sobre una opció de tipus (par/emp/aut).
var _currentFormMode = 'initial';

function _getAllFormCards() {
  var form = document.getElementById('altaForm');
  if (!form) return [];
  return Array.prototype.slice.call(form.querySelectorAll(':scope > .card'));
}

function _showCard(card) { if (card) card.style.display = ''; }
function _hideCard(card) { if (card) card.style.display = 'none'; }

// Doble-clic sobre una opció de tipus client → toggle Ident + Contacto
// Cridada des de #tipo-par / #tipo-emp / #tipo-aut (ondblclick)
function toggleIdentContactCards() {
  var ident = document.getElementById('card-identificacion');
  var contact = document.getElementById('card-contacto');
  if (!ident || !contact) return;
  var anyHidden = (ident.style.display === 'none') || (contact.style.display === 'none');
  if (anyHidden) {
    ident.style.display = '';
    contact.style.display = '';
    // Scroll suau cap a Identificación
    setTimeout(function(){ ident.scrollIntoView({behavior:'smooth', block:'center'}); }, 50);
    console.log('[toggle Ident+Contacto] → MOSTRATS');
  } else {
    // Si estem en mode client-loaded, podem amagar de nou
    if (_currentFormMode === 'client-loaded-directe' || _currentFormMode === 'client-loaded-fitxa') {
      ident.style.display = 'none';
      contact.style.display = 'none';
      console.log('[toggle Ident+Contacto] → AMAGATS');
    } else {
      console.log('[toggle Ident+Contacto] → ja visibles, mode actual no permet amagar');
    }
  }
}

// Mantenim toggleHiddenCard per cridada externa (Console o futurs botons)
function toggleHiddenCard(cardId) {
  var card = document.getElementById(cardId);
  if (!card) return;
  if (card.style.display === 'none') {
    card.style.display = '';
    setTimeout(function(){ card.scrollIntoView({behavior:'smooth', block:'center'}); }, 50);
  } else {
    card.style.display = 'none';
  }
}

function setFormMode(mode) {
  _currentFormMode = mode;
  var cardTipo = document.getElementById('card-tipo-cliente');
  var cardIdent = document.getElementById('card-identificacion');
  var cardContact = document.getElementById('card-contacto');
  var allCards = _getAllFormCards();

  if (mode === 'initial') {
    allCards.forEach(_hideCard);
    _showCard(cardTipo);
  } else if (mode === 'nou-client') {
    allCards.forEach(_showCard);
    if (typeof updateVisibility === 'function') updateVisibility();
  } else if (mode === 'client-loaded-directe' || mode === 'client-loaded-fitxa') {
    allCards.forEach(_showCard);
    if (typeof updateVisibility === 'function') updateVisibility();
    _hideCard(cardIdent);
    _hideCard(cardContact);
  }

  console.log('[FormMode] →', mode);
}

// ─── TOGGLE MODAL OD + FITXA-MODAL CARD 0 ───────────────────
// Toggle ON  (defecte) → en seleccionar un client existent, mostra fitxa al card 0 (cal confirmar càrrega)
// Toggle OFF           → en seleccionar un client existent, carrega directe al formulari
var _showOdModal = true;

function _readShowOdModal() {
  try {
    var v = localStorage.getItem('optiAlta_showOdModal');
    return v === null ? true : v === '1';
  } catch(e) { return true; }
}
function _writeShowOdModal(on) {
  try { localStorage.setItem('optiAlta_showOdModal', on ? '1' : '0'); } catch(e) {}
}

function _applyToggleVisual(on) {
  var track = document.getElementById('toggleModalODTrack');
  var knob  = document.getElementById('toggleModalODKnob');
  var lbl   = document.getElementById('toggleModalODLbl');
  if (track) track.style.background = on ? '#10B981' : '#6B7280';
  if (knob)  knob.style.left = on ? '18px' : '2px';
  if (lbl)   lbl.textContent = on ? 'Fitxa' : 'Directe';
}

function onToggleModalOD(checked) {
  _showOdModal = !!checked;
  _writeShowOdModal(_showOdModal);
  _applyToggleVisual(_showOdModal);
}

function initToggleModalOD() {
  _showOdModal = _readShowOdModal();
  var cb = document.getElementById('toggleModalOD');
  if (cb) cb.checked = _showOdModal;
  _applyToggleVisual(_showOdModal);
}

// Renderitza fitxa del client substituint el card-tipo-cliente (card 0)
// L'usuari ha de confirmar (Carregar) o cancel·lar.
function showClientFitxaCard0(client) {
  var cardTipo = document.getElementById('card-tipo-cliente');
  if (!cardTipo) { loadClientIntoForm(client); return; }

  // Origen visual
  var origen = '';
  var hasOD = client._source && (client._source === 'od' || client._source === 'both');
  var hasCRM = client._source && (client._source === 'crm' || client._source === 'both');
  if (hasOD && hasCRM) origen = '<span style="background:#10B981;color:#fff;padding:3px 9px;border-radius:5px;font-size:11px;font-weight:700;letter-spacing:.3px">📁 OD + 💼 CRM</span>';
  else if (hasOD) origen = '<span style="background:#3B82F6;color:#fff;padding:3px 9px;border-radius:5px;font-size:11px;font-weight:700;letter-spacing:.3px">📁 ONEDRIVE</span>';
  else if (hasCRM) origen = '<span style="background:#DC0028;color:#fff;padding:3px 9px;border-radius:5px;font-size:11px;font-weight:700;letter-spacing:.3px">💼 CRM</span>';

  // Camps formulari estàndard
  var nom = client.nombre_completo || client.nom ||
            ((client.par_nombre||client.aut_nombre||client.emp_razon||'') + ' ' + (client.par_ap1||client.aut_ap1||'')).trim() || '—';
  // Si nom és null (CRM no l'ha extret), provar dels rebuts
  if (!nom || nom === '—' || nom === 'null') {
    var fromReceipts = (typeof _extractNameFromReceipts === 'function') ? _extractNameFromReceipts(client) : '';
    if (fromReceipts) nom = fromReceipts;
  }
  var nif = client.nif_cif || client.par_nif || client.aut_nif || client.emp_cif || client.nif || '—';
  var tel = client.tel1 || client.telefon || '—';
  var email = client.email1 || client.email2 || client.email || '—';
  var loc = (client.localidad || client.muni || '') + (client.cp ? ' (' + client.cp + ')' : '');
  if (!loc) loc = '—';
  var adreca = client.dir || client.direccion || client.adreca || '';
  var refp = client._refnumpers ? '<div style="font-size:11px;color:#999;margin-top:2px">RefPers: ' + client._refnumpers + '</div>' : '';
  var enriched = client._enrichedAt ? '<div style="font-size:10px;color:#10B981;margin-top:2px">✓ Enriquit ' + new Date(client._enrichedAt).toLocaleDateString('ca-ES') + '</div>' : '';

  // ── Info addicional CRM (només si està present) ──
  var infoCRM = [];
  if (client.eclient) infoCRM.push('<span style="background:#FEF3C7;color:#92400E;padding:2px 7px;border-radius:4px;font-size:11px;font-weight:600">⭐ eClient: ' + client.eclient + '</span>');
  if (client.antiguitat) infoCRM.push('<span style="background:#E0E7FF;color:#3730A3;padding:2px 7px;border-radius:4px;font-size:11px;font-weight:600">📅 Alta: ' + client.antiguitat + '</span>');
  if (client.segment) infoCRM.push('<span style="background:#DBEAFE;color:#1E40AF;padding:2px 7px;border-radius:4px;font-size:11px;font-weight:600">🎯 ' + client.segment + '</span>');
  if (client.rgpd && /si|s[ií]/i.test(client.rgpd)) infoCRM.push('<span style="background:#DCFCE7;color:#166534;padding:2px 7px;border-radius:4px;font-size:11px;font-weight:600">✓ RGPD OK</span>');
  if (client.idioma) infoCRM.push('<span style="background:#F3F4F6;color:#374151;padding:2px 7px;border-radius:4px;font-size:11px;font-weight:600">🌐 ' + client.idioma + '</span>');
  if (client.professio) infoCRM.push('<span style="background:#FCE7F3;color:#9F1239;padding:2px 7px;border-radius:4px;font-size:11px;font-weight:600">💼 ' + client.professio + '</span>');

  // Estadístiques CRM
  var stats = [];
  if (client.totalProductes !== undefined && client.totalProductes !== null) stats.push({n: client.totalProductes, l: 'Productes'});
  if (client.totalPrimes) stats.push({n: client.totalPrimes, l: 'Primes anuals'});
  if (Array.isArray(client.polisses)) stats.push({n: client.polisses.length, l: 'Pòlisses'});
  if (Array.isArray(client.rebuts)) stats.push({n: client.rebuts.length, l: 'Rebuts'});
  if (Array.isArray(client.sinistres)) stats.push({n: client.sinistres.length, l: 'Sinistres'});

  var statsHtml = '';
  if (stats.length) {
    statsHtml = '<div style="display:flex;gap:8px;margin-top:10px;padding:10px;background:#F9FAFB;border-radius:6px;border-left:3px solid #DC0028">';
    stats.forEach(function(s){
      statsHtml += '<div style="flex:1;text-align:center"><div style="font-size:16px;font-weight:700;color:#202020;font-family:Inter,sans-serif">' + s.n + '</div><div style="font-size:10px;color:#666;text-transform:uppercase;letter-spacing:.3px">' + s.l + '</div></div>';
    });
    statsHtml += '</div>';
  }

  var infoCRMHtml = infoCRM.length ? '<div style="display:flex;flex-wrap:wrap;gap:5px;margin-top:10px;padding-top:10px;border-top:1px dashed #E5E7EB">' + infoCRM.join('') + '</div>' : '';

  // Guardar referència del client per al botó "Carregar"
  window._pendingClient = client;

  var html =
    '<div class="card-body" style="padding:18px 22px">' +
      '<div style="display:flex;align-items:center;gap:12px;margin-bottom:14px;padding-bottom:12px;border-bottom:1px solid #eee">' +
        '<div style="width:42px;height:42px;border-radius:50%;background:#FEF3C7;display:flex;align-items:center;justify-content:center;font-size:22px">👤</div>' +
        '<div style="flex:1">' +
          '<div style="font-family:Poppins,sans-serif;font-size:16px;font-weight:700;color:#202020;display:flex;align-items:center;gap:8px">' + nom + '</div>' +
          '<div style="font-size:12px;color:#666;margin-top:2px">Client existent — revisa abans de carregar</div>' +
        '</div>' +
        origen +
      '</div>' +
      '<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px 24px;font-size:13px">' +
        '<div><div style="font-size:10px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px">NIF / CIF</div><div style="font-weight:600;color:#202020">' + nif + '</div>' + refp + '</div>' +
        '<div><div style="font-size:10px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px">Telèfon</div><div style="font-weight:600;color:#202020">' + tel + '</div>' + enriched + '</div>' +
        '<div><div style="font-size:10px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px">Email</div><div style="font-weight:600;color:#202020;word-break:break-all">' + email + '</div></div>' +
        '<div><div style="font-size:10px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px">Localitat</div><div style="font-weight:600;color:#202020">' + loc + '</div></div>' +
        (adreca ? '<div style="grid-column:1/-1"><div style="font-size:10px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px">Adreça</div><div style="font-weight:600;color:#202020">' + adreca + '</div></div>' : '') +
      '</div>' +
      statsHtml +
      infoCRMHtml +
      '<div style="display:flex;gap:10px;justify-content:flex-end;margin-top:18px;padding-top:14px;border-top:1px solid #eee">' +
        '<button type="button" onclick="cancelClientFitxaCard0()" style="padding:9px 16px;border-radius:8px;border:1.5px solid #D1D5DB;background:#fff;font-size:12.5px;font-weight:600;color:#414141;cursor:pointer">✕ Cancel·lar</button>' +
        '<button type="button" onclick="confirmClientFitxaCard0()" style="padding:9px 18px;border-radius:8px;border:none;background:#DC0028;color:#fff;font-size:12.5px;font-weight:700;cursor:pointer">📝 Carregar al formulari</button>' +
      '</div>' +
    '</div>';

  // Guardem el contingut original per restaurar després
  if (!cardTipo._originalHTML) cardTipo._originalHTML = cardTipo.innerHTML;
  cardTipo.innerHTML = html;
  cardTipo.style.display = 'block';
  // Amaguem el contenidor de suggerencies
  var sug = document.getElementById('searchSuggestions');
  if (sug) sug.style.display = 'none';
  // Scroll suau cap a la fitxa
  cardTipo.scrollIntoView({behavior:'smooth', block:'start'});
}

function confirmClientFitxaCard0() {
  var c = window._pendingClient;
  window._pendingClient = null;
  // Restaurar card 0 original
  var cardTipo = document.getElementById('card-tipo-cliente');
  if (cardTipo && cardTipo._originalHTML) {
    cardTipo.innerHTML = cardTipo._originalHTML;
    cardTipo._originalHTML = null;
  }
  if (c) {
    loadClientIntoForm(c);
    // Després de carregar: el card 0 ja és normal, passem a mode directe (identificación/contacto strip)
    setFormMode('client-loaded-directe');
  }
}

function cancelClientFitxaCard0() {
  window._pendingClient = null;
  // Restaurar card 0 original
  var cardTipo = document.getElementById('card-tipo-cliente');
  if (cardTipo && cardTipo._originalHTML) {
    cardTipo.innerHTML = cardTipo._originalHTML;
    cardTipo._originalHTML = null;
  }
  // Reset cercador
  var si = document.getElementById('searchInput');
  if (si) si.value = '';
  // Tornar a estat inicial
  setFormMode('initial');
}
// ──────────────────────────────────────────────────────────

function activarNouClient() {
  // Reset de la fitxa al card 0 si hi era
  var cardTipo = document.getElementById('card-tipo-cliente');
  if (cardTipo && cardTipo._originalHTML) {
    cardTipo.innerHTML = cardTipo._originalHTML;
    cardTipo._originalHTML = null;
  }
  window._pendingClient = null;
  // Netejar cercador
  var si = document.getElementById('searchInput');
  if (si) si.value = '';
  // Ocultar suggerencies
  var sug = document.getElementById('searchSuggestions');
  if (sug) sug.style.display = 'none';

  // Cridar funció reset existent
  resetFormulariNouClient();

  // Mostrar TOTES les cards (mode 'nou-client')
  setFormMode('nou-client');

  // Feedback visual
  mostrarToastToggle('✨ Nou client activat — formulari complet');

  // Focus al primer camp
  setTimeout(function() {
    var primer = document.getElementById('par-nombre');
    if (primer) primer.focus();
  }, 100);
}
// ──────────────────────────────────────────────────────────

// ─── INICIALITZACIÓ: CARREGAR CLIENTS ─────────────────────
// FONT 1A: parent.OD directe (mateix origen — extension popup)
// FONT 1B: postMessage al parent (cross-origin — iframe GitHub Pages ↔ CRM)
// FONT 2 : clientesDir (carpeta local File System Access)

function _requestIndexViaPostMessage(timeoutMs, _attempt) {
  timeoutMs = timeoutMs || 4000;
  _attempt = _attempt || 1;
  var maxAttempts = 4;       // 4 intents
  var retryDelayMs = 2000;   // 2s entre intents

  return new Promise(function(resolve) {
    var done = false;
    function handler(ev) {
      if (!ev.data || typeof ev.data !== 'object') return;
      if (ev.data.type !== 'opticrm_clients_index_response') return;
      if (done) return;
      done = true;
      window.removeEventListener('message', handler);
      var n = ev.data.data && ev.data.data.clients ? Object.keys(ev.data.data.clients).length : 0;
      console.log('☁️ [PM] Resposta del parent (intent ' + _attempt + '/' + maxAttempts + '): ' + n + ' clients');
      resolve(ev.data.data || null);
    }
    window.addEventListener('message', handler);
    try {
      if (window.parent && window.parent !== window) {
        window.parent.postMessage({type:'opticrm_request_clients_index'}, '*');
        console.log('☁️ [PM] Petició enviada al parent (intent ' + _attempt + '/' + maxAttempts + ')');
      } else {
        console.log('☁️ [PM] No hi ha parent (form standalone)');
        done = true;
        resolve(null);
        return;
      }
    } catch(e) {
      console.warn('☁️ [PM] No es pot enviar postMessage:', e && e.message);
    }
    setTimeout(function(){
      if (done) return;
      done = true;
      window.removeEventListener('message', handler);
      if (_attempt < maxAttempts) {
        console.log('☁️ [PM] Timeout intent ' + _attempt + ' — reintent en ' + retryDelayMs + 'ms');
        setTimeout(function(){
          _requestIndexViaPostMessage(timeoutMs, _attempt + 1).then(resolve);
        }, retryDelayMs);
      } else {
        console.warn('☁️ [PM] Sense resposta del parent després de ' + maxAttempts + ' intents — silentsync potser no actiu');
        resolve(null);
      }
    }, timeoutMs);
  });
}

// Demanar dades completes d'un client (cli_crm_<NIF>.json) al parent via postMessage
function _requestClientDetailsViaPostMessage(nif, timeoutMs) {
  timeoutMs = timeoutMs || 5000;
  return new Promise(function(resolve) {
    if (!nif) { resolve(null); return; }
    var done = false;
    function handler(ev) {
      if (!ev.data || typeof ev.data !== 'object') return;
      if (ev.data.type !== 'opticrm_client_details_response') return;
      if (ev.data.nif !== nif) return;
      if (done) return;
      done = true;
      window.removeEventListener('message', handler);
      console.log('☁️ [PM] Detalls client rebuts per NIF ' + nif + ':', ev.data.data ? 'OK' : 'BUIT');
      resolve(ev.data.data || null);
    }
    window.addEventListener('message', handler);
    try {
      if (window.parent && window.parent !== window) {
        window.parent.postMessage({ type:'opticrm_request_client_details', nif: nif }, '*');
        console.log('☁️ [PM] Petició detalls enviada per NIF ' + nif);
      } else {
        done = true;
        resolve(null);
        return;
      }
    } catch(e) {
      console.warn('☁️ [PM] Error enviant petició detalls:', e && e.message);
    }
    setTimeout(function() {
      if (done) return;
      done = true;
      window.removeEventListener('message', handler);
      console.warn('☁️ [PM] Timeout detalls per NIF ' + nif);
      resolve(null);
    }, timeoutMs);
  });
}

async function initLoadClients() {
  var loadedOD = 0;
  var loadedDir = 0;
  // Buidar la llista per evitar duplicats en refrescos periòdics
  allClients = [];

  // ── FONT 1A: parent.OD directament (només funciona same-origin) ───
  var indexData = null;
  try {
    var od = null;
    try {
      if (window.parent && window.parent !== window && window.parent.OD && window.parent.OD.isReady && window.parent.OD.isReady()) od = window.parent.OD;
    } catch(e) { /* cross-origin: ignorem */ }
    try { if (!od && window.OD && window.OD.isReady && window.OD.isReady()) od = window.OD; } catch(e) {}

    if (od && typeof od.loadFile === 'function') {
      indexData = await od.loadFile('opticrm-clients-index.json');
      if (indexData) console.log('☁️ [OD-direct] Lectura directa parent.OD OK');
    }
  } catch(e) {
    console.log('☁️ [OD-direct] Falla (probable cross-origin):', e && e.message);
  }

  // ── FONT 1B: postMessage (cross-origin fallback) ───────────────────
  if (!indexData) {
    indexData = await _requestIndexViaPostMessage();
  }

  // ── Processar índex rebut (sigui per FONT 1A o 1B) ─────────────────
  if (indexData && indexData.clients && typeof indexData.clients === 'object') {
    var keys = Object.keys(indexData.clients);
    for (var k = 0; k < keys.length; k++) {
      var src = indexData.clients[keys[k]];
      if (!src) continue;
      var enriched = !!(src.tel1 || src.email1);
      var c = {
        _filename: 'od:' + keys[k],
        _source: enriched ? 'both' : 'crm',
        _refnumpers: src.refnumpers || keys[k],
        _enrichedAt: src._enrichedAt || null,
        _odIndex: true,
        nombre_completo: src.nom || src.nombre_completo || '',
        nif_cif: src.nif || '',
        tel1: src.tel1 || '',
        email1: src.email1 || '',
        tipo_label: src.tipo_label || (src.nif && /^[A-Z]/i.test(src.nif) && src.nif.length===9 ? 'EMP' : 'PAR')
      };
      allClients.push(c);
      loadedOD++;
    }
    console.log('☁️ Índex processat: ' + loadedOD + ' clients (FONT: ' + (indexData._via || 'auto') + ')');
  } else {
    console.log('☁️ Índex OD buit o no rebut');
  }

  // ── FONT 2: clientesDir (carpeta local, opcional) ──────────────────
  if (clientesDir) {
    try {
      var entries = [];
      for await (var entry of clientesDir.values()) {
        if (entry.kind === 'file' && entry.name.endsWith('.json')) entries.push(entry);
      }
      console.log('📁 [DIR] Trobats ' + entries.length + ' fitxers JSON');
      for (var i = 0; i < entries.length; i++) {
        try {
          var fh = await clientesDir.getFileHandle(entries[i].name);
          var file = await fh.getFile();
          var data = JSON.parse(await file.text());
          data._filename = entries[i].name;
          var existingIdx = -1;
          if (data.nif_cif) {
            for (var j = 0; j < allClients.length; j++) {
              if (allClients[j].nif_cif && allClients[j].nif_cif.toUpperCase() === data.nif_cif.toUpperCase()) { existingIdx = j; break; }
            }
          }
          if (existingIdx >= 0) {
            data._source = 'both';
            data._refnumpers = allClients[existingIdx]._refnumpers || null;
            allClients[existingIdx] = data;
          } else {
            data._source = 'od';
            allClients.push(data);
          }
          loadedDir++;
        } catch(e) { console.error('Error carregant ' + entries[i].name, e); }
      }
      console.log('📁 [DIR] Carregats ' + loadedDir + ' clients de la carpeta');
    } catch(e) { console.error('📁 [DIR] Error:', e); }
  }

  var total = allClients.length;
  console.log('✅ Total clients disponibles: ' + total + ' (☁️ ' + loadedOD + ' · 📁 ' + loadedDir + ')');
  if (total > 0 && typeof showToast === 'function' && !window._clientsLoadedToastShown) {
    showToast(total + ' clients carregats (☁️ ' + loadedOD + ' · 📁 ' + loadedDir + ')', 'success');
    window._clientsLoadedToastShown = true;
  }
}

// Cridar inicialització al carregar la pàgina
window.addEventListener('DOMContentLoaded', function() {
  console.log('🚀 Formulari Alta Clients inicialitzat');

  // Inicialitzar toggle Modal OD ON/OFF (persistència localStorage)
  try { initToggleModalOD(); } catch(e) { console.warn('initToggleModalOD:', e); }

  // ESTAT INICIAL: només card 0 visible, resta amagada
  try { setFormMode('initial'); } catch(e) { console.warn('setFormMode initial:', e); }

  // Carregar clients (OD index + carpeta local). NO depèn només de clientesDir
  // perquè dins l'iframe OPTICRM la font primària és parent.OD.loadFile()
  setTimeout(function() {
    initLoadClients();
  }, 500);

  // Refresc periòdic de l'índex cada 10 min (alineat amb sync silenciós d'app.js)
  if (!window._clientsRefreshTimer) {
    window._clientsRefreshTimer = setInterval(function() {
      console.log('🔄 Refrescant índex de clients (10 min tick)');
      initLoadClients();
    }, 10 * 60 * 1000);
  }
});
// ──────────────────────────────────────────────────────────

