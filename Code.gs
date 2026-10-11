/**
 * CRM Fermetal — servidor (Google Apps Script)
 * ------------------------------------------------------------
 * Guarda los datos en la hoja "CRM Fermetal - Base de datos" (se crea sola).
 * Implementar como Aplicación web: Ejecutar como "Yo", acceso "Cualquier persona".
 * La seguridad la da el inicio de sesión: cada vendedor solo recibe sus clientes.
 */

const CRM = {
  NOMBRE_HOJA: 'CRM Fermetal - Base de datos',
  ADMIN_USUARIO: 'ivan.marin',
  ADMIN_NOMBRE: 'Ivan Marín',
  COTIZADOR_URL: 'https://script.google.com/macros/s/AKfycbzBztfVcAT_othSTnsW8V9v3pyCbGyOTuZvtn-R2-Cdu9Pllxodzt3qC0sN7raIpttI/exec',
  SESION_DIAS: 30,
  HASH_VUELTAS: 300,
  MAX_INTENTOS: 5,
  BLOQUEO_MIN: 10,
  MAX_CAMBIOS: 5000
};

const TABLAS = {
  Clientes: ['id','rif','nombre','contacto','telefono','email','direccion','ciudad','condicion','estado','vendedorId','notas','ultimoContacto','proximaAccion','proximaNota','origen','creado','modificado','modificadoPor','ver'],
  Usuarios: ['id','usuario','nombre','rol','email','telefono','activo','hash','salt','creado','ultimoAcceso'],
  Seguimientos: ['id','clienteId','fecha','tipo','nota','proximaFecha','usuarioId','usuarioNombre','creado'],
  Cambios: ['fecha','usuarioId','usuarioNombre','clienteId','clienteNombre','accion','detalle'],
  Config: ['clave','valor']
};

// Campos que el vendedor puede actualizar de sus clientes
const CAMPOS_VENDEDOR = ['rif','nombre','contacto','telefono','email','direccion','ciudad','condicion','estado','notas','proximaAccion','proximaNota'];
// Campos que además puede tocar el administrador
const CAMPOS_ADMIN = CAMPOS_VENDEDOR.concat(['vendedorId','origen']);
// Campos que vienen del cotizador
const CAMPOS_COTIZADOR = ['rif','nombre','contacto','telefono','email','direccion','condicion'];
const ESTADOS = ['Activo','Prospecto','Inactivo'];
// Datos de ejemplo de la plantilla del cotizador que no son reales
const RELLENOS = ['nombre apellido','correo@ejemplo.com','correo@empresa.com'];

/* ============================================================
 * Entrada HTTP
 * ============================================================ */
function doGet(e) {
  const a = e && e.parameter && e.parameter.action;
  if (a === 'ping') return json_({ ok: true, app: 'crm' });
  return ContentService.createTextOutput('Servidor del CRM activo.');
}

function doPost(e) {
  let p;
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'Solicitud inválida' }); }
  try {
    const r = enrutar_(p);
    r.ok = true;
    return json_(r);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err), code: err && err.code || '' });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function fallo_(msg, code) { const e = new Error(msg); e.code = code || ''; throw e; }

const ACCIONES = {
  // públicas
  ping: { pub: true }, login: { pub: true }, crearClave: { pub: true },
  // con sesión
  logout: {}, datos: {}, cambiarClave: {},
  guardarCliente: { escribe: true }, agregarSeguimiento: { escribe: true }, borrarSeguimiento: { escribe: true },
  // solo administrador
  borrarCliente: { admin: true, escribe: true }, asignar: { admin: true, escribe: true },
  guardarUsuario: { admin: true, escribe: true }, reiniciarClave: { admin: true, escribe: true },
  importar: { admin: true, escribe: true }, sincronizarCotizador: { admin: true, escribe: true },
  cambios: { admin: true }, guardarConfig: { admin: true, escribe: true }
};

function enrutar_(p) {
  const nombre = String(p.action || '');
  const def = ACCIONES[nombre];
  if (!def) fallo_('Acción desconocida: ' + nombre);
  preparar_();
  if (nombre === 'ping') return { app: 'crm' };
  let yo = null;
  if (!def.pub) {
    yo = sesion_(p.token);
    if (def.admin && yo.rol !== 'admin') fallo_('Solo el administrador puede hacer esto.', 'PERMISO');
  }
  const fn = globalThis['a_' + nombre];
  if (typeof fn !== 'function') fallo_('Acción no disponible: ' + nombre);
  if (!def.escribe) return fn(p, yo);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) fallo_('El servidor está ocupado. Intenta de nuevo en unos segundos.');
  try {
    const prev = rev_();
    const r = fn(p, yo) || {};
    r.prevRev = prev;
    r.rev = subirRev_();
    return r;
  } finally { lock.releaseLock(); }
}

/* ============================================================
 * Hoja de cálculo y tablas
 * ============================================================ */
let SS_ = null;
function ss_() {
  if (SS_) return SS_;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) { try { SS_ = SpreadsheetApp.openById(id); return SS_; } catch (e) { /* se recrea */ } }
  SS_ = SpreadsheetApp.create(CRM.NOMBRE_HOJA);
  props.setProperty('SHEET_ID', SS_.getId());
  return SS_;
}

function hoja_(nombre) {
  const s = ss_();
  let h = s.getSheetByName(nombre);
  const cols = TABLAS[nombre];
  if (!h) {
    h = s.insertSheet(nombre);
    h.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    h.setFrozenRows(1);
    h.getRange(1, 1, h.getMaxRows(), cols.length).setNumberFormat('@'); // todo como texto
    const sobra = s.getSheetByName('Hoja 1') || s.getSheetByName('Sheet1');
    if (sobra && s.getSheets().length > 1) s.deleteSheet(sobra);
  } else {
    // Agrega columnas nuevas si el esquema creció
    const enc = h.getRange(1, 1, 1, Math.max(h.getLastColumn(), 1)).getValues()[0].map(String);
    cols.forEach(c => { if (enc.indexOf(c) < 0) { enc.push(c); h.getRange(1, enc.length).setValue(c).setFontWeight('bold'); } });
  }
  return h;
}

function leer_(nombre) {
  const h = hoja_(nombre);
  const n = h.getLastRow(), m = h.getLastColumn();
  if (n < 2) return [];
  const v = h.getRange(1, 1, n, m).getValues();
  const enc = v[0].map(String);
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const o = {};
    for (let j = 0; j < enc.length; j++) o[enc[j]] = celda_(v[i][j]);
    o._fila = i + 1;
    out.push(o);
  }
  return out;
}

function celda_(x) {
  if (x instanceof Date) return Utilities.formatDate(x, 'America/Caracas', "yyyy-MM-dd'T'HH:mm:ss");
  return x === null || x === undefined ? '' : String(x);
}

function encabezado_(h) { return h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(String); }
function aFila_(enc, o) { return enc.map(c => (o[c] === undefined || o[c] === null) ? '' : String(o[c])); }

function escribirFila_(nombre, o) {
  const h = hoja_(nombre), enc = encabezado_(h);
  h.getRange(o._fila, 1, 1, enc.length).setNumberFormat('@').setValues([aFila_(enc, o)]);
}

function agregarFilas_(nombre, lista) {
  if (!lista.length) return;
  const h = hoja_(nombre), enc = encabezado_(h);
  const ini = h.getLastRow() + 1;
  const r = h.getRange(ini, 1, lista.length, enc.length);
  r.setNumberFormat('@').setValues(lista.map(o => aFila_(enc, o)));
  lista.forEach((o, i) => o._fila = ini + i);
}

/** Reescribe filas existentes en bloque (más rápido que una por una). */
function escribirFilas_(nombre, lista) {
  if (!lista.length) return;
  const h = hoja_(nombre), enc = encabezado_(h);
  const filas = lista.map(o => o._fila).sort((a, b) => a - b);
  const min = filas[0], max = filas[filas.length - 1];
  if (lista.length > 15 || max - min + 1 <= lista.length * 3) {
    const r = h.getRange(min, 1, max - min + 1, enc.length);
    const v = r.getValues().map(f => f.map(celda_));
    lista.forEach(o => { v[o._fila - min] = aFila_(enc, o); });
    r.setNumberFormat('@').setValues(v);
  } else {
    lista.forEach(o => h.getRange(o._fila, 1, 1, enc.length).setNumberFormat('@').setValues([aFila_(enc, o)]));
  }
}

function borrarFila_(nombre, fila) { hoja_(nombre).deleteRow(fila); }

/** Sin campos vacíos: aligera la descarga (la app trata lo que falta como vacío). */
function compacto_(o) { const c = {}; Object.keys(o).forEach(k => { if (k !== '_fila' && o[k] !== '') c[k] = o[k]; }); return c; }
function limpio_(o) { const c = Object.assign({}, o); delete c._fila; return c; }

/* ============================================================
 * Configuración y versión de datos
 * ============================================================ */
function cfg_(clave) {
  const f = leer_('Config').find(r => r.clave === clave);
  return f ? f.valor : '';
}
function setCfg_(clave, valor) {
  const f = leer_('Config').find(r => r.clave === clave);
  if (f) { f.valor = String(valor); escribirFila_('Config', f); }
  else agregarFilas_('Config', [{ clave: clave, valor: String(valor) }]);
}
function rev_() { return PropertiesService.getScriptProperties().getProperty('REV') || '0'; }
function subirRev_() {
  const r = Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
  PropertiesService.getScriptProperties().setProperty('REV', r);
  return r;
}

/** Crea las hojas y al administrador la primera vez. */
function preparar_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('LISTO') === '2') return;
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (props.getProperty('LISTO') === '2') return;
    Object.keys(TABLAS).forEach(hoja_);
    const us = leer_('Usuarios');
    if (!us.some(u => u.rol === 'admin')) {
      agregarFilas_('Usuarios', [{ id: uid_(), usuario: CRM.ADMIN_USUARIO, nombre: CRM.ADMIN_NOMBRE, rol: 'admin', email: '', telefono: '', activo: 'si', hash: '', salt: '', creado: ahora_(), ultimoAcceso: '' }]);
    }
    if (!cfg_('cotizadorUrl')) setCfg_('cotizadorUrl', CRM.COTIZADOR_URL);
    props.setProperty('LISTO', '2');
    subirRev_();
  } finally { lock.releaseLock(); }
}

/* ============================================================
 * Utilidades
 * ============================================================ */
function uid_() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
function ahora_() { return Utilities.formatDate(new Date(), 'America/Caracas', "yyyy-MM-dd'T'HH:mm:ss"); }
function txt_(v, max) { return String(v === undefined || v === null ? '' : v).trim().slice(0, max || 500); }
function normRif_(r) { return String(r || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function normUsuario_(u) {
  return String(u || '').trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, '.').replace(/[^a-z0-9._-]/g, '');
}
function relleno_(v) { return RELLENOS.indexOf(String(v || '').trim().toLowerCase()) >= 0; }
function fechaOk_(v) { v = txt_(v, 10); return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : ''; }

function hash_(clave, salt) {
  let h = salt + '|' + clave;
  for (let i = 0; i < CRM.HASH_VUELTAS; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  }
  return h;
}
function validarClave_(c) {
  c = String(c || '');
  if (c.length < 6) fallo_('La contraseña debe tener al menos 6 caracteres.');
  if (c.length > 100) fallo_('La contraseña es demasiado larga.');
  return c;
}
function ponerClave_(u, clave) { u.salt = uid_() + uid_(); u.hash = hash_(validarClave_(clave), u.salt); }
function usuarioPublico_(u) {
  return { id: u.id, usuario: u.usuario, nombre: u.nombre, rol: u.rol, email: u.email, telefono: u.telefono, activo: u.activo !== 'no', sinClave: !u.hash, ultimoAcceso: u.ultimoAcceso, creado: u.creado };
}

function registrar_(yo, cliente, accion, detalle) {
  return { fecha: ahora_(), usuarioId: yo ? yo.id : '', usuarioNombre: yo ? yo.nombre : 'Sistema', clienteId: cliente ? cliente.id : '', clienteNombre: cliente ? cliente.nombre : '', accion: accion, detalle: txt_(detalle, 2000) };
}
function guardarCambios_(lista) {
  if (!lista.length) return;
  agregarFilas_('Cambios', lista);
  const h = hoja_('Cambios'), sobra = h.getLastRow() - 1 - CRM.MAX_CAMBIOS;
  if (sobra > 200) h.deleteRows(2, sobra);
}

/* ============================================================
 * Sesiones
 * ============================================================ */
function sesion_(token) {
  if (!token) fallo_('Tu sesión terminó. Vuelve a entrar.', 'SESION');
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('S_' + token);
  if (!raw) fallo_('Tu sesión terminó. Vuelve a entrar.', 'SESION');
  const s = JSON.parse(raw);
  if (s.exp < Date.now()) { props.deleteProperty('S_' + token); fallo_('Tu sesión terminó. Vuelve a entrar.', 'SESION'); }
  const u = leer_('Usuarios').find(x => x.id === s.uid);
  if (!u || u.activo === 'no' || !u.hash || u.salt !== s.salt) { props.deleteProperty('S_' + token); fallo_('Tu sesión terminó. Vuelve a entrar.', 'SESION'); }
  return u;
}

function abrirSesion_(u) {
  const props = PropertiesService.getScriptProperties();
  // Limpia sesiones vencidas
  const todas = props.getProperties(), t0 = Date.now();
  Object.keys(todas).forEach(k => { if (k.indexOf('S_') === 0) { try { if (JSON.parse(todas[k]).exp < t0) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); } } });
  const token = Utilities.getUuid().replace(/-/g, '') + uid_();
  // La sal va en la sesión: si cambia la contraseña, las sesiones viejas dejan de valer
  props.setProperty('S_' + token, JSON.stringify({ uid: u.id, salt: u.salt, exp: t0 + CRM.SESION_DIAS * 86400000 }));
  u.ultimoAcceso = ahora_();
  escribirFila_('Usuarios', u);
  return { token: token, user: usuarioPublico_(u) };
}

function cerrarSesionesDe_(uid) {
  const props = PropertiesService.getScriptProperties(), todas = props.getProperties();
  Object.keys(todas).forEach(k => { if (k.indexOf('S_') === 0) { try { if (JSON.parse(todas[k]).uid === uid) props.deleteProperty(k); } catch (e) {} } });
}

function a_login(p) {
  const nu = normUsuario_(p.usuario);
  const cache = CacheService.getScriptCache(), ck = 'fail_' + nu;
  const fallos = +(cache.get(ck) || 0);
  if (fallos >= CRM.MAX_INTENTOS) fallo_('Demasiados intentos. Espera ' + CRM.BLOQUEO_MIN + ' minutos e intenta de nuevo.');
  const u = leer_('Usuarios').find(x => x.usuario === nu);
  if (u && u.activo !== 'no' && !u.hash) return { crearClave: true, nombre: u.nombre };
  if (!u || u.activo === 'no' || hash_(String(p.password || ''), u.salt) !== u.hash) {
    cache.put(ck, String(fallos + 1), CRM.BLOQUEO_MIN * 60);
    fallo_('Usuario o contraseña incorrectos.');
  }
  cache.remove(ck);
  return abrirSesion_(u);
}

/** Primera entrada: el usuario define su contraseña (solo si aún no tiene una). */
function a_crearClave(p) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const u = leer_('Usuarios').find(x => x.usuario === normUsuario_(p.usuario));
    if (!u || u.activo === 'no') fallo_('Usuario no encontrado.');
    if (u.hash) fallo_('Este usuario ya tiene contraseña. Entra con ella o pide al administrador que la reinicie.');
    ponerClave_(u, p.password);
    escribirFila_('Usuarios', u);
    return abrirSesion_(u);
  } finally { lock.releaseLock(); }
}

function a_logout(p) { PropertiesService.getScriptProperties().deleteProperty('S_' + p.token); return {}; }

function a_cambiarClave(p, yo) {
  if (hash_(String(p.actual || ''), yo.salt) !== yo.hash) fallo_('La contraseña actual no es correcta.');
  ponerClave_(yo, p.nueva);
  escribirFila_('Usuarios', yo);
  cerrarSesionesDe_(yo.id);
  return abrirSesion_(yo);
}

/* ============================================================
 * Lectura de datos (filtrada por vendedor)
 * ============================================================ */
function a_datos(p, yo) {
  const rev = rev_();
  if (p.rev && p.rev === rev) return { same: true, rev: rev };
  const admin = yo.rol === 'admin';
  const usuarios = leer_('Usuarios');
  let clientes = leer_('Clientes');
  if (!admin) clientes = clientes.filter(c => c.vendedorId === yo.id);
  const ids = {}; clientes.forEach(c => ids[c.id] = 1);
  const segs = leer_('Seguimientos').filter(s => ids[s.clienteId]);
  return {
    rev: rev,
    me: usuarioPublico_(yo),
    clientes: clientes.map(compacto_),
    seguimientos: segs.map(limpio_),
    // El vendedor solo necesita nombres; el admin ve todo menos la clave
    usuarios: admin ? usuarios.map(usuarioPublico_) : usuarios.filter(u => u.id === yo.id).map(usuarioPublico_),
    config: admin ? { cotizadorUrl: cfg_('cotizadorUrl'), ultimaSync: cfg_('ultimaSync') } : {}
  };
}

function a_cambios(p) {
  const lista = leer_('Cambios').map(limpio_);
  const lim = Math.min(+p.limite || 1000, 5000);
  return { cambios: lista.slice(-lim).reverse() };
}

/* ============================================================
 * Clientes
 * ============================================================ */
function puedeVer_(yo, c) { return yo.rol === 'admin' || c.vendedorId === yo.id; }

function limpiarCampo_(k, v, usuariosPorId) {
  if (k === 'rif') return txt_(v, 30).toUpperCase();
  if (k === 'email') { v = txt_(v, 120).toLowerCase(); if (v && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) fallo_('El email no es válido: ' + v); return v; }
  if (k === 'estado') { v = txt_(v, 20); return ESTADOS.indexOf(v) >= 0 ? v : 'Activo'; }
  if (k === 'proximaAccion') return fechaOk_(v);
  if (k === 'vendedorId') { v = txt_(v, 40); if (v && !usuariosPorId[v]) fallo_('El vendedor seleccionado no existe.'); return v; }
  if (k === 'notas') return txt_(v, 4000);
  if (k === 'direccion') return txt_(v, 400);
  return txt_(v, 200);
}

function a_guardarCliente(p, yo) {
  const datos = p.cliente || {};
  const admin = yo.rol === 'admin';
  const permitidos = admin ? CAMPOS_ADMIN : CAMPOS_VENDEDOR;
  const usuarios = leer_('Usuarios'), porId = {};
  usuarios.forEach(u => porId[u.id] = u);
  const todos = leer_('Clientes');
  let c = datos.id ? todos.find(x => x.id === datos.id) : null;
  if (datos.id && !c) fallo_('El cliente ya no existe. Actualiza la lista.');
  if (c && !puedeVer_(yo, c)) fallo_('Este cliente no está asignado a ti.', 'PERMISO');
  // Evita sobrescribir cambios de otra persona
  if (c && datos._ver !== undefined && datos._ver !== (c.ver || '')) {
    fallo_('Otra persona modificó este cliente mientras lo editabas (' + (c.modificadoPor || '') + '). Cierra la ficha, vuelve a abrirla y repite tu cambio.', 'CONFLICTO');
  }
  const nuevo = !c;
  if (nuevo) c = { id: uid_(), origen: 'CRM', creado: ahora_(), estado: 'Activo', vendedorId: admin ? '' : yo.id };
  const antes = Object.assign({}, c), difs = [];
  permitidos.forEach(k => {
    if (!(k in datos)) return;
    const v = limpiarCampo_(k, datos[k], porId);
    if (String(c[k] || '') !== v) { difs.push(k); c[k] = v; }
  });
  if (!txt_(c.nombre)) fallo_('El nombre del cliente es obligatorio.');
  if (c.rif) {
    const n = normRif_(c.rif);
    const dup = todos.find(x => x.id !== c.id && normRif_(x.rif) === n);
    if (dup) fallo_('Ya existe un cliente con el RIF/C.I. ' + c.rif + ': ' + dup.nombre + (puedeVer_(yo, dup) ? '' : ' (asignado a otro vendedor)') + '.');
  }
  if (!nuevo && !difs.length) return { cliente: limpio_(c), sinCambios: true };
  tocar_(c, yo);
  if (nuevo) agregarFilas_('Clientes', [c]); else escribirFila_('Clientes', c);
  const det = nuevo ? 'Cliente creado' : difs.map(k => k + ': «' + (k === 'vendedorId' ? nombreU_(porId, antes[k]) : (antes[k] || '')) + '» → «' + (k === 'vendedorId' ? nombreU_(porId, c[k]) : c[k]) + '»').join('; ');
  guardarCambios_([registrar_(yo, c, nuevo ? 'crear' : 'editar', det)]);
  return { cliente: limpio_(c) };
}

function tocar_(c, yo, t) { c.modificado = t || ahora_(); c.modificadoPor = yo.nombre; c.ver = uid_(); }

function nombreU_(porId, id) { return id ? (porId[id] ? porId[id].nombre : '(eliminado)') : 'Sin asignar'; }

function a_borrarCliente(p, yo) {
  const c = leer_('Clientes').find(x => x.id === p.id);
  if (!c) fallo_('El cliente ya no existe.');
  borrarFila_('Clientes', c._fila);
  // Borra sus seguimientos (de abajo hacia arriba para no mover filas)
  leer_('Seguimientos').filter(s => s.clienteId === c.id).sort((a, b) => b._fila - a._fila).forEach(s => borrarFila_('Seguimientos', s._fila));
  guardarCambios_([registrar_(yo, c, 'borrar', 'Cliente eliminado (RIF ' + c.rif + ')')]);
  return { id: c.id };
}

function a_asignar(p, yo) {
  const vid = txt_(p.vendedorId, 40);
  const porId = {}; leer_('Usuarios').forEach(u => porId[u.id] = u);
  if (vid && !porId[vid]) fallo_('El vendedor no existe.');
  const ids = {}; (p.ids || []).forEach(i => ids[i] = 1);
  const cambiados = leer_('Clientes').filter(c => ids[c.id] && c.vendedorId !== vid);
  const t = ahora_(), log = [];
  cambiados.forEach(c => {
    log.push(registrar_(yo, c, 'asignar', nombreU_(porId, c.vendedorId) + ' → ' + nombreU_(porId, vid)));
    c.vendedorId = vid; tocar_(c, yo, t);
  });
  escribirFilas_('Clientes', cambiados);
  guardarCambios_(log);
  return { cambiados: cambiados.length, ids: cambiados.map(c => c.id), vers: cambiados.map(c => c.ver), vendedorId: vid, modificado: t, modificadoPor: yo.nombre };
}

/* ============================================================
 * Seguimientos
 * ============================================================ */
function a_agregarSeguimiento(p, yo) {
  const s = p.seguimiento || {};
  const c = leer_('Clientes').find(x => x.id === s.clienteId);
  if (!c) fallo_('El cliente ya no existe.');
  if (!puedeVer_(yo, c)) fallo_('Este cliente no está asignado a ti.', 'PERMISO');
  const nota = txt_(s.nota, 2000);
  if (!nota) fallo_('Escribe qué pasó en el contacto.');
  const reg = {
    id: uid_(), clienteId: c.id, fecha: fechaOk_(s.fecha) || ahora_().slice(0, 10),
    tipo: txt_(s.tipo, 30) || 'Llamada', nota: nota, proximaFecha: fechaOk_(s.proximaFecha),
    usuarioId: yo.id, usuarioNombre: yo.nombre, creado: ahora_()
  };
  agregarFilas_('Seguimientos', [reg]);
  if (!c.ultimoContacto || reg.fecha >= c.ultimoContacto) c.ultimoContacto = reg.fecha;
  c.proximaAccion = reg.proximaFecha;
  c.proximaNota = reg.proximaFecha ? txt_(s.proximaNota, 200) : '';
  tocar_(c, yo);
  escribirFila_('Clientes', c);
  guardarCambios_([registrar_(yo, c, 'seguimiento', reg.tipo + ': ' + nota)]);
  return { seguimiento: limpio_(reg), cliente: limpio_(c) };
}

function a_borrarSeguimiento(p, yo) {
  const s = leer_('Seguimientos').find(x => x.id === p.id);
  if (!s) fallo_('El seguimiento ya no existe.');
  if (yo.rol !== 'admin' && s.usuarioId !== yo.id) fallo_('Solo puedes borrar tus propios seguimientos.', 'PERMISO');
  const c = leer_('Clientes').find(x => x.id === s.clienteId);
  borrarFila_('Seguimientos', s._fila);
  guardarCambios_([registrar_(yo, c || { id: s.clienteId, nombre: '' }, 'borrar seguimiento', s.tipo + ': ' + s.nota)]);
  return { id: s.id };
}

/* ============================================================
 * Usuarios
 * ============================================================ */
function a_guardarUsuario(p, yo) {
  const d = p.usuario || {};
  const lista = leer_('Usuarios');
  let u = d.id ? lista.find(x => x.id === d.id) : null;
  if (d.id && !u) fallo_('El usuario ya no existe.');
  const nuevo = !u;
  const usuario = normUsuario_(d.usuario || (u && u.usuario));
  if (!usuario || usuario.length < 3) fallo_('El usuario debe tener al menos 3 letras (ej. maria.perez).');
  if (lista.some(x => x.usuario === usuario && (!u || x.id !== u.id))) fallo_('Ya existe el usuario ' + usuario + '.');
  const nombre = txt_(d.nombre, 80);
  if (!nombre) fallo_('Escribe el nombre del usuario.');
  const rol = d.rol === 'admin' ? 'admin' : 'vendedor';
  const activo = d.activo === false || d.activo === 'no' ? 'no' : 'si';
  if (u && u.id === yo.id && (rol !== 'admin' || activo === 'no')) fallo_('No puedes quitarte el rol de administrador ni desactivarte a ti mismo.');
  if (nuevo) u = { id: uid_(), creado: ahora_(), hash: '', salt: '', ultimoAcceso: '' };
  const antes = nuevo ? null : Object.assign({}, u);
  u.usuario = usuario; u.nombre = nombre; u.rol = rol; u.activo = activo;
  u.email = txt_(d.email, 120).toLowerCase(); u.telefono = txt_(d.telefono, 40);
  if (d.clave) ponerClave_(u, d.clave);
  if (nuevo) agregarFilas_('Usuarios', [u]); else escribirFila_('Usuarios', u);
  if (antes && (activo === 'no' || d.clave || antes.rol !== rol)) cerrarSesionesDe_(u.id);
  guardarCambios_([registrar_(yo, null, nuevo ? 'crear usuario' : 'editar usuario', u.nombre + ' (' + u.usuario + ', ' + u.rol + (activo === 'no' ? ', inactivo' : '') + ')')]);
  return { usuario: usuarioPublico_(u) };
}

function a_reiniciarClave(p, yo) {
  const u = leer_('Usuarios').find(x => x.id === p.id);
  if (!u) fallo_('El usuario no existe.');
  if (u.id === yo.id) fallo_('Para cambiar tu propia contraseña usa "Mi cuenta".');
  u.hash = ''; u.salt = '';
  escribirFila_('Usuarios', u);
  cerrarSesionesDe_(u.id);
  guardarCambios_([registrar_(yo, null, 'reiniciar contraseña', u.nombre + ' (' + u.usuario + ')')]);
  return { usuario: usuarioPublico_(u) };
}

function a_guardarConfig(p) {
  const url = txt_(p.cotizadorUrl, 300);
  if (url && !/^https:\/\/script\.google(usercontent)?\.com\//.test(url)) fallo_('La dirección debe empezar por https://script.google.com/');
  setCfg_('cotizadorUrl', url);
  return {};
}

/* ============================================================
 * Importación (Excel) y sincronización con el cotizador
 * ============================================================ */

/**
 * Une una lista de clientes con la base, buscando por RIF/C.I.
 * modo 'agregar'    → solo crea los que no existen y completa campos vacíos.
 * modo 'actualizar' → además reemplaza los datos de los existentes con lo que trae el archivo.
 */
function unir_(filas, modo, yo, origen, conVendedor) {
  const usuarios = leer_('Usuarios');
  const porClave = {};
  usuarios.forEach(u => { porClave[u.usuario] = u; porClave[normUsuario_(u.nombre)] = u; porClave['id:' + u.id] = u; });
  const todos = leer_('Clientes'), porRif = {}, porNombre = {};
  todos.forEach(c => { const n = normRif_(c.rif); if (n) porRif[n] = c; else porNombre[txt_(c.nombre).toUpperCase()] = c; });
  const campos = CAMPOS_ADMIN.filter(k => k !== 'vendedorId' && k !== 'origen');
  const nuevos = [], cambiados = [], errores = [], t = ahora_();
  let iguales = 0, sinVendedor = 0;
  const vistos = {};
  filas.forEach((f, i) => {
    const fila = i + 2;
    const nombre = txt_(f.nombre, 200);
    if (!nombre) { if (Object.keys(f).some(k => txt_(f[k]))) errores.push('Fila ' + fila + ': falta el nombre'); return; }
    const rifN = normRif_(f.rif);
    const clave = rifN || ('N:' + nombre.toUpperCase());
    if (vistos[clave]) {
      errores.push(origen === 'Cotizador'
        ? 'RIF ' + (f.rif || '-') + ' repetido en el cotizador («' + nombre + '»); se tomó el primero'
        : 'Fila ' + fila + ': RIF repetido en el archivo (' + (f.rif || nombre) + ')');
      return;
    }
    vistos[clave] = 1;
    let vid;
    if (conVendedor && txt_(f.vendedor)) {
      const kv = normUsuario_(f.vendedor);
      const u = (kv === 'sin.asignar' || kv === 'ninguno') ? { id: '' } : porClave[kv];
      if (!u) { errores.push('Fila ' + fila + ': no existe el vendedor «' + f.vendedor + '» (se dejó sin cambiar)'); sinVendedor++; }
      else vid = u.id;
    }
    const limpio = {};
    campos.forEach(k => {
      if (!(k in f)) return;
      let v = f[k];
      if (relleno_(v)) v = '';
      try { limpio[k] = limpiarCampo_(k, v, {}); }
      catch (e) { if (k === 'email') { limpio[k] = ''; errores.push('Fila ' + fila + ': email inválido, se dejó vacío'); } else throw e; }
    });
    if ('estado' in f && !txt_(f.estado)) delete limpio.estado;
    const c = rifN ? porRif[rifN] : porNombre[nombre.toUpperCase()];
    if (!c) {
      const n = Object.assign({ id: uid_(), estado: 'Activo', origen: origen, creado: t, vendedorId: vid || '' }, limpio);
      tocar_(n, yo, t);
      nuevos.push(n);
      if (rifN) porRif[rifN] = n; else porNombre[nombre.toUpperCase()] = n;
      return;
    }
    let cambio = false;
    Object.keys(limpio).forEach(k => {
      const v = limpio[k];
      if (!v && modo !== 'actualizar') return;
      if (k === 'estado' && modo !== 'actualizar') return;
      if (modo === 'actualizar' ? (v && String(c[k] || '') !== v) : !txt_(c[k])) { c[k] = v; cambio = true; }
    });
    if (vid !== undefined && c.vendedorId !== vid) { c.vendedorId = vid; cambio = true; }
    if (cambio) { tocar_(c, yo, t); if (cambiados.indexOf(c) < 0 && c._fila) cambiados.push(c); }
    else iguales++;
  });
  escribirFilas_('Clientes', cambiados);
  agregarFilas_('Clientes', nuevos);
  const res = { nuevos: nuevos.length, actualizados: cambiados.length, iguales: iguales, errores: errores.slice(0, 200), totalErrores: errores.length };
  guardarCambios_([registrar_(yo, null, 'importar', origen + ' (' + modo + '): ' + res.nuevos + ' nuevos, ' + res.actualizados + ' actualizados, ' + res.iguales + ' sin cambios, ' + res.totalErrores + ' avisos')]);
  return res;
}

function a_importar(p, yo) {
  const filas = Array.isArray(p.filas) ? p.filas : [];
  if (!filas.length) fallo_('El archivo no trae clientes.');
  if (filas.length > 20000) fallo_('El archivo es demasiado grande (máximo 20.000 filas).');
  return unir_(filas, p.modo === 'actualizar' ? 'actualizar' : 'agregar', yo, 'Excel', true);
}

/** Trae los clientes del servidor del cotizador y los une por RIF. */
function a_sincronizarCotizador(p, yo) {
  const url = txt_(p.url, 300) || cfg_('cotizadorUrl') || CRM.COTIZADOR_URL;
  if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(url)) fallo_('La dirección del cotizador no es válida.');
  const r = UrlFetchApp.fetch(url + (url.indexOf('?') >= 0 ? '&' : '?') + 'action=all', { muteHttpExceptions: true, followRedirects: true });
  let j;
  try { j = JSON.parse(r.getContentText()); } catch (e) { fallo_('El cotizador no respondió con datos (código ' + r.getResponseCode() + '). Revisa la dirección.'); }
  if (!j.ok || !j.data || !Array.isArray(j.data.clientes)) fallo_('El cotizador respondió sin lista de clientes.');
  const filas = j.data.clientes.map(c => {
    const o = {}; CAMPOS_COTIZADOR.forEach(k => o[k] = c[k]); return o;
  });
  const res = unir_(filas, p.modo === 'actualizar' ? 'actualizar' : 'agregar', yo, 'Cotizador', false);
  setCfg_('ultimaSync', ahora_());
  if (p.url) setCfg_('cotizadorUrl', url);
  res.recibidos = filas.length;
  return res;
}

/* ============================================================
 * Funciones para ejecutar a mano desde el editor de Apps Script
 * ============================================================ */

/** Crea la hoja y el usuario administrador. Muestra el enlace de la hoja en el registro. */
function instalar() {
  preparar_();
  Logger.log('Listo. Hoja de datos: ' + ss_().getUrl());
  Logger.log('Administrador: ' + CRM.ADMIN_USUARIO + ' (define su contraseña la primera vez que entra).');
}

/** Si el administrador olvida su contraseña: la borra para que la vuelva a crear al entrar. */
function reiniciarClaveAdministrador() {
  preparar_();
  const u = leer_('Usuarios').find(x => x.usuario === CRM.ADMIN_USUARIO);
  if (!u) { Logger.log('No existe ' + CRM.ADMIN_USUARIO); return; }
  u.hash = ''; u.salt = ''; u.activo = 'si'; u.rol = 'admin';
  escribirFila_('Usuarios', u);
  cerrarSesionesDe_(u.id);
  subirRev_();
  Logger.log('Contraseña de ' + CRM.ADMIN_USUARIO + ' reiniciada. Entra al CRM con ese usuario y crea una nueva.');
}
