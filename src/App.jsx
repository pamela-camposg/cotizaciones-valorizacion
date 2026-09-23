import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import * as api from './api.js';
import { RESIDUOS_TREE, LISTAS, DESTINATARIOS_SEED } from './data.js';

/* ══════════════════════════════════════════════════════════════════════════
   1 · DATOS DERIVADOS
   ══════════════════════════════════════════════════════════════════════════ */
const REGION_BY_COMUNA = Object.fromEntries(LISTAS.COMUNA_REGION);
const COMUNAS = LISTAS.COMUNA_REGION.map((c) => c[0]);
const BASE = import.meta.env.BASE_URL;

/* Cada nodo del árbol viene como "Nombre|COD": se parte por el último "|". */
function splitKey(key) {
  const i = key.lastIndexOf('|');
  return { name: key.slice(0, i), code: key.slice(i + 1) };
}
function nameOf(key) { return key ? splitKey(key).name : ''; }

function todayISO() { return new Date().toISOString().slice(0, 10); }
function fmtNum(n) {
  if (n === '' || n === null || n === undefined || isNaN(n)) return '';
  return new Intl.NumberFormat('es-CL').format(n);
}

/* ══════════════════════════════════════════════════════════════════════════
   2 · MARCA
   ══════════════════════════════════════════════════════════════════════════ */
/* El logo y la "a" son los mismos archivos del Planificador. Si no están
   subidos al repositorio, cae al dibujo de respaldo y la página no se rompe. */
function LogoAmbipar({ className = 'logo' }) {
  const [falla, setFalla] = useState(false);
  if (falla) {
    return (
      <svg viewBox="0 0 260 46" className={className} aria-label="Ambipar">
        <text x="0" y="35" fontFamily="'Barlow Semi Condensed', sans-serif" fontSize="42"
          fontWeight="600" letterSpacing="-1" fill="#031716">ambipar</text>
        <g transform="translate(220 3)" fill="none" stroke="#031716" strokeLinecap="butt">
          <rect x="5" y="5" width="26" height="26" rx="6" transform="rotate(45 18 18)" strokeWidth="3.2" />
          <circle cx="17.4" cy="19.8" r="4.9" strokeWidth="3" />
          <path d="M22.3 13.4 V 26.2" strokeWidth="3" />
        </g>
      </svg>
    );
  }
  return <img src={BASE + 'ambipar-logo.png'} alt="Ambipar" className={className}
    onError={() => setFalla(true)} />;
}

function MarcaDeAgua() {
  const [falla, setFalla] = useState(false);
  return (
    <div aria-hidden="true" className="watermark">
      {falla ? (
        <svg viewBox="0 0 100 100" style={{ padding: '18%' }}>
          <g fill="none" stroke="#CCFF00" strokeLinecap="butt">
            <rect x="22" y="22" width="56" height="56" rx="13" transform="rotate(45 50 50)" strokeWidth="7" />
            <circle cx="48.5" cy="55" r="10.5" strokeWidth="6.5" />
            <path d="M59 41.5 V 68.5" strokeWidth="6.5" />
          </g>
        </svg>
      ) : (
        <img src={BASE + 'ambipar-a-verde.png'} alt="" onError={() => setFalla(true)} />
      )}
    </div>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   3 · PIEZAS DE FORMULARIO
   ══════════════════════════════════════════════════════════════════════════ */
function Field({ label, req, nota, auto, children }) {
  return (
    <div className="field">
      <label>
        {label}{req && <em>*</em>}
        {nota && <i> — {nota}</i>}
        {auto && <span className="auto-badge">automático</span>}
      </label>
      {children}
    </div>
  );
}

function Hint({ children }) { return <div className="hint">{children}</div>; }

function Chips({ options, value, onChange, disabled }) {
  return (
    <div className="chips">
      {options.map((o) => (
        <button key={o} type="button" className="chip" disabled={disabled}
          aria-pressed={value === o} onClick={() => onChange(o)}>{o}</button>
      ))}
    </div>
  );
}

/* Campo de solo lectura, para lo que el sistema completa por su cuenta. */
function Auto({ value, placeholder }) {
  return <input type="text" value={value || ''} readOnly placeholder={placeholder || '—'} />;
}

/* ══════════════════════════════════════════════════════════════════════════
   4 · BUSCADOR DE RESIDUO POR TIPO  [F2]
   ══════════════════════════════════════════════════════════════════════════ */
function BuscadorResiduo({ onClose, onSelect }) {
  const [q, setQ] = useState('');

  const resultados = useMemo(() => {
    const needle = q.trim().toLowerCase();
    if (needle.length < 2) return [];
    const out = [];
    for (const pk of Object.keys(RESIDUOS_TREE)) {
      for (const fk of Object.keys(RESIDUOS_TREE[pk])) {
        for (const mk of Object.keys(RESIDUOS_TREE[pk][fk])) {
          for (const tk of Object.keys(RESIDUOS_TREE[pk][fk][mk])) {
            if (splitKey(tk).name.toLowerCase().includes(needle)) out.push({ pk, fk, mk, tk });
            if (out.length >= 40) return out;
          }
        }
      }
    }
    return out;
  }, [q]);

  return (
    <div className="overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <div className="kicker">Buscar residuo por tipo · F2</div>
          <div className="field" style={{ marginBottom: 0 }}>
            <input autoFocus value={q} onChange={(e) => setQ(e.target.value)}
              placeholder="Escriba, por ejemplo: aceite, cartón, lodo…" />
          </div>
        </div>
        <div className="modal-body">
          {resultados.length === 0 && (
            <div className="empty-note" style={{ border: 0 }}>
              {q.trim().length < 2 ? 'Escriba al menos 2 caracteres.' : 'Sin resultados para esa búsqueda.'}
            </div>
          )}
          {resultados.map((r, i) => (
            <div key={i} className="op" onClick={() => onSelect(r)}>
              <b>{nameOf(r.tk)}</b>
              <span>{nameOf(r.pk)} · {nameOf(r.fk)} · {nameOf(r.mk)}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   5 · PANTALLA DE ACCESO
   ══════════════════════════════════════════════════════════════════════════ */
function LoginScreen({ onLogin }) {
  const [usuario, setUsuario] = useState('');
  const [pin, setPin] = useState('');
  const [error, setError] = useState('');
  const [cargando, setCargando] = useState(false);

  async function submit(e) {
    e.preventDefault();
    if (!usuario) { setError('Seleccione su usuario.'); return; }
    if (!pin.trim()) { setError('Ingrese su PIN.'); return; }
    setError(''); setCargando(true);
    try {
      const s = await api.login(usuario.trim(), pin.trim());
      onLogin({ usuario: s.usuario, pin: pin.trim(), nombre: s.nombre });
    } catch (err) {
      setError(err.message);
      setCargando(false);
    }
  }

  return (
    <div className="login-wrap">
      <MarcaDeAgua />
      <form className="login-card" onSubmit={submit}>
        <LogoAmbipar className="login-logo" />
        <span className="lime-rule" />
        <h1 className="premium-heading">Cotizaciones de Valorización</h1>
        <p className="sub">Registro de cotizaciones de valorización y eliminación de residuos.</p>

        <Field label="Usuario" req>
          <select value={usuario} onChange={(e) => setUsuario(e.target.value)} disabled={cargando}>
            <option value="">Seleccione su usuario…</option>
            {LISTAS.RESPONSABLE.map((r) => <option key={r} value={r}>{r}</option>)}
          </select>
        </Field>

        <Field label="PIN de acceso" req>
          <input type="password" value={pin} onChange={(e) => setPin(e.target.value)}
            placeholder="••••" disabled={cargando} autoComplete="off" />
        </Field>

        {error && <div className="login-error">{error}</div>}

        <button type="submit" className="btn ancho" disabled={cargando}>
          {cargando ? 'Verificando…' : 'Ingresar'}
        </button>

        <p className="login-pie">
          El usuario y el PIN se validan contra la hoja USUARIOS de la base compartida.
          Para revocar un acceso, cambie ACTIVO a NO en esa hoja.
        </p>
      </form>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   6 · FORMULARIO DE COTIZACIÓN
   ══════════════════════════════════════════════════════════════════════════ */
const VACIO_LOG = { tipoServicio: 'En planta', canalCotizacion: 'Correo', comunaOrigen: '', origen: '' };
const VACIO_EST = {
  nombreFantasia: '', comunaEstablecimiento: '', sistemaDeclaracion: '', codigoEstablecimiento: '',
  razonSocial: '', rut: '', nombreEstablecimiento: '', regionEstablecimiento: '', esNuevo: false,
};
const VACIO_RES = { pk: '', fk: '', mk: '', tk: '', ek: '', fok: '', otroNombre: '' };
const VACIO_COM = { tipoNegocio: 'Venta', unidad: '', valor: '' };
const VACIO_ADJ = { documentos: '', comentarios: '' };

function QuoteForm({ sesion, destinatarios, onSaveQuote, onAddDestinatario, notify }) {
  const [log, setLog] = useState(VACIO_LOG);
  const [est, setEst] = useState(VACIO_EST);
  const [res, setRes] = useState(VACIO_RES);
  const [com, setCom] = useState(VACIO_COM);
  const [adj, setAdj] = useState(VACIO_ADJ);
  const [buscando, setBuscando] = useState(false);
  const [guardando, setGuardando] = useState(false);
  const [query, setQuery] = useState('');

  /* ---------- establecimiento ---------- */
  const nombres = useMemo(
    () => [...new Set(destinatarios.map((d) => d.NOMBRE_FANTASIA).filter(Boolean))].sort(),
    [destinatarios]
  );
  const sugerencias = query.trim()
    ? nombres.filter((n) => n.toLowerCase().includes(query.trim().toLowerCase())).slice(0, 8)
    : [];

  const comunasDe = useMemo(() => {
    if (!est.nombreFantasia) return [];
    return [...new Set(destinatarios
      .filter((d) => d.NOMBRE_FANTASIA === est.nombreFantasia)
      .map((d) => d.COMUNA_ESTABLECIMIENTO).filter(Boolean))];
  }, [destinatarios, est.nombreFantasia]);

  const sistemasDe = useMemo(() => {
    if (!est.nombreFantasia || !est.comunaEstablecimiento) return [];
    return [...new Set(destinatarios
      .filter((d) => d.NOMBRE_FANTASIA === est.nombreFantasia
        && d.COMUNA_ESTABLECIMIENTO === est.comunaEstablecimiento)
      .map((d) => d.SISTEMA_DECLARACION).filter(Boolean))];
  }, [destinatarios, est.nombreFantasia, est.comunaEstablecimiento]);

  function elegirFantasia(nombre) {
    const esNuevo = !nombres.includes(nombre);
    setEst({ ...VACIO_EST, nombreFantasia: nombre, esNuevo });
    setQuery('');
    /* Si solo hay una comuna, se elige sola: un clic menos. */
    if (!esNuevo) {
      const cs = [...new Set(destinatarios.filter((d) => d.NOMBRE_FANTASIA === nombre)
        .map((d) => d.COMUNA_ESTABLECIMIENTO).filter(Boolean))];
      if (cs.length === 1) setTimeout(() => elegirComuna(cs[0], nombre), 0);
    }
  }

  function elegirComuna(comuna, fantasia) {
    const nf = fantasia || est.nombreFantasia;
    setEst((p) => ({ ...p, nombreFantasia: nf, comunaEstablecimiento: comuna, sistemaDeclaracion: '',
      codigoEstablecimiento: '', razonSocial: '', rut: '', nombreEstablecimiento: '', regionEstablecimiento: '' }));
    const ss = [...new Set(destinatarios
      .filter((d) => d.NOMBRE_FANTASIA === nf && d.COMUNA_ESTABLECIMIENTO === comuna)
      .map((d) => d.SISTEMA_DECLARACION).filter(Boolean))];
    if (ss.length === 1) setTimeout(() => elegirSistema(ss[0], nf, comuna), 0);
  }

  function elegirSistema(sistema, fantasia, comuna) {
    const nf = fantasia || est.nombreFantasia;
    const cm = comuna || est.comunaEstablecimiento;
    const fila = destinatarios.find((d) => d.NOMBRE_FANTASIA === nf
      && d.COMUNA_ESTABLECIMIENTO === cm && d.SISTEMA_DECLARACION === sistema);
    setEst((p) => ({
      ...p, nombreFantasia: nf, comunaEstablecimiento: cm, sistemaDeclaracion: sistema,
      codigoEstablecimiento: fila ? fila.CODIGO_ESTABLECIMIENTO : '',
      razonSocial: fila ? fila.RAZON_SOCIAL : '',
      rut: fila ? fila.RUT : '',
      nombreEstablecimiento: fila ? fila.NOMBRE_ESTABLECIMIENTO : '',
      regionEstablecimiento: fila ? fila.REGION_ESTABLECIMIENTO : '',
    }));
  }

  /* ---------- cascada del residuo ---------- */
  const T = RESIDUOS_TREE;
  const opP = Object.keys(T);
  const opF = res.pk ? Object.keys(T[res.pk]) : [];
  const opM = res.fk ? Object.keys(T[res.pk][res.fk]) : [];
  const opT = res.mk ? Object.keys(T[res.pk][res.fk][res.mk]) : [];
  const opE = res.tk ? Object.keys(T[res.pk][res.fk][res.mk][res.tk]) : [];
  const opFo = res.ek ? T[res.pk][res.fk][res.mk][res.tk][res.ek] : [];

  const codRes = useMemo(() => {
    const { pk, fk, mk, tk, ek, fok } = res;
    if (!(pk && fk && mk && tk && ek && fok)) return '';
    return [pk, fk, mk, tk, ek, fok].map((k) => splitKey(k).code).join('-');
  }, [res]);

  /* ---------- guardar ---------- */
  function limpiarParcial() {
    /* Los apartados 1 y 2 quedan cargados: lo habitual es registrar varias
       cotizaciones seguidas del mismo establecimiento. */
    setRes(VACIO_RES); setCom(VACIO_COM); setAdj(VACIO_ADJ);
  }
  function limpiarTodo() {
    setLog(VACIO_LOG); setEst(VACIO_EST); setQuery('');
    setRes(VACIO_RES); setCom(VACIO_COM); setAdj(VACIO_ADJ);
  }

  function validar() {
    if (!est.nombreFantasia) return 'Falta el establecimiento (Apartado 2).';
    if (est.esNuevo && (!est.razonSocial || !est.rut || !est.comunaEstablecimiento))
      return 'Del destinatario nuevo faltan razón social, RUT o comuna.';
    if (!est.esNuevo && !est.comunaEstablecimiento) return 'Falta la comuna del establecimiento.';
    if (!codRes) return 'Falta completar la caracterización del residuo (Apartado 3).';
    if (log.tipoServicio === 'Con retiro' && !log.comunaOrigen)
      return 'Un servicio con retiro necesita la comuna de origen.';
    if (!com.unidad) return 'Falta la unidad de la condición comercial (Apartado 4).';
    if (com.valor === '' || isNaN(Number(com.valor))) return 'Falta el valor de la condición comercial.';
    return null;
  }

  const handleSave = useCallback(async () => {
    const err = validar();
    if (err) { notify(err, true); return; }
    setGuardando(true);
    try {
      if (est.esNuevo) {
        await onAddDestinatario({
          NOMBRE_FANTASIA: est.nombreFantasia,
          SISTEMA_DECLARACION: est.sistemaDeclaracion || 'N/A',
          CODIGO_ESTABLECIMIENTO: est.codigoEstablecimiento || '(pendiente)',
          RAZON_SOCIAL: est.razonSocial,
          RUT: est.rut,
          NOMBRE_ESTABLECIMIENTO: est.nombreEstablecimiento || est.razonSocial,
          COMUNA_ESTABLECIMIENTO: est.comunaEstablecimiento,
          REGION_ESTABLECIMIENTO: REGION_BY_COMUNA[est.comunaEstablecimiento] || '',
        });
      }
      const registro = {
        TIPO_SERVICIO: log.tipoServicio,
        CANAL_COTIZACION: log.canalCotizacion,
        ORIGEN: log.origen,
        COMUNA_ORIGEN: log.comunaOrigen,
        REGION_ORIGEN: REGION_BY_COMUNA[log.comunaOrigen] || '',
        NOMBRE_FANTASIA: est.nombreFantasia,
        COMUNA_ESTABLECIMIENTO: est.comunaEstablecimiento,
        REGION_ESTABLECIMIENTO: est.esNuevo
          ? (REGION_BY_COMUNA[est.comunaEstablecimiento] || '') : est.regionEstablecimiento,
        SISTEMA_DECLARACION: est.sistemaDeclaracion,
        CODIGO_ESTABLECIMIENTO: est.codigoEstablecimiento,
        RAZON_SOCIAL: est.razonSocial,
        RUT: est.rut,
        NOMBRE_ESTABLECIMIENTO: est.nombreEstablecimiento,
        COD_RES: codRes,
        PELIGROSIDAD: nameOf(res.pk),
        FAMILIA: nameOf(res.fk),
        MATERIAL: nameOf(res.mk),
        TIPO: nameOf(res.tk),
        ESTADO: nameOf(res.ek),
        FORMATO: nameOf(res.fok),
        OTRO_NOMBRE: res.otroNombre,
        TIPO_NEGOCIO: com.tipoNegocio,
        VALOR: Number(com.valor),
        UNIDAD: com.unidad,
        DOCUMENTOS: adj.documentos,
        COMENTARIOS: adj.comentarios,
      };
      const id = await onSaveQuote(registro);
      notify(`Cotización ${id} guardada en la base compartida.`);
      limpiarParcial();
    } catch (e) {
      /* No se limpia nada: así no se pierde lo escrito y se puede reintentar. */
      notify('No se pudo guardar: ' + e.message, true);
    } finally {
      setGuardando(false);
    }
  }, [log, est, res, com, adj, codRes, onSaveQuote, onAddDestinatario, notify]);

  /* ---------- atajos ---------- */
  useEffect(() => {
    const h = (e) => {
      if (e.key === 'F2') { e.preventDefault(); setBuscando(true); }
      if (e.key === 'Escape') setBuscando(false);
      if (e.ctrlKey && e.key === 'Enter') { e.preventDefault(); handleSave(); }
    };
    window.addEventListener('keydown', h);
    return () => window.removeEventListener('keydown', h);
  }, [handleSave]);

  const conRetiro = log.tipoServicio === 'Con retiro';

  return (
    <>
      {buscando && (
        <BuscadorResiduo onClose={() => setBuscando(false)}
          onSelect={(r) => { setRes({ ...VACIO_RES, ...r }); setBuscando(false); }} />
      )}

      <div className="page-head">
        <span className="lime-rule" />
        <h1 className="premium-heading">Nueva cotización</h1>
        <p className="bajada">
          Los campos marcados con <b>*</b> son obligatorios. El código de residuo, el ID
          y la fecha se generan solos. <b>F2</b> busca un residuo por su tipo;
          <b> Ctrl+Enter</b> guarda.
        </p>
      </div>

      {/* ══════════ 01 · LOGÍSTICA ══════════ */}
      <section className="card">
        <div className="card-head">
          <span className="fmtNum">01</span>
          <h2>Logística y control general</h2>
          <span className="tag">{todayISO()} · {sesion.usuario}</span>
        </div>
        <p className="card-note">
          La fecha, el correlativo y el responsable los asigna el sistema al guardar.
        </p>

        <div className="form-section">Cómo se cotiza</div>
        <div className="row2">
          <Field label="Tipo de servicio" req>
            <Chips options={LISTAS.TIPO_SERVICIO} value={log.tipoServicio}
              onChange={(v) => setLog((p) => ({ ...p, tipoServicio: v,
                comunaOrigen: v === 'En planta' ? '' : p.comunaOrigen,
                origen: v === 'En planta' ? '' : p.origen }))} />
          </Field>
          <Field label="Canal de cotización" req>
            <Chips options={LISTAS.CANAL_COTIZACION} value={log.canalCotizacion}
              onChange={(v) => setLog((p) => ({ ...p, canalCotizacion: v }))} />
          </Field>
        </div>

        <div className="form-section">Origen del residuo</div>
        {!conRetiro && (
          <div className="aviso">
            En un servicio <b>en planta</b> el cliente lleva el residuo, así que no
            corresponde indicar origen. Cambie a <b>Con retiro</b> si hay que ir a buscarlo.
          </div>
        )}
        {conRetiro && (
          <>
            <div className="row2">
              <Field label="Comuna de origen" req>
                <select value={log.comunaOrigen}
                  onChange={(e) => setLog((p) => ({ ...p, comunaOrigen: e.target.value }))}>
                  <option value="">Seleccione comuna…</option>
                  {COMUNAS.map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
              </Field>
              <Field label="Región de origen" auto>
                <Auto value={REGION_BY_COMUNA[log.comunaOrigen]} placeholder="Según la comuna" />
              </Field>
            </div>
            <Field label="Dirección o referencia" nota="opcional">
              <input value={log.origen} onChange={(e) => setLog((p) => ({ ...p, origen: e.target.value }))}
                placeholder="Ej: Av. Industrial 1234, bodega 5" />
            </Field>
          </>
        )}
      </section>

      {/* ══════════ 02 · ESTABLECIMIENTO ══════════ */}
      <section className="card">
        <div className="card-head">
          <span className="fmtNum">02</span>
          <h2>Establecimiento de destino</h2>
          {est.nombreFantasia && (
            <span className="tag">{est.esNuevo ? 'NUEVO' : 'EN LA BASE'}</span>
          )}
        </div>
        <p className="card-note">
          Escriba el nombre de fantasía. Si ya está en la base, el resto se completa solo.
        </p>

        <div className="form-section">Identificación</div>
        <Field label="Nombre de fantasía" req>
          <div className="ac-wrap">
            <input value={est.nombreFantasia || query}
              onChange={(e) => { setQuery(e.target.value); setEst(VACIO_EST); }}
              placeholder="Escriba para buscar…" autoComplete="off" />
            {sugerencias.length > 0 && (
              <div className="ac-list">
                {sugerencias.map((n) => (
                  <div key={n} onClick={() => elegirFantasia(n)}>{n}</div>
                ))}
              </div>
            )}
          </div>
          {query.trim() && sugerencias.length === 0 && (
            <div className="ac-new">
              Sin coincidencias.{' '}
              <button type="button" onClick={() => elegirFantasia(query.trim())}>
                Registrar «{query.trim()}» como destinatario nuevo
              </button>
            </div>
          )}
        </Field>

        {est.nombreFantasia && !est.esNuevo && (
          <>
            <div className="row2">
              <Field label="Comuna del establecimiento" req>
                <select value={est.comunaEstablecimiento} onChange={(e) => elegirComuna(e.target.value)}>
                  <option value="">Seleccione…</option>
                  {comunasDe.map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
              </Field>
              <Field label="Sistema de declaración" req>
                <select value={est.sistemaDeclaracion} disabled={!est.comunaEstablecimiento}
                  onChange={(e) => elegirSistema(e.target.value)}>
                  <option value="">Seleccione…</option>
                  {sistemasDe.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
              </Field>
            </div>

            <div className="form-section">Datos que trae la base</div>
            <div className="row2">
              <Field label="Razón social" auto><Auto value={est.razonSocial} /></Field>
              <Field label="RUT" auto><Auto value={est.rut} /></Field>
            </div>
            <div className="row3">
              <Field label="Código establecimiento" auto><Auto value={est.codigoEstablecimiento} /></Field>
              <Field label="Nombre establecimiento" auto><Auto value={est.nombreEstablecimiento} /></Field>
              <Field label="Región" auto><Auto value={est.regionEstablecimiento} /></Field>
            </div>
          </>
        )}

        {est.esNuevo && (
          <>
            <div className="aviso nuevo">
              <b>Destinatario nuevo.</b> Al guardar la cotización se agrega a la base
              compartida y queda disponible para todo el equipo.
            </div>
            <div className="form-section">Datos del nuevo destinatario</div>
            <div className="row2">
              <Field label="Razón social" req>
                <input value={est.razonSocial}
                  onChange={(e) => setEst((p) => ({ ...p, razonSocial: e.target.value }))} />
              </Field>
              <Field label="RUT" req>
                <input value={est.rut} placeholder="12345678-9"
                  onChange={(e) => setEst((p) => ({ ...p, rut: e.target.value }))} />
              </Field>
            </div>
            <div className="row2">
              <Field label="Comuna del establecimiento" req>
                <select value={est.comunaEstablecimiento}
                  onChange={(e) => setEst((p) => ({ ...p, comunaEstablecimiento: e.target.value }))}>
                  <option value="">Seleccione…</option>
                  {COMUNAS.map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
                {est.comunaEstablecimiento && (
                  <Hint>Región: {REGION_BY_COMUNA[est.comunaEstablecimiento]}</Hint>
                )}
              </Field>
              <Field label="Sistema de declaración" req>
                <select value={est.sistemaDeclaracion}
                  onChange={(e) => setEst((p) => ({ ...p, sistemaDeclaracion: e.target.value }))}>
                  <option value="">Seleccione…</option>
                  <option value="SINADER">SINADER</option>
                  <option value="N/A">N/A</option>
                </select>
              </Field>
            </div>
            <div className="row2">
              <Field label="Código establecimiento" nota="si aún no se conoce, queda pendiente">
                <input value={est.codigoEstablecimiento}
                  onChange={(e) => setEst((p) => ({ ...p, codigoEstablecimiento: e.target.value }))} />
              </Field>
              <Field label="Nombre del establecimiento" nota="opcional">
                <input value={est.nombreEstablecimiento}
                  onChange={(e) => setEst((p) => ({ ...p, nombreEstablecimiento: e.target.value }))} />
              </Field>
            </div>
          </>
        )}
      </section>

      {/* ══════════ 03 · RESIDUO ══════════ */}
      <section className="card">
        <div className="card-head">
          <span className="fmtNum">03</span>
          <h2>Caracterización del residuo</h2>
          <button className="link-btn" style={{ marginLeft: 'auto' }}
            onClick={() => setBuscando(true)}>F2 · Buscar por tipo</button>
        </div>
        <p className="card-note">
          Cada nivel filtra al siguiente. El código se arma con los seis segmentos.
        </p>

        <div className="form-section">Clasificación</div>
        <div className="row2">
          <Field label="Peligrosidad" req>
            <select value={res.pk} onChange={(e) => setRes({ ...VACIO_RES, pk: e.target.value })}>
              <option value="">Seleccione…</option>
              {opP.map((k) => <option key={k} value={k}>{nameOf(k)}</option>)}
            </select>
          </Field>
          <Field label="Familia" req>
            <select value={res.fk} disabled={!res.pk}
              onChange={(e) => setRes((p) => ({ ...VACIO_RES, pk: p.pk, fk: e.target.value }))}>
              <option value="">Seleccione…</option>
              {opF.map((k) => <option key={k} value={k}>{nameOf(k)}</option>)}
            </select>
          </Field>
        </div>
        <div className="row2">
          <Field label="Material" req>
            <select value={res.mk} disabled={!res.fk}
              onChange={(e) => setRes((p) => ({ ...p, mk: e.target.value, tk: '', ek: '', fok: '' }))}>
              <option value="">Seleccione…</option>
              {opM.map((k) => <option key={k} value={k}>{nameOf(k)}</option>)}
            </select>
          </Field>
          <Field label="Tipo" req>
            <select value={res.tk} disabled={!res.mk}
              onChange={(e) => setRes((p) => ({ ...p, tk: e.target.value, ek: '', fok: '' }))}>
              <option value="">Seleccione…</option>
              {opT.map((k) => <option key={k} value={k}>{nameOf(k)}</option>)}
            </select>
          </Field>
        </div>

        <div className="form-section">Cómo viene</div>
        <div className="row2">
          <Field label="Estado" req>
            <select value={res.ek} disabled={!res.tk}
              onChange={(e) => setRes((p) => ({ ...p, ek: e.target.value, fok: '' }))}>
              <option value="">Seleccione…</option>
              {opE.map((k) => <option key={k} value={k}>{nameOf(k)}</option>)}
            </select>
          </Field>
          <Field label="Formato" req>
            <select value={res.fok} disabled={!res.ek}
              onChange={(e) => setRes((p) => ({ ...p, fok: e.target.value }))}>
              <option value="">Seleccione…</option>
              {opFo.map((k) => <option key={k} value={k}>{nameOf(k)}</option>)}
            </select>
          </Field>
        </div>

        <div className="codres">
          <span>Código de residuo</span>
          <b className={codRes ? '' : 'vacio'}>{codRes || '— — — — — —'}</b>
        </div>

        <Field label="Otro nombre" nota="nombre comercial, opcional">
          <input maxLength={40} value={res.otroNombre}
            onChange={(e) => setRes((p) => ({ ...p, otroNombre: e.target.value }))}
            placeholder="Como lo llama el cliente" />
        </Field>
      </section>

      {/* ══════════ 04 · COMERCIAL ══════════ */}
      <section className="card">
        <div className="card-head">
          <span className="fmtNum">04</span>
          <h2>Condición comercial</h2>
        </div>
        <p className="card-note">
          <b>Venta</b>: el cliente paga por el servicio. <b>Compra/Pago</b>: Ambipar paga
          por el material.
        </p>

        <div className="form-section">Negocio</div>
        <Field label="Tipo de negocio" req>
          <Chips options={LISTAS.TIPO_NEGOCIO} value={com.tipoNegocio}
            onChange={(v) => setCom((p) => ({ ...p, tipoNegocio: v }))} />
        </Field>
        <div className="row2">
          <Field label="Unidad" req>
            <select value={com.unidad} onChange={(e) => setCom((p) => ({ ...p, unidad: e.target.value }))}>
              <option value="">Moneda y medida…</option>
              {LISTAS.UNIDAD.map((u) => <option key={u} value={u}>{u}</option>)}
            </select>
          </Field>
          <Field label="Valor" req>
            <input type="number" min="0" step="any" value={com.valor}
              onChange={(e) => setCom((p) => ({ ...p, valor: e.target.value }))} placeholder="0" />
            {com.valor !== '' && com.unidad && !isNaN(Number(com.valor)) && (
              <Hint>{fmtNum(Number(com.valor))} {com.unidad}</Hint>
            )}
          </Field>
        </div>
      </section>

      {/* ══════════ 05 · RESPALDO ══════════ */}
      <section className="card">
        <div className="card-head">
          <span className="fmtNum">05</span>
          <h2>Respaldo y comentarios</h2>
        </div>

        <div className="form-section">Antecedentes</div>
        <Field label="Referencia del documento" nota="nombre o enlace, opcional">
          <input value={adj.documentos}
            onChange={(e) => setAdj((p) => ({ ...p, documentos: e.target.value }))}
            placeholder="Ej: cotizacion_cliente.pdf" />
        </Field>
        <Field label="Comentarios" nota="máx. 100 caracteres">
          <textarea rows={2} maxLength={100} value={adj.comentarios}
            onChange={(e) => setAdj((p) => ({ ...p, comentarios: e.target.value }))} />
          <Hint>{adj.comentarios.length}/100</Hint>
        </Field>
      </section>

      <div className="toolbar" style={{ justifyContent: 'flex-end', marginBottom: '1rem' }}>
        <button className="btn secondary" onClick={limpiarTodo} disabled={guardando}>
          Limpiar todo
        </button>
        <button className="btn" onClick={handleSave} disabled={guardando}>
          {guardando ? 'Guardando…' : 'Guardar cotización'}
        </button>
      </div>
    </>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   7 · HISTORIAL
   ══════════════════════════════════════════════════════════════════════════ */
function HistorialPanel({ cotizaciones, onImport }) {
  const [fMaterial, setFMaterial] = useState('');
  const [fRegion, setFRegion] = useState('');
  const [fNegocio, setFNegocio] = useState('');
  const archivo = useRef(null);

  const materiales = useMemo(
    () => [...new Set(cotizaciones.map((c) => c.MATERIAL).filter(Boolean))].sort(), [cotizaciones]);
  const regiones = useMemo(
    () => [...new Set(cotizaciones.map((c) => c.REGION_ESTABLECIMIENTO).filter(Boolean))].sort(), [cotizaciones]);

  const filtradas = useMemo(() => cotizaciones
    .filter((c) => (!fMaterial || c.MATERIAL === fMaterial)
      && (!fRegion || c.REGION_ESTABLECIMIENTO === fRegion)
      && (!fNegocio || c.TIPO_NEGOCIO === fNegocio))
    .sort((a, b) => String(b.FECHA || '').localeCompare(String(a.FECHA || ''))),
    [cotizaciones, fMaterial, fRegion, fNegocio]);

  const valores = filtradas.map((c) => Number(c.VALOR)).filter((v) => !isNaN(v) && v !== 0);
  const promedio = valores.length ? valores.reduce((s, v) => s + v, 0) / valores.length : 0;
  const establecimientos = new Set(filtradas.map((c) => c.NOMBRE_FANTASIA)).size;

  function leerArchivo(e) {
    const f = e.target.files[0];
    if (!f) return;
    const lector = new FileReader();
    lector.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array', cellDates: true });
        if (!wb.SheetNames.includes('COTIZACIONES')) {
          alert('El archivo no tiene una hoja llamada COTIZACIONES.');
          return;
        }
        const filas = XLSX.utils.sheet_to_json(wb.Sheets['COTIZACIONES'], { defval: '' })
          .filter((r) => r.ID)
          .map((r) => ({ ...r, FECHA: r.FECHA instanceof Date
            ? r.FECHA.toISOString().slice(0, 10) : String(r.FECHA || '') }));
        onImport(filas);
      } catch (err) {
        alert('No se pudo leer el archivo: ' + err.message);
      }
    };
    lector.readAsArrayBuffer(f);
    e.target.value = '';
  }

  return (
    <>
      <div className="page-head">
        <span className="lime-rule" />
        <h1 className="premium-heading">Historial de cotizaciones</h1>
        <p className="bajada">
          Las cotizaciones registradas en la base compartida. El histórico anterior
          a esta herramienta vive en el Excel y se puede cargar aquí para consultarlo.
        </p>
      </div>

      <section className="card">
        <div className="card-head">
          <span className="fmtNum">01</span>
          <h2>Consulta</h2>
          <span className="tag">{filtradas.length} REGISTROS</span>
        </div>

        <div className="toolbar">
          <div className="field">
            <label>Material</label>
            <select value={fMaterial} onChange={(e) => setFMaterial(e.target.value)}>
              <option value="">Todos</option>
              {materiales.map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
          </div>
          <div className="field">
            <label>Región</label>
            <select value={fRegion} onChange={(e) => setFRegion(e.target.value)}>
              <option value="">Todas</option>
              {regiones.map((r) => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
          <div className="field">
            <label>Negocio</label>
            <select value={fNegocio} onChange={(e) => setFNegocio(e.target.value)}>
              <option value="">Ambos</option>
              {LISTAS.TIPO_NEGOCIO.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </div>
          <div className="right">
            <button className="btn secondary" onClick={() => archivo.current.click()}>
              Cargar histórico (.xlsx)
            </button>
            <input ref={archivo} type="file" accept=".xlsx,.xls"
              style={{ display: 'none' }} onChange={leerArchivo} />
          </div>
        </div>

        <div className="kpis">
          <div className="kpi"><b>{filtradas.length}</b><span>Cotizaciones</span></div>
          <div className="kpi"><b>{establecimientos}</b><span>Establecimientos</span></div>
          <div className="kpi"><b>{fmtNum(Math.round(promedio))}</b><span>Valor promedio</span></div>
          <div className="kpi"><b>{materiales.length}</b><span>Materiales</span></div>
        </div>

        <div className="tabla-box">
          <div className="tabla-scroll">
            <table className="datos">
              <thead>
                <tr>
                  {['Fecha', 'ID', 'Establecimiento', 'Región', 'Material', 'Tipo',
                    'Negocio', 'Valor', 'Unidad', 'Responsable'].map((h) => <th key={h}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {filtradas.slice(0, 400).map((c, i) => (
                  <tr key={c.ID || i}>
                    <td>{c.FECHA}</td>
                    <td className="destacado">{c.ID}</td>
                    <td>{c.NOMBRE_FANTASIA || c.NOMBRE_ESTABLECIMIENTO}</td>
                    <td>{c.REGION_ESTABLECIMIENTO}</td>
                    <td>{c.MATERIAL}</td>
                    <td>{c.TIPO}</td>
                    <td>{c.TIPO_NEGOCIO}</td>
                    <td className="num">{fmtNum(c.VALOR)}</td>
                    <td>{c.UNIDAD}</td>
                    <td>{c.RESPONSABLE}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {filtradas.length === 0 && (
              <div className="empty-note" style={{ border: 0, margin: '1rem' }}>
                {cotizaciones.length === 0
                  ? 'Todavía no hay cotizaciones registradas. Use «Cargar histórico» para revisar las anteriores.'
                  : 'Ningún registro coincide con estos filtros.'}
              </div>
            )}
          </div>
        </div>
        {filtradas.length > 400 && (
          <p className="card-note">Se muestran las 400 más recientes de {filtradas.length}.</p>
        )}
      </section>
    </>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   8 · APLICACIÓN
   ══════════════════════════════════════════════════════════════════════════ */
export default function App() {
  /* La sesión vive solo en memoria: al recargar hay que ingresar el PIN. */
  const [sesion, setSesion] = useState(null);
  const [vista, setVista] = useState('form');
  const [destinatarios, setDestinatarios] = useState([]);
  const [cotizaciones, setCotizaciones] = useState([]);
  const [cargado, setCargado] = useState(false);
  const [errorCarga, setErrorCarga] = useState('');
  const [toast, setToast] = useState({ txt: '', malo: false });
  const toastTimer = useRef(null);

  const notify = useCallback((txt, malo) => {
    setToast({ txt, malo: !!malo });
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToast({ txt: '', malo: false }), 4200);
  }, []);

  useEffect(() => {
    if (!sesion) return;
    let cancelado = false;
    (async () => {
      setCargado(false); setErrorCarga('');
      try {
        const [dest, cots] = await Promise.all([
          api.getDestinatarios(sesion),
          api.getCotizaciones(sesion),
        ]);
        if (cancelado) return;
        setDestinatarios(dest);
        setCotizaciones(cots);
      } catch (err) {
        if (cancelado) return;
        setErrorCarga(err.message);
        /* Respaldo local: permite consultar aunque no se pueda guardar. */
        setDestinatarios(DESTINATARIOS_SEED.rows.map((r) => ({
          NOMBRE_FANTASIA: r[0], SISTEMA_DECLARACION: r[1], CODIGO_ESTABLECIMIENTO: r[2],
          RAZON_SOCIAL: r[3], RUT: r[4], NOMBRE_ESTABLECIMIENTO: r[5],
          COMUNA_ESTABLECIMIENTO: r[6], REGION_ESTABLECIMIENTO: r[7],
        })));
      } finally {
        if (!cancelado) setCargado(true);
      }
    })();
    return () => { cancelado = true; };
  }, [sesion]);

  async function handleAddDestinatario(fila) {
    const yaEstaba = await api.agregarDestinatario(sesion, fila);
    if (!yaEstaba) setDestinatarios((p) => [...p, fila]);
  }

  async function handleSaveQuote(registro) {
    /* El ID, la fecha y el responsable los pone el servidor, para que el
       correlativo no se rompa si dos personas guardan a la vez. */
    const id = await api.guardarCotizacion(sesion, registro);
    setCotizaciones((p) => [...p, { ...registro, ID: id, FECHA: todayISO(),
      RESPONSABLE: sesion.usuario }]);
    return id;
  }

  function handleImport(filas) {
    /* Solo alimenta la vista: el histórico sigue viviendo en el Excel. */
    setCotizaciones((p) => {
      const ids = new Set(p.map((c) => c.ID));
      const nuevas = filas.filter((f) => !ids.has(f.ID));
      notify(`${nuevas.length} cotizaciones cargadas para consulta.`);
      return [...p, ...nuevas];
    });
  }

  if (!sesion) return <LoginScreen onLogin={setSesion} />;

  return (
    <>
      <MarcaDeAgua />

      <header className="topbar">
        <div className="topbar-in">
          <LogoAmbipar />
          <span className="divider" />
          <div className="app-title">
            <p className="premium-heading">Cotizaciones de Valorización</p>
            <p>Ambipar Chile</p>
          </div>
          <div className="topbar-right">
            <div className="navtabs">
              <button aria-pressed={vista === 'form'} onClick={() => setVista('form')}>Nueva</button>
              <button aria-pressed={vista === 'hist'} onClick={() => setVista('hist')}>Historial</button>
            </div>
            <div className="pill" title={errorCarga ? 'Sin conexión con la base' : 'Conectado a la base compartida'}>
              <span className="dot" style={errorCarga ? { background: '#FF6B5A' } : undefined} />
              <span>{errorCarga ? 'Sin conexión' : sesion.usuario}</span>
            </div>
            <button className="link-btn" onClick={() => setSesion(null)}>Salir</button>
          </div>
        </div>
      </header>

      <div className={'canvas' + (vista === 'hist' ? ' ancho' : '')}>
        {errorCarga && (
          <div className="aviso malo" style={{ marginBottom: '1.25rem' }}>
            <b>No se pudo conectar con la base compartida.</b> Se está mostrando el
            catálogo local: puede consultar, pero lo que registre no se guardará.
            Detalle: {errorCarga}
          </div>
        )}

        {!cargado ? (
          <div className="empty-note" style={{ marginTop: '3rem' }}>Cargando la base…</div>
        ) : vista === 'form' ? (
          <QuoteForm sesion={sesion} destinatarios={destinatarios}
            onSaveQuote={handleSaveQuote} onAddDestinatario={handleAddDestinatario} notify={notify} />
        ) : (
          <HistorialPanel cotizaciones={cotizaciones} onImport={handleImport} />
        )}
      </div>

      <footer className="footer">
        <div className="footer-in">
          <span>Ambipar Chile - Cotizaciones de Valorización</span>
          <span>Inteligencia Circular LATAM</span>
        </div>
      </footer>

      <div id="toast" className={(toast.txt ? 'on' : '') + (toast.malo ? ' malo' : '')}>
        {toast.txt}
      </div>
    </>
  );
}
