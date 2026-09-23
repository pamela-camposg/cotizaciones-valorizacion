import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import * as api from './api.js';

import { RESIDUOS_TREE, LISTAS, DESTINATARIOS_SEED } from './data.js';

const COLORS = {
  teal: '#032024',
  tealSoft: '#0B3A40',
  lime: '#CDFF00',
  cream: '#F5F4ED',
  creamDark: '#E8E5D8',
  ink: '#0F1B1C',
  gray: '#5B6B6C',
  border: '#D8D4C4',
  danger: '#C1440E',
};

const REGION_BY_COMUNA = Object.fromEntries(LISTAS.COMUNA_REGION);
const COMUNAS = LISTAS.COMUNA_REGION.map((c) => c[0]);

function splitKey(key) {
  const i = key.lastIndexOf('|');
  return { name: key.slice(0, i), code: key.slice(i + 1) };
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function fmtCLP(n) {
  if (n === '' || n === null || n === undefined || isNaN(n)) return '';
  return new Intl.NumberFormat('es-CL').format(n);
}

/* ============================== ALMACENAMIENTO ============================== */


/* ============================== UI ATOMS ============================== */
function Field({ label, required, hint, children, auto }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{
        display: 'flex', alignItems: 'baseline', gap: 6,
        fontSize: 12.5, fontWeight: 600, color: COLORS.teal,
        letterSpacing: '0.01em', marginBottom: 5,
      }}>
        {label}
        {required && <span style={{ color: COLORS.danger }}>*</span>}
        {auto && <span style={{ fontSize: 10.5, fontWeight: 500, color: COLORS.gray }}>· automático</span>}
      </label>
      {children}
      {hint && <div style={{ fontSize: 11, color: COLORS.gray, marginTop: 3 }}>{hint}</div>}
    </div>
  );
}

const inputBase = {
  width: '100%', boxSizing: 'border-box', padding: '9px 11px',
  borderRadius: 6, border: `1px solid ${COLORS.border}`,
  fontSize: 13.5, fontFamily: 'inherit', color: COLORS.ink,
  background: '#fff', outline: 'none',
};
const inputAuto = {
  ...inputBase, background: COLORS.creamDark, color: COLORS.gray, cursor: 'not-allowed',
};

function TextInput(props) {
  const { auto, ...rest } = props;
  return <input {...rest} style={auto ? inputAuto : inputBase} readOnly={auto || rest.readOnly} />;
}

function Select({ auto, children, ...rest }) {
  return <select {...rest} style={auto ? inputAuto : inputBase} disabled={auto || rest.disabled}>{children}</select>;
}

function RadioGroup({ options, value, onChange, disabled }) {
  return (
    <div style={{ display: 'flex', gap: 18 }}>
      {options.map((opt) => (
        <label key={opt} style={{
          display: 'flex', alignItems: 'center', gap: 6, fontSize: 13.5,
          color: disabled ? COLORS.gray : COLORS.ink, cursor: disabled ? 'default' : 'pointer',
        }}>
          <input type="radio" checked={value === opt} disabled={disabled}
            onChange={() => onChange(opt)}
            style={{ accentColor: COLORS.teal, width: 15, height: 15 }} />
          {opt}
        </label>
      ))}
    </div>
  );
}

function SectionCard({ color, title, subtitle, shortcut, children }) {
  return (
    <div style={{
      background: '#fff', border: `1px solid ${COLORS.border}`, borderRadius: 10,
      marginBottom: 16, overflow: 'hidden',
    }}>
      <div style={{
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        padding: '11px 16px', background: color, borderBottom: `1px solid ${COLORS.border}`,
      }}>
        <div>
          <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.teal }}>{title}</div>
          {subtitle && <div style={{ fontSize: 11, color: COLORS.gray, marginTop: 1 }}>{subtitle}</div>}
        </div>
        {shortcut && (
          <div style={{
            fontSize: 10.5, color: COLORS.teal, background: '#fff', border: `1px solid ${COLORS.border}`,
            borderRadius: 5, padding: '3px 8px', fontFamily: 'ui-monospace, monospace',
          }}>{shortcut}</div>
        )}
      </div>
      <div style={{ padding: 16 }}>{children}</div>
    </div>
  );
}

function Toast({ message, type }) {
  if (!message) return null;
  return (
    <div style={{
      position: 'fixed', bottom: 20, left: '50%', transform: 'translateX(-50%)',
      background: type === 'error' ? COLORS.danger : COLORS.teal, color: '#fff',
      padding: '11px 20px', borderRadius: 8, fontSize: 13.5, fontWeight: 600,
      boxShadow: '0 8px 24px rgba(0,0,0,0.25)', zIndex: 1000, maxWidth: 420, textAlign: 'center',
    }}>{message}</div>
  );
}

/* ============================== LOGIN ============================== */
function LoginScreen({ onLogin }) {
  const [user, setUser] = useState('');
  const [pin, setPin] = useState('');
  const [error, setError] = useState('');
  const [cargando, setCargando] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    if (!user) { setError('Seleccione su usuario.'); return; }
    if (!pin.trim()) { setError('Ingrese su PIN.'); return; }

    setError('');
    setCargando(true);
    try {
      // La validación ocurre en el servidor, contra la hoja USUARIOS.
      const sesion = await api.login(user.trim(), pin.trim());
      onLogin({ usuario: sesion.usuario, pin: pin.trim(), nombre: sesion.nombre });
    } catch (err) {
      setError(err.message);
      setCargando(false);
    }
  };

  return (
    <div style={{
      minHeight: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center',
      background: `linear-gradient(160deg, ${COLORS.teal} 0%, #051A1D 100%)`,
      fontFamily: "'Inter', system-ui, sans-serif", padding: 24,
    }}>
      <form onSubmit={submit} style={{
        width: 360, background: COLORS.cream, borderRadius: 14, padding: '30px 28px',
        boxShadow: '0 20px 60px rgba(0,0,0,0.4)',
      }}>
        <div style={{
          width: 40, height: 6, background: COLORS.lime, borderRadius: 3, marginBottom: 18,
        }} />
        <h1 style={{ margin: '0 0 4px', fontSize: 19, color: COLORS.teal, fontWeight: 800 }}>
          Cotizaciones de Valorización
        </h1>
        <p style={{ margin: '0 0 22px', fontSize: 12.5, color: COLORS.gray, lineHeight: 1.5 }}>
          Acceso restringido — equipo de Inteligencia Circular y Gerencia de Valorización.
        </p>

        <Field label="Usuario" required>
          <Select value={user} onChange={(e) => setUser(e.target.value)} disabled={cargando}>
            <option value="">Seleccione su usuario...</option>
            {LISTAS.RESPONSABLE.map((r) => <option key={r} value={r}>{r}</option>)}
          </Select>
        </Field>

        <Field label="PIN de acceso" required>
          <TextInput type="password" value={pin} onChange={(e) => setPin(e.target.value)}
            placeholder="••••" disabled={cargando} />
        </Field>

        {error && <div style={{ color: COLORS.danger, fontSize: 12.5, marginBottom: 12 }}>{error}</div>}

        <button type="submit" style={{
          width: '100%', padding: '11px 0', border: 'none', borderRadius: 8,
          background: COLORS.lime, color: COLORS.teal, fontWeight: 800, fontSize: 13.5,
          cursor: 'pointer', marginTop: 6,
        }} disabled={cargando}>{cargando ? 'Verificando...' : 'Ingresar'}</button>

        <p style={{ fontSize: 10.5, color: COLORS.gray, marginTop: 16, lineHeight: 1.5 }}>
          El usuario y el PIN se validan en el servidor, contra la hoja USUARIOS del
          Google Sheet. Para revocar un acceso, cambie ACTIVO a NO en esa hoja.
        </p>
      </form>
    </div>
  );
}

/* ============================== BUSCADOR DE RESIDUO (F2) ============================== */
function ResidueSearchModal({ onClose, onSelect }) {
  const [q, setQ] = useState('');
  const results = useMemo(() => {
    if (q.trim().length < 2) return [];
    const needle = q.toLowerCase();
    const out = [];
    for (const pk of Object.keys(RESIDUOS_TREE)) {
      for (const fk of Object.keys(RESIDUOS_TREE[pk])) {
        for (const mk of Object.keys(RESIDUOS_TREE[pk][fk])) {
          for (const tk of Object.keys(RESIDUOS_TREE[pk][fk][mk])) {
            const tName = splitKey(tk).name;
            if (tName.toLowerCase().includes(needle)) {
              out.push({ pk, fk, mk, tk });
            }
          }
        }
      }
    }
    return out.slice(0, 30);
  }, [q]);

  return (
    <div style={{
      position: 'fixed', inset: 0, background: 'rgba(3,32,36,0.55)', zIndex: 900,
      display: 'flex', alignItems: 'flex-start', justifyContent: 'center', paddingTop: '10vh',
    }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{
        width: 480, maxHeight: '70vh', background: '#fff', borderRadius: 12, overflow: 'hidden',
        boxShadow: '0 30px 80px rgba(0,0,0,0.4)', display: 'flex', flexDirection: 'column',
      }}>
        <div style={{ padding: 14, borderBottom: `1px solid ${COLORS.border}` }}>
          <div style={{ fontSize: 11, color: COLORS.gray, marginBottom: 6 }}>
            Buscador de residuos por TIPO · [F2]
          </div>
          <TextInput autoFocus value={q} onChange={(e) => setQ(e.target.value)}
            placeholder="Escriba, ej: neumático, aceite, cartón..." />
        </div>
        <div style={{ overflowY: 'auto' }}>
          {results.length === 0 && (
            <div style={{ padding: 16, fontSize: 12.5, color: COLORS.gray }}>
              {q.trim().length < 2 ? 'Escriba al menos 2 caracteres.' : 'Sin resultados.'}
            </div>
          )}
          {results.map((r, i) => {
            const p = splitKey(r.pk), f = splitKey(r.fk), m = splitKey(r.mk), t = splitKey(r.tk);
            return (
              <div key={i} onClick={() => onSelect(r)} style={{
                padding: '10px 14px', cursor: 'pointer', fontSize: 12.5,
                borderBottom: `1px solid ${COLORS.creamDark}`,
              }}
                onMouseEnter={(e) => e.currentTarget.style.background = COLORS.cream}
                onMouseLeave={(e) => e.currentTarget.style.background = '#fff'}>
                <div style={{ fontWeight: 700, color: COLORS.teal }}>{t.name}</div>
                <div style={{ color: COLORS.gray, fontSize: 11 }}>{p.name} · {f.name} · {m.name}</div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ============================== FORMULARIO ============================== */
const EMPTY_LOGISTICA = {
  tipoServicio: 'En planta', canalCotizacion: 'Correo', comunaOrigen: '', origen: '',
};
const EMPTY_ESTABLECIMIENTO = {
  nombreFantasia: '', comunaEstablecimiento: '', sistemaDeclaracion: '',
  codigoEstablecimiento: '', razonSocial: '', rut: '', nombreEstablecimiento: '',
  regionEstablecimiento: '', esNuevo: false,
};
const EMPTY_RESIDUO = { pk: '', fk: '', mk: '', tk: '', ek: '', fok: '', otroNombre: '' };
const EMPTY_COMERCIAL = { tipoNegocio: 'Venta', unidad: '', valor: '' };
const EMPTY_ADJUNTOS = { documentos: '', comentarios: '' };

function QuoteForm({ user, destinatarios, onSaveQuote, onAddDestinatario, notify }) {
  const [logistica, setLogistica] = useState(EMPTY_LOGISTICA);
  const [establecimiento, setEstablecimiento] = useState(EMPTY_ESTABLECIMIENTO);
  const [residuo, setResiduo] = useState(EMPTY_RESIDUO);
  const [comercial, setComercial] = useState(EMPTY_COMERCIAL);
  const [adjuntos, setAdjuntos] = useState(EMPTY_ADJUNTOS);
  const [showSearch, setShowSearch] = useState(false);
  const [guardando, setGuardando] = useState(false);
  const [fantasiaQuery, setFantasiaQuery] = useState('');
  const formRef = useRef(null);

  useEffect(() => {
    const handler = (e) => {
      if (e.key === 'F2') { e.preventDefault(); setShowSearch(true); }
      if (e.key === 'Escape') { setShowSearch(false); }
      if (e.ctrlKey && e.key === 'Enter') { e.preventDefault(); handleSave(); }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
    // eslint-disable-next-line
  }, [logistica, establecimiento, residuo, comercial, adjuntos]);

  const nombresFantasia = useMemo(
    () => [...new Set(destinatarios.map((d) => d.NOMBRE_FANTASIA))].sort(),
    [destinatarios]
  );
  const matchesEstablecimiento = fantasiaQuery.length > 0
    ? nombresFantasia.filter((n) => n.toLowerCase().includes(fantasiaQuery.toLowerCase()))
    : [];
  const exactMatch = nombresFantasia.includes(establecimiento.nombreFantasia);

  const comunasParaFantasia = useMemo(() => {
    if (!establecimiento.nombreFantasia) return [];
    return [...new Set(destinatarios
      .filter((d) => d.NOMBRE_FANTASIA === establecimiento.nombreFantasia)
      .map((d) => d.COMUNA_ESTABLECIMIENTO))];
  }, [destinatarios, establecimiento.nombreFantasia]);

  const sistemasParaSeleccion = useMemo(() => {
    if (!establecimiento.nombreFantasia || !establecimiento.comunaEstablecimiento) return [];
    return [...new Set(destinatarios
      .filter((d) => d.NOMBRE_FANTASIA === establecimiento.nombreFantasia
        && d.COMUNA_ESTABLECIMIENTO === establecimiento.comunaEstablecimiento)
      .map((d) => d.SISTEMA_DECLARACION))];
  }, [destinatarios, establecimiento.nombreFantasia, establecimiento.comunaEstablecimiento]);

  function pickFantasia(name) {
    const isNew = !nombresFantasia.includes(name);
    setEstablecimiento({ ...EMPTY_ESTABLECIMIENTO, nombreFantasia: name, esNuevo: isNew });
    setFantasiaQuery('');
  }

  function pickComuna(comuna) {
    setEstablecimiento((prev) => ({ ...prev, comunaEstablecimiento: comuna, sistemaDeclaracion: '', codigoEstablecimiento: '', razonSocial: '', rut: '', nombreEstablecimiento: '', regionEstablecimiento: '' }));
  }

  function pickSistema(sistema) {
    const row = destinatarios.find((d) =>
      d.NOMBRE_FANTASIA === establecimiento.nombreFantasia &&
      d.COMUNA_ESTABLECIMIENTO === establecimiento.comunaEstablecimiento &&
      d.SISTEMA_DECLARACION === sistema);
    if (row) {
      setEstablecimiento((prev) => ({
        ...prev, sistemaDeclaracion: sistema,
        codigoEstablecimiento: row.CODIGO_ESTABLECIMIENTO, razonSocial: row.RAZON_SOCIAL,
        rut: row.RUT, nombreEstablecimiento: row.NOMBRE_ESTABLECIMIENTO,
        regionEstablecimiento: row.REGION_ESTABLECIMIENTO,
      }));
    }
  }

  // Cascada de residuo
  const pOptions = Object.keys(RESIDUOS_TREE);
  const fOptions = residuo.pk ? Object.keys(RESIDUOS_TREE[residuo.pk]) : [];
  const mOptions = residuo.pk && residuo.fk ? Object.keys(RESIDUOS_TREE[residuo.pk][residuo.fk]) : [];
  const tOptions = residuo.pk && residuo.fk && residuo.mk ? Object.keys(RESIDUOS_TREE[residuo.pk][residuo.fk][residuo.mk]) : [];
  const eOptions = residuo.pk && residuo.fk && residuo.mk && residuo.tk ? Object.keys(RESIDUOS_TREE[residuo.pk][residuo.fk][residuo.mk][residuo.tk]) : [];
  const foOptions = residuo.pk && residuo.fk && residuo.mk && residuo.tk && residuo.ek
    ? RESIDUOS_TREE[residuo.pk][residuo.fk][residuo.mk][residuo.tk][residuo.ek] : [];

  const codRes = useMemo(() => {
    if (!(residuo.pk && residuo.fk && residuo.mk && residuo.tk && residuo.ek && residuo.fok)) return '';
    return [residuo.pk, residuo.fk, residuo.mk, residuo.tk, residuo.ek, residuo.fok]
      .map((k) => splitKey(k).code).join('-');
  }, [residuo]);

  function onResidueSearchSelect(r) {
    setResiduo({ ...EMPTY_RESIDUO, pk: r.pk, fk: r.fk, mk: r.mk, tk: r.tk });
    setShowSearch(false);
  }

  function resetAfterSave() {
    // Apartados 1 y 2 se mantienen precargados para la siguiente carga (según especificación).
    setResiduo(EMPTY_RESIDUO);
    setComercial(EMPTY_COMERCIAL);
    setAdjuntos(EMPTY_ADJUNTOS);
  }

  function validate() {
    if (!residuo.pk || !residuo.fk || !residuo.mk || !residuo.tk || !residuo.ek || !residuo.fok) {
      return 'Complete la caracterización del residuo (Apartado 3).';
    }
    if (!establecimiento.nombreFantasia) return 'Ingrese el nombre de fantasía del establecimiento.';
    if (establecimiento.esNuevo && (!establecimiento.razonSocial || !establecimiento.rut || !establecimiento.comunaEstablecimiento)) {
      return 'Complete razón social, RUT y comuna del nuevo destinatario.';
    }
    if (logistica.tipoServicio === 'Con retiro' && !logistica.comunaOrigen) {
      return 'Seleccione la comuna de origen (servicio con retiro).';
    }
    if (!comercial.unidad || comercial.valor === '') return 'Complete la condición comercial (Apartado 4).';
    return null;
  }

  async function handleSave() {
    const err = validate();
    if (err) { notify(err, 'error'); return; }

    setGuardando(true);
    try {
    if (establecimiento.esNuevo) {
      await onAddDestinatario({
        NOMBRE_FANTASIA: establecimiento.nombreFantasia,
        SISTEMA_DECLARACION: establecimiento.sistemaDeclaracion || 'N/A',
        CODIGO_ESTABLECIMIENTO: establecimiento.codigoEstablecimiento || '(pendiente)',
        RAZON_SOCIAL: establecimiento.razonSocial,
        RUT: establecimiento.rut,
        NOMBRE_ESTABLECIMIENTO: establecimiento.nombreEstablecimiento || establecimiento.razonSocial,
        COMUNA_ESTABLECIMIENTO: establecimiento.comunaEstablecimiento,
        REGION_ESTABLECIMIENTO: REGION_BY_COMUNA[establecimiento.comunaEstablecimiento] || '',
      });
    }

    const record = {
      FECHA: todayISO(),
      TIPO_SERVICIO: logistica.tipoServicio,
      CANAL_COTIZACION: logistica.canalCotizacion,
      ORIGEN: logistica.origen,
      COMUNA_ORIGEN: logistica.comunaOrigen,
      REGION_ORIGEN: REGION_BY_COMUNA[logistica.comunaOrigen] || '',
      NOMBRE_FANTASIA: establecimiento.nombreFantasia,
      COMUNA_ESTABLECIMIENTO: establecimiento.comunaEstablecimiento,
      REGION_ESTABLECIMIENTO: establecimiento.esNuevo ? (REGION_BY_COMUNA[establecimiento.comunaEstablecimiento] || '') : establecimiento.regionEstablecimiento,
      SISTEMA_DECLARACION: establecimiento.sistemaDeclaracion,
      CODIGO_ESTABLECIMIENTO: establecimiento.codigoEstablecimiento,
      RAZON_SOCIAL: establecimiento.razonSocial,
      RUT: establecimiento.rut,
      NOMBRE_ESTABLECIMIENTO: establecimiento.nombreEstablecimiento,
      COD_RES: codRes,
      PELIGROSIDAD: splitKey(residuo.pk).name,
      FAMILIA: splitKey(residuo.fk).name,
      MATERIAL: splitKey(residuo.mk).name,
      TIPO: splitKey(residuo.tk).name,
      ESTADO: splitKey(residuo.ek).name,
      FORMATO: splitKey(residuo.fok).name,
      OTRO_NOMBRE: residuo.otroNombre,
      TIPO_NEGOCIO: comercial.tipoNegocio,
      VALOR: Number(comercial.valor),
      UNIDAD: comercial.unidad,
      RESPONSABLE: user,
      DOCUMENTOS: adjuntos.documentos,
      COMENTARIOS: adjuntos.comentarios,
    };

    const id = await onSaveQuote(record);
    notify(`Cotización ${id} guardada en la base de datos.`, 'ok');
    resetAfterSave();
    } catch (err) {
      // El formulario NO se limpia: así no se pierde lo escrito y se
      // puede reintentar.
      notify('No se pudo guardar: ' + err.message, 'error');
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div ref={formRef} style={{ maxWidth: 720, margin: '0 auto' }}>
      {showSearch && <ResidueSearchModal onClose={() => setShowSearch(false)} onSelect={onResidueSearchSelect} />}

      <SectionCard color="#EAF3EC" title="Apartado 1 · Logística y control general"
        subtitle="Fecha, ID y responsable se generan automáticamente al guardar."
        shortcut="Tab avanza · F1 clona 1+2">
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14, marginBottom: 4 }}>
          <Field label="Fecha" auto><TextInput auto value={todayISO()} /></Field>
          <Field label="Responsable" auto><TextInput auto value={user} /></Field>
        </div>
        <Field label="Tipo de servicio" required>
          <RadioGroup options={LISTAS.TIPO_SERVICIO} value={logistica.tipoServicio}
            onChange={(v) => setLogistica((p) => ({ ...p, tipoServicio: v, comunaOrigen: v === 'En planta' ? '' : p.comunaOrigen, origen: v === 'En planta' ? '' : p.origen }))} />
        </Field>
        <Field label="Canal de cotización" required>
          <RadioGroup options={LISTAS.CANAL_COTIZACION} value={logistica.canalCotizacion}
            onChange={(v) => setLogistica((p) => ({ ...p, canalCotizacion: v }))} />
        </Field>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
          <Field label="Comuna de origen" required={logistica.tipoServicio === 'Con retiro'}
            hint={logistica.tipoServicio === 'En planta' ? 'No aplica en planta.' : undefined}>
            <Select disabled={logistica.tipoServicio === 'En planta'} value={logistica.comunaOrigen}
              onChange={(e) => setLogistica((p) => ({ ...p, comunaOrigen: e.target.value }))}>
              <option value="">Seleccione comuna...</option>
              {COMUNAS.map((c) => <option key={c} value={c}>{c}</option>)}
            </Select>
          </Field>
          <Field label="Región de origen" auto><TextInput auto value={REGION_BY_COMUNA[logistica.comunaOrigen] || ''} /></Field>
        </div>
        <Field label="Origen (dirección / referencia)" hint="Manual, solo si es con retiro.">
          <TextInput disabled={logistica.tipoServicio === 'En planta'} value={logistica.origen}
            onChange={(e) => setLogistica((p) => ({ ...p, origen: e.target.value }))} />
        </Field>
      </SectionCard>

      <SectionCard color="#EAF0F3" title="Apartado 2 · Identificación del establecimiento"
        subtitle="Escriba el nombre de fantasía; si ya existe, el resto se completa solo.">
        <Field label="Nombre de fantasía" required>
          <TextInput
            value={establecimiento.nombreFantasia || fantasiaQuery}
            onChange={(e) => { setFantasiaQuery(e.target.value); setEstablecimiento(EMPTY_ESTABLECIMIENTO); }}
            placeholder="Escriba para buscar empresa..." />
          {fantasiaQuery && matchesEstablecimiento.length > 0 && (
            <div style={{ border: `1px solid ${COLORS.border}`, borderRadius: 6, marginTop: 4, maxHeight: 160, overflowY: 'auto', background: '#fff' }}>
              {matchesEstablecimiento.slice(0, 8).map((n) => (
                <div key={n} onClick={() => pickFantasia(n)} style={{ padding: '7px 10px', fontSize: 12.5, cursor: 'pointer' }}
                  onMouseEnter={(e) => e.currentTarget.style.background = COLORS.cream}
                  onMouseLeave={(e) => e.currentTarget.style.background = '#fff'}>{n}</div>
              ))}
            </div>
          )}
          {fantasiaQuery && matchesEstablecimiento.length === 0 && (
            <div style={{ marginTop: 6, fontSize: 12, color: COLORS.gray }}>
              Sin coincidencias.{' '}
              <span style={{ color: COLORS.teal, fontWeight: 700, cursor: 'pointer', textDecoration: 'underline' }}
                onClick={() => pickFantasia(fantasiaQuery)}>
                Registrar "{fantasiaQuery}" como nuevo destinatario
              </span>
            </div>
          )}
        </Field>

        {establecimiento.nombreFantasia && !establecimiento.esNuevo && (
          <>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
              <Field label="Comuna del establecimiento" required>
                <Select value={establecimiento.comunaEstablecimiento} onChange={(e) => pickComuna(e.target.value)}>
                  <option value="">Seleccione...</option>
                  {comunasParaFantasia.map((c) => <option key={c} value={c}>{c}</option>)}
                </Select>
              </Field>
              <Field label="Sistema de declaración" required>
                <Select disabled={!establecimiento.comunaEstablecimiento} value={establecimiento.sistemaDeclaracion} onChange={(e) => pickSistema(e.target.value)}>
                  <option value="">Seleccione...</option>
                  {sistemasParaSeleccion.map((s) => <option key={s} value={s}>{s}</option>)}
                </Select>
              </Field>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
              <Field label="Código de establecimiento" auto><TextInput auto value={establecimiento.codigoEstablecimiento} /></Field>
              <Field label="RUT" auto><TextInput auto value={establecimiento.rut} /></Field>
            </div>
            <Field label="Razón social" auto><TextInput auto value={establecimiento.razonSocial} /></Field>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
              <Field label="Nombre del establecimiento" auto><TextInput auto value={establecimiento.nombreEstablecimiento} /></Field>
              <Field label="Región del establecimiento" auto><TextInput auto value={establecimiento.regionEstablecimiento} /></Field>
            </div>
          </>
        )}

        {establecimiento.esNuevo && (
          <div style={{ background: '#FFF7E6', border: '1px solid #E8CE8B', borderRadius: 8, padding: 12, marginTop: 4 }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: '#7A5A00', marginBottom: 10 }}>
              Nuevo destinatario — se agregará a la base de datos DESTINATARIOS al guardar.
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
              <Field label="Razón social" required>
                <TextInput value={establecimiento.razonSocial} onChange={(e) => setEstablecimiento((p) => ({ ...p, razonSocial: e.target.value }))} />
              </Field>
              <Field label="RUT" required>
                <TextInput value={establecimiento.rut} onChange={(e) => setEstablecimiento((p) => ({ ...p, rut: e.target.value }))} placeholder="12345678-9" />
              </Field>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
              <Field label="Comuna del establecimiento" required>
                <Select value={establecimiento.comunaEstablecimiento}
                  onChange={(e) => setEstablecimiento((p) => ({ ...p, comunaEstablecimiento: e.target.value }))}>
                  <option value="">Seleccione...</option>
                  {COMUNAS.map((c) => <option key={c} value={c}>{c}</option>)}
                </Select>
              </Field>
              <Field label="Sistema de declaración" required>
                <Select value={establecimiento.sistemaDeclaracion}
                  onChange={(e) => setEstablecimiento((p) => ({ ...p, sistemaDeclaracion: e.target.value }))}>
                  <option value="">Seleccione...</option>
                  <option value="SINADER">SINADER</option>
                  <option value="N/A">N/A</option>
                </Select>
              </Field>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
              <Field label="Código de establecimiento" hint="Si aún no se conoce, se guarda como pendiente.">
                <TextInput value={establecimiento.codigoEstablecimiento} onChange={(e) => setEstablecimiento((p) => ({ ...p, codigoEstablecimiento: e.target.value }))} />
              </Field>
              <Field label="Nombre del establecimiento">
                <TextInput value={establecimiento.nombreEstablecimiento} onChange={(e) => setEstablecimiento((p) => ({ ...p, nombreEstablecimiento: e.target.value }))} />
              </Field>
            </div>
          </div>
        )}
      </SectionCard>

      <SectionCard color="#F1EAF3" title="Apartado 3 · Caracterización del residuo"
        shortcut="F2 buscar por tipo">
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
          <Field label="Peligrosidad" required>
            <Select value={residuo.pk} onChange={(e) => setResiduo({ ...EMPTY_RESIDUO, pk: e.target.value })}>
              <option value="">Seleccione...</option>
              {pOptions.map((k) => <option key={k} value={k}>{splitKey(k).name}</option>)}
            </Select>
          </Field>
          <Field label="Familia" required>
            <Select disabled={!residuo.pk} value={residuo.fk} onChange={(e) => setResiduo((p) => ({ ...EMPTY_RESIDUO, pk: p.pk, fk: e.target.value }))}>
              <option value="">Seleccione...</option>
              {fOptions.map((k) => <option key={k} value={k}>{splitKey(k).name}</option>)}
            </Select>
          </Field>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
          <Field label="Material" required>
            <Select disabled={!residuo.fk} value={residuo.mk} onChange={(e) => setResiduo((p) => ({ ...p, mk: e.target.value, tk: '', ek: '', fok: '' }))}>
              <option value="">Seleccione...</option>
              {mOptions.map((k) => <option key={k} value={k}>{splitKey(k).name}</option>)}
            </Select>
          </Field>
          <Field label="Tipo" required>
            <Select disabled={!residuo.mk} value={residuo.tk} onChange={(e) => setResiduo((p) => ({ ...p, tk: e.target.value, ek: '', fok: '' }))}>
              <option value="">Seleccione...</option>
              {tOptions.map((k) => <option key={k} value={k}>{splitKey(k).name}</option>)}
            </Select>
          </Field>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
          <Field label="Estado" required>
            <Select disabled={!residuo.tk} value={residuo.ek} onChange={(e) => setResiduo((p) => ({ ...p, ek: e.target.value, fok: '' }))}>
              <option value="">Seleccione...</option>
              {eOptions.map((k) => <option key={k} value={k}>{splitKey(k).name}</option>)}
            </Select>
          </Field>
          <Field label="Formato" required>
            <Select disabled={!residuo.ek} value={residuo.fok} onChange={(e) => setResiduo((p) => ({ ...p, fok: e.target.value }))}>
              <option value="">Seleccione...</option>
              {foOptions.map((k) => <option key={k} value={k}>{splitKey(k).name}</option>)}
            </Select>
          </Field>
        </div>
        <Field label="Código de residuo (COD_RES)" auto><TextInput auto value={codRes} /></Field>
        <Field label="Otro nombre" hint="Manual, opcional (máx. 40 caracteres). Nombre comercial.">
          <TextInput maxLength={40} value={residuo.otroNombre} onChange={(e) => setResiduo((p) => ({ ...p, otroNombre: e.target.value }))} />
        </Field>
      </SectionCard>

      <SectionCard color="#F5EDE5" title="Apartado 4 · Condición comercial">
        <Field label="Tipo de negocio" required>
          <RadioGroup options={LISTAS.TIPO_NEGOCIO} value={comercial.tipoNegocio}
            onChange={(v) => setComercial((p) => ({ ...p, tipoNegocio: v }))} />
        </Field>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
          <Field label="Unidad" required>
            <Select value={comercial.unidad} onChange={(e) => setComercial((p) => ({ ...p, unidad: e.target.value }))}>
              <option value="">Seleccione moneda/medida...</option>
              {LISTAS.UNIDAD.map((u) => <option key={u} value={u}>{u}</option>)}
            </Select>
          </Field>
          <Field label="Valor" required>
            <TextInput type="number" min="0" value={comercial.valor}
              onChange={(e) => setComercial((p) => ({ ...p, valor: e.target.value }))} placeholder="Ingrese monto" />
          </Field>
        </div>
      </SectionCard>

      <SectionCard color="#EDEDED" title="Apartado 5 · Adjuntos y comentarios">
        <Field label="Referencia de archivo" hint='Nombre/enlace del documento (ej. repositorio SharePoint). Se renombrará como ID_Fecha_NombreFantasia.'>
          <TextInput value={adjuntos.documentos} onChange={(e) => setAdjuntos((p) => ({ ...p, documentos: e.target.value }))} placeholder="Ej: cotizacion_polambiente.pdf" />
        </Field>
        <Field label="Comentarios" hint="Manual, libre (máx. 100 caracteres).">
          <textarea maxLength={100} value={adjuntos.comentarios}
            onChange={(e) => setAdjuntos((p) => ({ ...p, comentarios: e.target.value }))}
            style={{ ...inputBase, minHeight: 60, resize: 'vertical' }} />
        </Field>
      </SectionCard>

      <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginBottom: 40 }}>
        <button onClick={resetAfterSave} style={{
          padding: '10px 18px', borderRadius: 7, border: `1px solid ${COLORS.border}`,
          background: '#fff', color: COLORS.teal, fontSize: 13, fontWeight: 700, cursor: 'pointer',
        }}>Cancelar (Esc)</button>
        <button onClick={handleSave} style={{
          padding: '10px 22px', borderRadius: 7, border: 'none',
          background: COLORS.lime, color: COLORS.teal, fontSize: 13, fontWeight: 800, cursor: 'pointer',
        }} disabled={guardando}>{guardando ? 'Guardando...' : 'Guardar (Ctrl+Enter)'}</button>
      </div>
    </div>
  );
}

/* ============================== HISTORIAL ============================== */
function HistorialPanel({ cotizaciones, onImport }) {
  const [filtroMaterial, setFiltroMaterial] = useState('');
  const [filtroZona, setFiltroZona] = useState('');
  const fileInput = useRef(null);

  const materiales = useMemo(() => [...new Set(cotizaciones.map((c) => c.MATERIAL))].filter(Boolean).sort(), [cotizaciones]);
  const zonas = useMemo(() => [...new Set(cotizaciones.map((c) => c.REGION_ESTABLECIMIENTO))].filter(Boolean).sort(), [cotizaciones]);

  const filtradas = cotizaciones.filter((c) =>
    (!filtroMaterial || c.MATERIAL === filtroMaterial) &&
    (!filtroZona || c.REGION_ESTABLECIMIENTO === filtroZona)
  ).sort((a, b) => (a.FECHA < b.FECHA ? 1 : -1));

  const promedio = filtradas.length
    ? filtradas.reduce((s, c) => s + (Number(c.VALOR) || 0), 0) / filtradas.length
    : 0;

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: 'array', cellDates: true });
      if (!wb.SheetNames.includes('COTIZACIONES')) {
        alert('El archivo no contiene una hoja "COTIZACIONES".');
        return;
      }
      const ws = wb.Sheets['COTIZACIONES'];
      const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
      const normalized = json.map((r) => ({
        ...r,
        FECHA: r.FECHA instanceof Date ? r.FECHA.toISOString().slice(0, 10) : String(r.FECHA),
      }));
      onImport(normalized);
    };
    reader.readAsArrayBuffer(file);
  }

  return (
    <div style={{ maxWidth: 1000, margin: '0 auto' }}>
      <div style={{ display: 'flex', gap: 12, marginBottom: 18, flexWrap: 'wrap', alignItems: 'flex-end' }}>
        <div>
          <div style={{ fontSize: 11.5, fontWeight: 700, color: COLORS.teal, marginBottom: 4 }}>Material</div>
          <Select value={filtroMaterial} onChange={(e) => setFiltroMaterial(e.target.value)} style={{ minWidth: 180 }}>
            <option value="">Todos</option>
            {materiales.map((m) => <option key={m} value={m}>{m}</option>)}
          </Select>
        </div>
        <div>
          <div style={{ fontSize: 11.5, fontWeight: 700, color: COLORS.teal, marginBottom: 4 }}>Región</div>
          <Select value={filtroZona} onChange={(e) => setFiltroZona(e.target.value)} style={{ minWidth: 180 }}>
            <option value="">Todas</option>
            {zonas.map((z) => <option key={z} value={z}>{z}</option>)}
          </Select>
        </div>
        <div style={{ marginLeft: 'auto', display: 'flex', gap: 16, alignItems: 'center' }}>
          <div style={{ fontSize: 12, color: COLORS.gray }}>
            <b style={{ color: COLORS.teal }}>{filtradas.length}</b> cotizaciones ·
            valor promedio <b style={{ color: COLORS.teal }}>{fmtCLP(Math.round(promedio))}</b>
          </div>
          <button onClick={() => fileInput.current.click()} style={{
            padding: '8px 14px', borderRadius: 7, border: `1px solid ${COLORS.teal}`,
            background: '#fff', color: COLORS.teal, fontSize: 12, fontWeight: 700, cursor: 'pointer',
          }}>Importar histórico (.xlsx)</button>
          <input ref={fileInput} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleFile} />
        </div>
      </div>

      <div style={{ border: `1px solid ${COLORS.border}`, borderRadius: 10, overflow: 'hidden', background: '#fff' }}>
        <div style={{ maxHeight: 480, overflowY: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead style={{ position: 'sticky', top: 0, background: COLORS.teal, color: '#fff' }}>
              <tr>
                {['Fecha', 'ID', 'Material', 'Tipo', 'Establecimiento', 'Región', 'Negocio', 'Valor', 'Unidad', 'Responsable'].map((h) => (
                  <th key={h} style={{ textAlign: 'left', padding: '8px 10px', fontWeight: 700, whiteSpace: 'nowrap' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtradas.slice(0, 300).map((c, i) => (
                <tr key={i} style={{ borderBottom: `1px solid ${COLORS.creamDark}`, background: i % 2 ? COLORS.cream : '#fff' }}>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>{c.FECHA}</td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>{c.ID}</td>
                  <td style={{ padding: '7px 10px' }}>{c.MATERIAL}</td>
                  <td style={{ padding: '7px 10px' }}>{c.TIPO}</td>
                  <td style={{ padding: '7px 10px' }}>{c.NOMBRE_FANTASIA || c.NOMBRE_ESTABLECIMIENTO}</td>
                  <td style={{ padding: '7px 10px' }}>{c.REGION_ESTABLECIMIENTO}</td>
                  <td style={{ padding: '7px 10px' }}>{c.TIPO_NEGOCIO}</td>
                  <td style={{ padding: '7px 10px', textAlign: 'right' }}>{fmtCLP(c.VALOR)}</td>
                  <td style={{ padding: '7px 10px' }}>{c.UNIDAD}</td>
                  <td style={{ padding: '7px 10px' }}>{c.RESPONSABLE}</td>
                </tr>
              ))}
              {filtradas.length === 0 && (
                <tr><td colSpan={10} style={{ padding: 24, textAlign: 'center', color: COLORS.gray }}>
                  Aún no hay cotizaciones registradas con estos filtros. Cargue el histórico o registre una cotización nueva.
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ============================== APP ============================== */
export default function App() {
  // sesion = { usuario, pin, nombre }. Se mantiene solo en memoria: al
  // recargar la página hay que volver a ingresar el PIN.
  const [sesion, setSesion] = useState(null);
  const [view, setView] = useState('form');
  const [destinatarios, setDestinatarios] = useState([]);
  const [cotizaciones, setCotizaciones] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [errorCarga, setErrorCarga] = useState('');
  const [toast, setToast] = useState({ message: '', type: 'ok' });

  const user = sesion ? sesion.usuario : null;

  // Al iniciar sesión, trae destinatarios y cotizaciones desde el Sheet.
  useEffect(() => {
    if (!sesion) return;
    let cancelado = false;
    (async () => {
      setLoaded(false);
      setErrorCarga('');
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
        // Si el Sheet no responde, se usa el catálogo local como respaldo
        // para que al menos se pueda consultar, aunque no guardar.
        setDestinatarios(DESTINATARIOS_SEED.rows.map((r) => ({
          NOMBRE_FANTASIA: r[0],
          SISTEMA_DECLARACION: r[1],
          CODIGO_ESTABLECIMIENTO: r[2],
          RAZON_SOCIAL: r[3],
          RUT: r[4],
          NOMBRE_ESTABLECIMIENTO: r[5],
          COMUNA_ESTABLECIMIENTO: r[6],
          REGION_ESTABLECIMIENTO: r[7],
        })));
      } finally {
        if (!cancelado) setLoaded(true);
      }
    })();
    return () => { cancelado = true; };
  }, [sesion]);

  function notify(message, type) {
    setToast({ message, type });
    setTimeout(() => setToast({ message: '', type: 'ok' }), 4000);
  }

  async function handleAddDestinatario(newRow) {
    const yaExistia = await api.agregarDestinatario(sesion, newRow);
    if (!yaExistia) {
      setDestinatarios((prev) => [...prev, newRow]);
    }
  }

  async function handleSaveQuote(record) {
    // El ID, la FECHA y el RESPONSABLE los asigna el servidor, para que el
    // correlativo no se rompa si dos personas guardan a la vez.
    const id = await api.guardarCotizacion(sesion, record);
    setCotizaciones((prev) => [...prev, { ...record, ID: id }]);
    return id;
  }

  async function handleImportCotizaciones(rows) {
    // La importación solo alimenta la vista de historial; no escribe en el
    // Sheet, porque el histórico vive en el Excel.
    setCotizaciones((prev) => [...prev, ...rows]);
    notify(`${rows.length} cotizaciones cargadas en la vista de historial.`, 'ok');
    setView('historial');
  }

  if (!sesion) {
    return <LoginScreen onLogin={setSesion} />;
  }

  return (
    <div style={{ minHeight: '100%', background: COLORS.cream, fontFamily: "'Inter', system-ui, sans-serif" }}>
      <div style={{
        background: COLORS.teal, color: '#fff', padding: '14px 22px',
        display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 10,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 10, height: 10, borderRadius: 3, background: COLORS.lime }} />
          <div style={{ fontWeight: 800, fontSize: 14.5 }}>Cotizaciones de Valorización</div>
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <button onClick={() => setView('form')} style={{
            padding: '7px 14px', borderRadius: 6, border: 'none', cursor: 'pointer',
            background: view === 'form' ? COLORS.lime : 'transparent',
            color: view === 'form' ? COLORS.teal : '#fff', fontWeight: 700, fontSize: 12.5,
          }}>Nueva cotización</button>
          <button onClick={() => setView('historial')} style={{
            padding: '7px 14px', borderRadius: 6, border: 'none', cursor: 'pointer',
            background: view === 'historial' ? COLORS.lime : 'transparent',
            color: view === 'historial' ? COLORS.teal : '#fff', fontWeight: 700, fontSize: 12.5,
          }}>Historial</button>
          <div style={{ width: 1, height: 20, background: 'rgba(255,255,255,0.25)', margin: '0 4px' }} />
          <div style={{ fontSize: 12, color: '#CFE0E0' }}>{user}</div>
          <button onClick={() => setSesion(null)} style={{
            padding: '6px 10px', borderRadius: 6, border: '1px solid rgba(255,255,255,0.3)',
            background: 'transparent', color: '#fff', fontSize: 11.5, cursor: 'pointer',
          }}>Salir</button>
        </div>
      </div>

      <div style={{ padding: '26px 20px' }}>
        {errorCarga && (
          <div style={{
            maxWidth: 720, margin: '0 auto 16px', padding: '12px 16px',
            background: '#FDECEA', border: `1px solid ${COLORS.danger}`, borderRadius: 8,
            fontSize: 12.5, color: '#7A2410', lineHeight: 1.5,
          }}>
            <b>No se pudo conectar con la base de datos.</b> Se está mostrando el catálogo
            local, y lo que registre no se guardará. Detalle: {errorCarga}
          </div>
        )}
        {!loaded ? (
          <div style={{ textAlign: 'center', color: COLORS.gray, padding: 60 }}>Cargando datos...</div>
        ) : view === 'form' ? (
          <QuoteForm user={user} destinatarios={destinatarios} onSaveQuote={handleSaveQuote}
            onAddDestinatario={handleAddDestinatario} notify={notify} />
        ) : (
          <HistorialPanel cotizaciones={cotizaciones} onImport={handleImportCotizaciones} />
        )}
      </div>

      <Toast message={toast.message} type={toast.type} />
    </div>
  );
}
