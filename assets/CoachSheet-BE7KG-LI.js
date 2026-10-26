import{a as e,n as t,t as n}from"./jsx-runtime-Wy9s6hX0.js";import{D as r,a as i,d as a,i as o,n as s,r as c,t as l}from"./ui-B9bBRzoM.js";import{t as u}from"./Icon-DSCEE_A-.js";import{d,f as ee,h as f}from"./settings-BvLJRgrv.js";import{t as p}from"./toast-CbIM49A2.js";import{n as te}from"./context-DFwHpQoh.js";import{Dt as ne,Ht as m,Ot as h,ht as re}from"./index-SF452oCg.js";/* empty css                 */import{c as g,d as ie,i as _,l as v,r as y}from"./coachSheet-D_KMxqK8.js";var b=e(t(),1),ae=`/**
 * Hybrid ↔ Excel del entrenador (Hoja de cálculo de Google)
 *
 * Script de Google Apps Script que hace de puente entre la app Hybrid y la hoja del entrenador:
 *  - «tabs»:  pestañas visibles (la última visible es la semana actual)
 *  - «read»:  días de una pestaña con sus ejercicios (para crear las rutinas en Hybrid)
 *  - «write»: escribe lo hecho en un entreno: repeticiones conseguidas, KG (solo si cambia)
 *             y las respuestas de sensaciones del día
 *
 * Solo toca la hoja indicada en SHEET_ID y solo responde si se le manda la clave KEY.
 * Instalación (una vez), en la cuenta de Google que puede editar la hoja:
 *  1. script.google.com → Nuevo proyecto → pegar este archivo entero.
 *  2. Rellenar SHEET_ID (el trozo largo del enlace de la hoja, entre /d/ y /edit) y KEY (inventada, larga).
 *  3. Implementar → Nueva implementación → Tipo: Aplicación web → Ejecutar como: Yo →
 *     Quién tiene acceso: Cualquier usuario → Implementar → Autorizar.
 *  4. Copiar la URL de la aplicación web y pegarla en Hybrid (Ajustes → Excel del entrenador) junto con la clave.
 */

const SHEET_ID = 'PEGA_AQUI_EL_ID_DE_LA_HOJA'
const KEY = 'PEGA_AQUI_UNA_CLAVE_LARGA_INVENTADA'

/* ---------------------------------------------------------------------------------------------- */
/* Entrada                                                                                        */
/* ---------------------------------------------------------------------------------------------- */

function doPost(e) {
  let body
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}')
  } catch (err) {
    return reply({ ok: false, error: 'bad-json' })
  }
  if (!body || body.key !== KEY) return reply({ ok: false, error: 'bad-key' })
  const lock = LockService.getScriptLock()
  lock.waitLock(20000)
  try {
    const book = SpreadsheetApp.openById(SHEET_ID)
    if (body.action === 'tabs') return reply({ ok: true, title: book.getName(), tabs: visibleTabs(book) })
    const sheet = pickTab(book, body.tab)
    if (!sheet) return reply({ ok: false, error: 'no-tab' })
    if (body.action === 'read') return reply({ ok: true, tab: sheet.getName(), days: readDays(sheetGrid(sheet)) })
    if (body.action === 'write') return reply(Object.assign({ ok: true, tab: sheet.getName() }, writeWorkout(sheetGrid(sheet), body)))
    return reply({ ok: false, error: 'bad-action' })
  } catch (err) {
    return reply({ ok: false, error: String(err && err.message ? err.message : err) })
  } finally {
    lock.releaseLock()
  }
}

/** Para comprobar desde el navegador que la URL funciona */
function doGet() {
  return reply({ ok: true, hello: 'Hybrid · Excel del entrenador' })
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)
}

function visibleTabs(book) {
  return book
    .getSheets()
    .filter((s) => !s.isSheetHidden())
    .map((s) => s.getName())
}

/** La pestaña pedida o, si no, la última visible (la semana actual) */
function pickTab(book, name) {
  if (name) {
    const s = book.getSheetByName(name)
    if (s) return s
  }
  const visible = book.getSheets().filter((s) => !s.isSheetHidden())
  return visible[visible.length - 1] || null
}

/** Rejilla de la pestaña: valores y escritura por fila/columna (1 = primera) */
function sheetGrid(sheet) {
  const values = sheet.getDataRange().getValues()
  return {
    values,
    set(row, col, value) {
      sheet.getRange(row, col).setValue(value)
    },
  }
}

/* ---------------------------------------------------------------------------------------------- */
/* Lógica (sin nada de Google: se prueba aparte con los datos de la hoja)                          */
/* ---------------------------------------------------------------------------------------------- */

/** Texto para comparar: sin tildes, en minúsculas y sin espacios de más */
function fold(s) {
  return String(s == null ? '' : s)
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/\\s+/g, ' ')
    .trim()
}

/** Valor de celda como texto. «8-10» escrito a mano lo guarda la hoja como fecha (8 de octubre): se deshace */
function cellText(v) {
  if (v == null) return ''
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return v.getDate() + '-' + (v.getMonth() + 1)
  if (typeof v === 'number') return String(v).replace('.', ',')
  return String(v).trim()
}

const COLS = {
  series: 'series',
  target: 'repeticiones objetivo',
  done: 'repeticiones conseguidas',
  kg: 'kg',
  rir: 'rir',
  rest: 'descanso',
  feel: 'sensacion de fatiga / cansancio',
}

const ANSWERS = [
  ['motivacion', 'motivacion'],
  ['tiempo', 'tiempo de entrenamiento'],
  ['sueno', 'sueno'],
  ['exigencia', 'exigencia'],
  ['comentarios', 'comentarios'],
]

/** Bloques de días: cabecera («Repeticiones conseguidas»), título, columnas y filas de ejercicios */
function findBlocks(values) {
  const heads = []
  values.forEach((row, r) => {
    if (row.some((v) => fold(v) === COLS.done)) heads.push(r)
  })
  return heads.map((h, i) => {
    const row = values[h]
    const col = {}
    row.forEach((v, c) => {
      const f = fold(v)
      for (const k in COLS) if (f === COLS[k] && col[k] == null) col[k] = c
    })
    // Título del día: la fila de encima («LUNES · PIERNA» o «TIPO DE ENTRENAMIENTO: · PIERNA»);
    // si encima no hay nada, el de la propia cabecera («OPCIONAL»)
    const above = h > 0 ? values[h - 1] : []
    const aboveA = cellText(above[0])
    const headA = cellText(row[0])
    const type = aboveA ? cellText(above[2]) : ''
    // Formato antiguo: «TIPO DE ENTRENAMIENTO: · DIA 1 TREN INFERIOR» (el nombre del día es el tipo)
    let label = fold(aboveA) === 'tipo de entrenamiento:' ? type : aboveA || headA
    if (!label) label = type || 'Día ' + (i + 1)
    // Ejercicios: filas seguidas con nombre debajo de la cabecera
    const end = i + 1 < heads.length ? heads[i + 1] - 1 : values.length
    const rows = []
    const data = ['series', 'target', 'done', 'kg', 'rir'].map((k) => col[k]).filter((c) => c != null)
    for (let r = h + 1; r < end; r++) {
      const name = cellText(values[r][0])
      if (!name || /^(comentarios y valoraciones|tipo de entrenamiento)/.test(fold(name))) break
      // Filas sin nada que hacer (p. ej. «DESCANSO OBLIGATORIO» o una indicación) no son ejercicios
      if (!data.some((c) => cellText(values[r][c]))) continue
      rows.push(r)
    }
    return { head: h, label, type, col, rows }
  })
}

/** Días con sus ejercicios (para importar las rutinas) */
function readDays(grid) {
  const v = grid.values
  return findBlocks(v)
    .filter((b) => b.rows.length)
    .map((b) => ({
      label: b.label,
      type: b.type,
      exercises: b.rows.map((r) => ({
        name: cellText(v[r][0]),
        series: b.col.series != null ? cellText(v[r][b.col.series]) : '',
        target: b.col.target != null ? cellText(v[r][b.col.target]) : '',
        kg: b.col.kg != null ? cellText(kgOf(v[r][b.col.kg])) : '',
        rir: b.col.rir != null ? cellText(v[r][b.col.rir]) : '',
        rest: b.col.rest != null ? cellText(v[r][b.col.rest]) : '',
      })),
    }))
}

/** Peso de una celda: «12,5» escrito con punto («12.5») lo guarda la hoja como fecha (12 de mayo) */
function kgOf(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return Number(v.getDate() + '.' + (v.getMonth() + 1))
  return v
}

/** Número con coma decimal («22,5») */
function fmtNum(n) {
  return String(Math.round(n * 100) / 100).replace('.', ',')
}

/**
 * KG: el peso de cada serie, separado por guiones («30-30-30», «30-40-40»; con una sola serie, «30»).
 * Se respeta lo que acompaña al número («25/LADO» → «30-30/LADO», «50 + barra» → «60-60 + barra»).
 * Devuelve el valor a escribir o null si ya pone eso.
 */
function newKg(old, kgs) {
  old = kgOf(old)
  const list = (kgs || []).map((x) => Math.round(Number(x) * 100) / 100).filter((x) => isFinite(x) && x > 0)
  if (!list.length) return null
  const repl = list.map(fmtNum).join('-')
  if (typeof old === 'number') {
    if (list.length === 1 && Math.abs(old - list[0]) < 0.001) return null
    return list.length === 1 ? list[0] : repl
  }
  const text = typeof old === 'string' ? old : ''
  // El número o la lista de pesos que ya hubiera («25-30-30»): se sustituye entera
  const m = text.match(/\\d+(?:[.,]\\d+)?(?:\\s*-\\s*\\d+(?:[.,]\\d+)?)*/)
  if (m) {
    if (m[0].replace(/\\s+/g, '').replace(/\\./g, ',') === repl) return null
    return text.replace(m[0], repl)
  }
  return list.length === 1 ? list[0] : repl
}

/** Lo que se escribe como texto se marca con ' para que la hoja no lo convierta (p. ej. «10-9» en fecha) */
function asCell(x) {
  if (typeof x === 'number') return x
  const s = String(x)
  if (/^-?\\d+(\\.\\d+)?$/.test(s)) return Number(s)
  return "'" + s
}

/**
 * Escribe un entreno: { day, exercises: [{ name, reps, kgs }], answers: { general, motivacion, … } }.
 * El día se busca por los ejercicios (el bloque que tenga más de ellos), por si cambia el título.
 */
function writeWorkout(grid, body) {
  const v = grid.values
  const blocks = findBlocks(v)
  const wanted = (body.exercises || []).map((e) => Object.assign({}, e, { key: fold(e.name) }))
  const nameAt = (r) => fold(cellText(v[r][0]))
  let best = null
  let bestScore = -1
  for (const b of blocks) {
    let score = wanted.filter((e) => b.rows.some((r) => nameAt(r) === e.key)).length * 10
    if (body.day && fold(body.day) === fold(b.label)) score += 1
    if (score > bestScore) {
      best = b
      bestScore = score
    }
  }
  if (!best || bestScore < 10) return { ok: false, error: 'no-day', written: [], missing: wanted.map((e) => e.name) }

  const written = []
  const missing = []
  const used = {}
  for (const e of wanted) {
    const r = best.rows.find((x) => !used[x] && nameAt(x) === e.key)
    if (r == null) {
      missing.push(e.name)
      continue
    }
    used[r] = true
    if (e.reps && best.col.done != null) grid.set(r + 1, best.col.done + 1, asCell(e.reps))
    if (best.col.kg != null) {
      const kg = newKg(v[r][best.col.kg], e.kgs)
      if (kg != null) grid.set(r + 1, best.col.kg + 1, asCell(kg))
    }
    written.push(e.name)
  }

  // Respuestas de sensaciones: a la derecha de cada pregunta (la general, en la propia cabecera)
  const a = body.answers || {}
  const feel = best.col.feel
  if (feel != null) {
    const answerCol = feel + 2
    if (a.general) grid.set(best.head + 1, answerCol, asCell(a.general))
    const last = best.rows.length ? best.rows[best.rows.length - 1] : best.head
    const done = {}
    for (let r = best.head + 1; r <= last + 2 && r < v.length; r++) {
      const label = fold(v[r][feel])
      for (const pair of ANSWERS) {
        const k = pair[0]
        if (!done[k] && label.indexOf(pair[1]) >= 0) {
          done[k] = true
          if (a[k]) grid.set(r + 1, answerCol, asCell(a[k]))
        }
      }
    }
  }
  return { day: best.label, written, missing }
}

/* Para las pruebas fuera de Google */
if (typeof module !== 'undefined') module.exports = { findBlocks, readDays, writeWorkout, newKg, cellText, asCell }
`,x=n(),S=`Entrenador`,C=e=>a(e).replace(/\s+/g,` `).trim(),w=(e,t)=>`${C(e)}|${C(t)}`,T=e=>g(e),E=e=>e.replace(/\bE\.\s*/gi,`elevaciones `).replace(/\bext\.\s*/gi,`extension `);function D(){let e=r(),{exercises:t,findExercise:n,routines:a,folders:D,userId:O,refresh:oe}=te(),k=ee().coachSheet,[A,j]=(0,b.useState)(k?.url??``),[M,N]=(0,b.useState)(k?.key??``),[P,F]=(0,b.useState)(null),[I,L]=(0,b.useState)(null),[R,z]=(0,b.useState)(null),[B,V]=(0,b.useState)(null),[H,U]=(0,b.useState)({}),[W,G]=(0,b.useState)({}),[K,q]=(0,b.useState)(null),[J,Y]=(0,b.useState)(null),X=a.filter(e=>e.user_id===O),Z=(0,b.useMemo)(()=>t.filter(e=>!e.custom),[t]),Q=(0,b.useMemo)(()=>ne(Z),[Z]);(0,b.useEffect)(()=>{if(!k?.url||!k.key)return;let e=!0;return _(k).then(t=>e&&L(t)).catch(()=>{}),()=>{e=!1}},[]);async function se(){let e={url:A.trim(),key:M.trim()};if(!/^https:\/\/script\.google(usercontent)?\.com\//.test(e.url))return p(`Pega la URL de la aplicación web del script (script.google.com/…/exec)`);if(e.key.length<8)return p(`Pega la clave que pusiste en el script`);F(`connect`);try{let t=await _(e);L(t),d({coachSheet:{...e,tab:k?.tab??``}}),p(`Conectado con «${t.title}»`)}catch(e){p(e.message)}finally{F(null)}}async function ce(){F(`read`);try{let[e,t]=await Promise.all([y(),O?re(O):Promise.resolve({})]);if(!e.days.length)return p(`No hay días con ejercicios en esa pestaña`);let r=new Map;for(let e of X)for(let t of e.exercises)t.coach&&n(t.exerciseId)&&r.set(C(t.coach.name),t.exerciseId);let i={};for(let n of e.days)for(let e of n.exercises){let a=r.get(C(e.name))??h(E(e.name),Q.index,Q.libById,Q.libByName,t).exercise?.id;i[w(n.label,e.name)]=a?{id:a}:{create:e.name}}U(i),G({}),V(e)}catch(e){p(e.message)}finally{F(null)}}async function le(){if(!B||!O)return;let t=B.days.filter(e=>!W[e.label]);if(!t.length)return p(`Elige al menos un día`);F(`create`);try{let n=D.find(e=>e.user_id===O&&e.name===S);if(!n){let{data:e,error:t}=await f.from(`routine_folders`).insert({name:S,position:D.filter(e=>e.user_id===O).length}).select(`*`).single();if(t||!e)throw Error(`No se pudo crear la carpeta`);n=e}let r=n.id,i=new Map;for(let e of t)for(let t of e.exercises){let n=H[w(e.label,t.name)];if(!n||!(`create`in n)||i.has(C(t.name)))continue;let r=h(E(t.name),Q.index,Q.libById,Q.libByName).suggestions[0]?.exercise,{data:a,error:o}=await f.from(`custom_exercises`).insert({name:g(t.name).slice(0,80),equipment:r?.equipment??`other`,muscle:r?.muscle??`full_body`,secondary_muscles:[],exercise_type:ie(t.target)}).select(`id`).single();if(o||!a)throw Error(`No se pudo crear «${g(t.name)}»`);i.set(C(t.name),`c:${a.id}`)}let a=Math.max(0,...X.map(e=>e.position+1));for(let e of t){let t=X.find(t=>t.folder_id===r&&t.exercises.some(t=>t.coach&&C(t.coach.day)===C(e.label))),n=e.exercises.map(n=>{let r=H[w(e.label,n.name)],a=r&&`id`in r?r.id:i.get(C(n.name)),o=t?.exercises.find(e=>e.coach&&C(e.coach.name)===C(n.name));return v(n,e.label,a,o?.restSeconds)}),o=(e.type&&C(e.type)!==C(e.label)?`${T(e.type)} · ${T(e.label)}`:T(e.label)).slice(0,80);if((t?await f.from(`routines`).update({title:o,exercises:n}).eq(`id`,t.id):await f.from(`routines`).insert({title:o,exercises:n,folder_id:r,position:a++})).error)throw Error(`No se pudo guardar «${o}»`)}await oe(),p(`Rutinas del entrenador listas (${t.length})`),V(null),e(`/entrenamiento`)}catch(e){p(e.message)}finally{F(null)}}if(K)return(0,x.jsx)(m,{single:!0,onCancel:()=>q(null),onDone:([e])=>{e&&U(t=>({...t,[K.key]:{id:e}})),q(null)}});if(B){let e=B.days.filter(e=>!W[e.label]).length,t=J?H[J]:void 0,r=J?B.days.flatMap(e=>e.exercises.map(t=>[w(e.label,t.name),t.name])).find(([e])=>e===J)?.[1]:``;return(0,x.jsxs)(`div`,{className:`st`,children:[(0,x.jsx)(s,{title:`Importar «${B.tab}»`,left:(0,x.jsx)(`button`,{className:`text-btn`,onClick:()=>V(null),children:`Cancelar`}),right:(0,x.jsx)(`button`,{className:`save-btn`,disabled:!e||P!=null,onClick:le,children:P===`create`?`Creando…`:`Crear`})}),(0,x.jsx)(`div`,{className:`st-body`,children:B.days.map(e=>(0,x.jsxs)(`section`,{className:`st-group ${W[e.label]?`ch-off`:``}`,children:[(0,x.jsxs)(`label`,{className:`ch-day`,children:[(0,x.jsx)(`input`,{type:`checkbox`,checked:!W[e.label],onChange:t=>G(n=>({...n,[e.label]:!t.target.checked}))}),(0,x.jsx)(`b`,{children:T(e.label)}),e.type&&C(e.type)!==C(e.label)&&(0,x.jsx)(`span`,{children:T(e.type)})]}),(0,x.jsx)(`div`,{className:`st-list`,children:e.exercises.map(t=>{let r=w(e.label,t.name),i=H[r],a=i&&`id`in i?n(i.id):void 0;return(0,x.jsxs)(`button`,{className:`st-row ch-row`,onClick:()=>Y(r),children:[a?(0,x.jsx)(l,{exercise:a,size:36}):(0,x.jsx)(`span`,{className:`ch-new`,"aria-hidden":`true`,children:(0,x.jsx)(u,{name:`plus`,size:18})}),(0,x.jsxs)(`span`,{className:`st-text`,children:[(0,x.jsx)(`b`,{children:a?a.name:`${g(t.name)} (nuevo)`}),(0,x.jsxs)(`small`,{children:[g(t.name),` · `,t.series||`?`,` × `,t.target||`?`,t.kg?` · ${t.kg}`:``]})]}),(0,x.jsx)(u,{name:`chevronRight`,size:18,className:`st-chev`})]},r)})})]},e.label))}),(0,x.jsx)(c,{open:!!J,onClose:()=>Y(null),title:r?g(r):``,children:J&&(0,x.jsxs)(x.Fragment,{children:[h(E(r??``),Q.index,Q.libById,Q.libByName).suggestions.filter((e,t,n)=>n.findIndex(t=>t.exercise.name===e.exercise.name)===t).slice(0,3).map(({exercise:e})=>(0,x.jsx)(o,{icon:(0,x.jsx)(l,{exercise:e,size:36}),selected:!!t&&`id`in t&&t.id===e.id,onClick:()=>(U(t=>({...t,[J]:{id:e.id}})),Y(null)),children:e.name},e.id)),(0,x.jsx)(o,{icon:(0,x.jsx)(u,{name:`search`,size:20}),onClick:()=>(q({key:J,name:r??``}),Y(null)),children:`Elegir otro ejercicio`}),(0,x.jsxs)(o,{icon:(0,x.jsx)(u,{name:`plus`,size:20}),selected:!!t&&`create`in t,onClick:()=>(U(e=>({...e,[J]:{create:r??``}})),Y(null)),children:[`Crear «`,g(r??``),`» como personalizado`]})]})})]})}let $=I?.tabs.at(-1);return(0,x.jsxs)(`div`,{className:`st`,children:[(0,x.jsx)(s,{title:`Excel del entrenador`}),(0,x.jsxs)(`div`,{className:`st-body`,children:[(0,x.jsxs)(`section`,{className:`st-group`,children:[(0,x.jsx)(`h2`,{className:`st-title`,children:`Conexión`}),(0,x.jsx)(`div`,{className:`st-list`,children:(0,x.jsxs)(`div`,{className:`st-row st-row-col`,children:[(0,x.jsx)(`label`,{className:`ch-label`,htmlFor:`coach-url`,children:`URL del script`}),(0,x.jsx)(`input`,{id:`coach-url`,className:`sheet-input ch-input`,placeholder:`https://script.google.com/macros/s/…/exec`,value:A,onChange:e=>j(e.target.value),autoComplete:`off`,inputMode:`url`}),(0,x.jsx)(`label`,{className:`ch-label`,htmlFor:`coach-key`,children:`Clave`}),(0,x.jsx)(`input`,{id:`coach-key`,className:`sheet-input ch-input`,placeholder:`La que pusiste en el script`,value:M,onChange:e=>N(e.target.value),autoComplete:`off`,type:`password`}),(0,x.jsx)(`button`,{className:`btn btn-primary btn-block`,disabled:P!=null,onClick:se,children:P===`connect`?`Conectando…`:I?`Volver a conectar`:`Conectar`})]})}),(0,x.jsx)(`p`,{className:`st-note`,children:I?(0,x.jsxs)(x.Fragment,{children:[(0,x.jsx)(u,{name:`check`,size:14,stroke:2.6}),` Conectado con «`,I.title,`».`]}):(0,x.jsx)(`button`,{className:`text-btn ch-help`,onClick:()=>z(`help`),children:`¿Cómo se consigue la URL y la clave?`})})]}),I&&(0,x.jsxs)(`section`,{className:`st-group`,children:[(0,x.jsx)(`h2`,{className:`st-title`,children:`Semana`}),(0,x.jsxs)(`div`,{className:`st-list`,children:[(0,x.jsxs)(`button`,{className:`st-row`,onClick:()=>z(`tab`),children:[(0,x.jsx)(`span`,{className:`st-icon`,children:(0,x.jsx)(u,{name:`calendar`,size:19})}),(0,x.jsxs)(`span`,{className:`st-text`,children:[(0,x.jsx)(`b`,{children:`Pestaña de esta semana`}),(0,x.jsx)(`small`,{children:k?.tab?`Elegida a mano`:`Automática: la última pestaña visible`})]}),(0,x.jsx)(`span`,{className:`st-value`,children:k?.tab||$}),(0,x.jsx)(u,{name:`chevronRight`,size:18,className:`st-chev`})]}),(0,x.jsxs)(`button`,{className:`st-row`,disabled:P!=null,onClick:ce,children:[(0,x.jsx)(`span`,{className:`st-icon`,children:(0,x.jsx)(u,{name:`arrowDown`,size:19})}),(0,x.jsxs)(`span`,{className:`st-text`,children:[(0,x.jsx)(`b`,{children:P===`read`?`Leyendo el Excel…`:`Importar rutinas de esta semana`}),(0,x.jsxs)(`small`,{children:[`Crea o actualiza las rutinas de la carpeta «`,S,`»`]})]}),P===`read`?(0,x.jsx)(i,{}):(0,x.jsx)(u,{name:`chevronRight`,size:18,className:`st-chev`})]})]})]}),I&&(0,x.jsx)(`section`,{className:`st-group`,children:(0,x.jsx)(`div`,{className:`st-list`,children:(0,x.jsx)(`button`,{className:`st-row danger`,onClick:()=>{d({coachSheet:null}),L(null),j(``),N(``)},children:`Desconectar`})})})]}),(0,x.jsxs)(c,{open:R===`tab`,onClose:()=>z(null),title:`Pestaña de esta semana`,children:[(0,x.jsxs)(o,{selected:!k?.tab,onClick:()=>(k&&d({coachSheet:{...k,tab:``}}),z(null)),children:[`Automática (la última visible: `,$,`)`]}),I?.tabs.map(e=>(0,x.jsx)(o,{selected:k?.tab===e,onClick:()=>(k&&d({coachSheet:{...k,tab:e}}),z(null)),children:e},e))]}),(0,x.jsxs)(c,{open:R===`help`,onClose:()=>z(null),title:`Conectar con la Hoja de Google`,children:[(0,x.jsxs)(`ol`,{className:`ch-steps`,children:[(0,x.jsxs)(`li`,{children:[`En el ordenador, entra en `,(0,x.jsx)(`b`,{children:`script.google.com`}),` con la cuenta de Google que puede editar el Excel y crea un `,(0,x.jsx)(`b`,{children:`Nuevo proyecto`}),`.`]}),(0,x.jsx)(`li`,{children:`Borra lo que haya y pega el código del script (botón de abajo).`}),(0,x.jsxs)(`li`,{children:[`Arriba del código, en `,(0,x.jsx)(`b`,{children:`SHEET_ID`}),` pega el trozo largo del enlace del Excel (entre `,(0,x.jsx)(`code`,{children:`/d/`}),` y `,(0,x.jsx)(`code`,{children:`/edit`}),`) y en`,` `,(0,x.jsx)(`b`,{children:`KEY`}),` inventa una clave larga.`]}),(0,x.jsxs)(`li`,{children:[(0,x.jsx)(`b`,{children:`Implementar → Nueva implementación →`}),` tipo `,(0,x.jsx)(`b`,{children:`Aplicación web`}),`, ejecutar como `,(0,x.jsx)(`b`,{children:`Yo`}),`, acceso `,(0,x.jsx)(`b`,{children:`Cualquier usuario`}),`. Autoriza los permisos.`]}),(0,x.jsxs)(`li`,{children:[`Copia la URL que termina en `,(0,x.jsx)(`b`,{children:`/exec`}),` y pégala aquí con la clave.`]})]}),(0,x.jsxs)(`button`,{className:`btn btn-primary btn-block`,onClick:async()=>{try{await navigator.clipboard.writeText(ae),p(`Código del script copiado`)}catch{p(`No se pudo copiar`)}},children:[(0,x.jsx)(u,{name:`copy`,size:18}),`Copiar el código del script`]})]})]})}export{D as CoachSheet};