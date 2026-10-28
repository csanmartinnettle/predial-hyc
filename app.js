/* Administración Predial - H&C
   App web instalable (PWA) para reportes de visitas a terrenos.
   Datos: Firebase (Auth anónima + Firestore con caché sin conexión). */

import { firebaseConfig } from "./firebase-config.js";
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getAuth, signInAnonymously, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager,
  collection, doc, getDoc, getDocs, setDoc, updateDoc, writeBatch, onSnapshot, query, orderBy, limit
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

/* ---------- constantes ---------- */
const MAX_FOTOS = 8;
const TERRENOS_INICIALES = ["Los Andes", "Suizos", "Reina Norte", "Batuco", "Lote G", "Talagante"];
const PREGUNTAS_INICIALES = [
  { id: "cierre", texto: "Estado del cierre perimetral", tipo: "scale", lo: "Muy malo", hi: "Muy bueno", req: true },
  { id: "personas", texto: "Presencia de personas ajenas al terreno", tipo: "yn", req: true },
  { id: "vandalismo", texto: "Vandalismo en terreno", tipo: "scale", lo: "Grave", hi: "Sin vandalismo", req: true },
  { id: "animales", texto: "Presencia de animales no deseados", tipo: "scale", lo: "Muchos", hi: "Ninguno", req: true }
].map((q, i) => ({ ...q, orden: i + 1, activa: true }));
const SCALE_TXT = { 1: "Muy malo", 2: "Malo", 3: "Regular", 4: "Bueno", 5: "Muy bueno" };

const $ = id => document.getElementById(id);
const S = {
  uid: null, me: null, terrenos: [], reportes: [], preguntas: [], preguntasLive: false,
  usuarios: [], cur: null, draft: null, unsubUsers: null, share: null
};

/* ---------- utilidades ---------- */
function toast(t) { const el = $("toast"); el.textContent = t; el.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => el.hidden = true, 3500); }
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
function slug(s) { return s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40) || uid(); }
const fmtDate = iso => new Date(iso).toLocaleDateString("es-CL", { day: "numeric", month: "short", year: "numeric" });
const fmtDT = iso => new Date(iso).toLocaleString("es-CL", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
function repName(r) { const d = new Date(r.fecha); return `${d.getFullYear()} ${String(d.getMonth() + 1).padStart(2, "0")} ${String(d.getDate()).padStart(2, "0")} ${(r.terrenoNombre || "").toUpperCase()}`; }
function ansTxt(q, v) {
  if (v == null || v === "") return "—";
  if (q.tipo === "yn") return v === "si" ? "Sí" : "No";
  if (q.tipo === "text") return String(v);
  return v + " / 5";
}
function ansLong(q, v) {
  let a = ansTxt(q, v);
  if (q.tipo === "scale" && v) a += ` (${v === 1 ? (q.lo || SCALE_TXT[1]) : v === 5 ? (q.hi || SCALE_TXT[5]) : SCALE_TXT[v]})`;
  return a;
}
const isAdmin = () => S.me && S.me.rol === "admin";
const activeQuestions = () => S.preguntas.filter(q => q.activa !== false);
const myName = () => (S.me && S.me.nombre) || "";

async function compress(file, max = 1280, q = 0.72) {
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = url; });
    let sc = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
    for (let tries = 0; tries < 4; tries++) {
      const w = Math.round(img.naturalWidth * sc), h = Math.round(img.naturalHeight * sc);
      const c = document.createElement("canvas"); c.width = w; c.height = h; c.getContext("2d").drawImage(img, 0, 0, w, h);
      const data = c.toDataURL("image/jpeg", q);
      if (data.length < 900000) return { data, w, h };   // cada foto debe caber en un documento de Firestore (1 MB)
      sc *= 0.8;
    }
    throw new Error("Foto demasiado grande");
  } finally { URL.revokeObjectURL(url); }
}

/* ---------- borradores en el teléfono (IndexedDB) ---------- */
const mem = {};
let idbP = null;
function idb() {
  if (idbP) return idbP;
  idbP = new Promise(res => {
    try {
      const r = indexedDB.open("hyc-predial", 1);
      r.onupgradeneeded = () => r.result.createObjectStore("drafts");
      r.onsuccess = () => res(r.result); r.onerror = () => res(null);
    } catch (e) { res(null); }
  });
  return idbP;
}
async function draftGet(k) {
  const d = await idb(); if (!d) return mem[k];
  return new Promise(res => { try { const r = d.transaction("drafts").objectStore("drafts").get(k); r.onsuccess = () => res(r.result); r.onerror = () => res(undefined); } catch (e) { res(undefined); } });
}
async function draftPut(k, v) {
  const d = await idb(); if (!d) { mem[k] = v; return; }
  return new Promise(res => { try { const tx = d.transaction("drafts", "readwrite"); tx.objectStore("drafts").put(v, k); tx.oncomplete = res; tx.onerror = res; } catch (e) { res(); } });
}
async function draftDel(k) {
  const d = await idb(); if (!d) { delete mem[k]; return; }
  return new Promise(res => { try { const tx = d.transaction("drafts", "readwrite"); tx.objectStore("drafts").delete(k); tx.oncomplete = res; tx.onerror = res; } catch (e) { res(); } });
}

/* ---------- Firebase ---------- */
if (!firebaseConfig || !firebaseConfig.apiKey || firebaseConfig.apiKey.startsWith("PEGAR")) {
  document.body.innerHTML = '<main class="wrap"><h1>Falta configurar</h1><p class="sub">Copia los datos de tu proyecto Firebase en el archivo <b>firebase-config.js</b>, como indica la guía de instalación.</p></main>';
  throw new Error("firebase-config.js sin completar");
}
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = initializeFirestore(app, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });

onAuthStateChanged(auth, async user => {
  if (!user) {
    try { await signInAnonymously(auth); }
    catch (e) { showFatal("No se pudo iniciar la app. Revisa la conexión a internet y vuelve a abrirla."); }
    return;
  }
  S.uid = user.uid;
  let snap = null;
  try { snap = await getDoc(doc(db, "usuarios", user.uid)); } catch (e) { /* sin conexión y sin caché */ }
  if (snap && snap.exists()) { S.me = snap.data(); start(); }
  else { showOnboard(true); }
});

function showFatal(msg) {
  $("vHome").hidden = false; $("lots").innerHTML = ""; const d = document.createElement("div"); d.className = "notice err"; d.textContent = msg; $("lots").append(d);
}

/* ---------- bienvenida: nombre + código ---------- */
function showOnboard(isNew) {
  $("codeField").hidden = !isNew; $("accessCode").required = isNew;
  $("userName").value = myName(); $("nameErr").textContent = "";
  $("nameBtn").textContent = isNew ? "Entrar" : "Guardar nombre";
  $("onboard").hidden = false; setTimeout(() => $("userName").focus(), 60);
}
$("nameForm").onsubmit = async e => {
  e.preventDefault();
  const nombre = $("userName").value.trim().replace(/\s+/g, " ");
  if (nombre.split(" ").length < 2) { $("nameErr").textContent = "Escribe nombre y apellido."; return; }
  $("nameBtn").disabled = true; $("nameErr").textContent = "";
  try {
    if (!S.me) {
      const codigo = $("accessCode").value.trim();
      const data = { nombre, codigo, rol: "usuario", creado: new Date().toISOString() };
      await setDoc(doc(db, "usuarios", S.uid), data);
      S.me = data; $("onboard").hidden = true; toast("Hola, " + nombre.split(" ")[0] + "."); start();
    } else {
      await updateDoc(doc(db, "usuarios", S.uid), { nombre });
      S.me.nombre = nombre; showWho(); $("onboard").hidden = true; toast("Nombre actualizado.");
    }
  } catch (err) {
    $("nameErr").textContent = (err && err.code === "permission-denied")
      ? "El código de acceso no es correcto. Pídeselo al administrador."
      : "No se pudo guardar. Revisa la conexión e inténtalo de nuevo.";
  } finally { $("nameBtn").disabled = false; }
};
$("changeName").onclick = () => showOnboard(false);
function showWho() { $("whoName").textContent = myName() || "—"; $("adminBtn").hidden = !isAdmin(); }

/* ---------- arranque con sesión ---------- */
let started = false;
function start() {
  if (started) return; started = true;
  showWho();
  onSnapshot(doc(db, "usuarios", S.uid), s => { if (s.exists()) { S.me = s.data(); showWho(); } });

  onSnapshot(collection(db, "terrenos"), s => {
    S.terrenos = s.docs.map(d => ({ id: d.id, ...d.data() })).filter(t => !t.archivado).sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
    if (s.empty && !s.metadata.fromCache) seedTerrenos();
    route();
  }, () => showFatal("No tienes acceso a los datos. Pide al administrador que revise tu usuario."));

  onSnapshot(query(collection(db, "preguntas"), orderBy("orden")), s => {
    S.preguntasLive = !s.empty;
    S.preguntas = s.empty ? PREGUNTAS_INICIALES.slice() : s.docs.map(d => ({ id: d.id, ...d.data() }));
    if (s.empty && !s.metadata.fromCache && isAdmin()) seedPreguntas();
    if (S.cur && !$("pVisit").hidden) renderSurvey();
    if (!$("vAdmin").hidden) renderQAdmin();
  });

  onSnapshot(query(collection(db, "reportes"), orderBy("fecha", "desc"), limit(1000)), { includeMetadataChanges: true }, s => {
    S.reportes = s.docs.map(d => ({ id: d.id, ...d.data(), _pend: d.metadata.hasPendingWrites }));
    refresh();
  });
  route();
}
async function seedTerrenos() {
  for (const n of TERRENOS_INICIALES) { try { await setDoc(doc(db, "terrenos", slug(n)), { nombre: n, creado: new Date().toISOString() }); } catch (e) { } }
}
async function seedPreguntas() {
  try { const b = writeBatch(db); for (const q of PREGUNTAS_INICIALES) { const { id, ...rest } = q; b.set(doc(db, "preguntas", id), rest); } await b.commit(); } catch (e) { }
}

/* ---------- navegación ---------- */
function route() {
  if (!started) return;
  const h = location.hash.replace("#", "");
  if (h === "admin" && isAdmin()) return showAdmin();
  const id = h.startsWith("t-") ? h.slice(2) : null;
  const lot = id && S.terrenos.find(t => t.id === id);
  if (lot) { if (!S.cur || S.cur.id !== lot.id) openLot(lot); else refresh(); }
  else showHome();
}
addEventListener("hashchange", route);
$("back").onclick = () => { history.length > 1 ? history.back() : (location.hash = ""); };
$("adminBtn").onclick = () => { location.hash = "admin"; };

function setView(v) {
  for (const id of ["vHome", "vLot", "vAdmin"]) $(id).hidden = id !== v;
  $("back").hidden = v === "vHome"; $("logoImg").hidden = v !== "vHome";
  $("actionbar").hidden = !(v === "vLot" && !$("pVisit").hidden);
}
function showHome() {
  S.cur = null; setView("vHome"); $("brand").textContent = "Administración Predial"; renderHome();
}

/* ---------- inicio ---------- */
function renderHome() {
  const box = $("lots"); box.textContent = "";
  if (!S.terrenos.length) { const d = document.createElement("div"); d.className = "empty"; d.textContent = "Cargando terrenos…"; box.append(d); }
  for (const t of S.terrenos) {
    const reps = S.reportes.filter(r => r.terrenoId === t.id);
    const pend = reps.filter(r => r._pend).length;
    const b = document.createElement("button"); b.className = "lot"; b.type = "button";
    b.innerHTML = '<span class="pin"></span><span><div class="lot-name"></div><div class="lot-meta"></div></span><span class="chev">›</span>';
    b.querySelector(".pin").textContent = t.nombre.trim().charAt(0).toUpperCase();
    b.querySelector(".lot-name").textContent = t.nombre;
    const meta = b.querySelector(".lot-meta");
    meta.textContent = reps.length ? `Última visita ${fmtDate(reps[0].fecha)} · ${reps.length} reporte${reps.length > 1 ? "s" : ""}` : "Sin visitas registradas";
    if (pend) { const s = document.createElement("span"); s.className = "pend"; s.textContent = `${pend} por enviar`; meta.append(" ", s); }
    b.onclick = () => { location.hash = "t-" + t.id; };
    box.append(b);
  }
  const pend = S.reportes.filter(r => r._pend).length;
  $("pendNote").hidden = !pend;
  $("pendNote").textContent = pend ? `${pend} reporte${pend > 1 ? "s" : ""} guardado${pend > 1 ? "s" : ""} en este teléfono. Se enviará${pend > 1 ? "n" : ""} solo${pend > 1 ? "s" : ""} cuando haya señal.` : "";
}
$("addForm").onsubmit = async e => {
  e.preventDefault(); const name = $("newLot").value.trim(); if (!name) return;
  if (S.terrenos.some(t => t.nombre.toLowerCase() === name.toLowerCase())) { toast("Ese terreno ya existe."); return; }
  let id = slug(name); if (S.terrenos.some(t => t.id === id)) id += "-" + uid().slice(-4);
  setDoc(doc(db, "terrenos", id), { nombre: name, creado: new Date().toISOString(), creadoPor: S.uid }).catch(() => toast("No se pudo agregar el terreno."));
  $("newLot").value = ""; toast("Terreno agregado: " + name);
};

/* ---------- terreno ---------- */
async function openLot(lot) {
  S.cur = lot; setView("vLot"); $("brand").textContent = lot.nombre; $("lotName").textContent = lot.nombre;
  S.draft = (await draftGet(lot.id)) || { respuestas: {}, obs: "", fotos: [], inicio: new Date().toISOString() };
  if (S.cur !== lot) return;
  $("obs").value = S.draft.obs || "";
  renderSurvey(); renderThumbs(); setTab("visit"); resetConfirm(); refresh(); scrollTo(0, 0);
}
function setTab(t) {
  $("tabVisit").setAttribute("aria-selected", t === "visit"); $("tabReps").setAttribute("aria-selected", t === "reps");
  $("pVisit").hidden = t !== "visit"; $("pReps").hidden = t !== "reps"; $("actionbar").hidden = t !== "visit";
  if (t === "reps") renderReps();
}
$("tabVisit").onclick = () => setTab("visit"); $("tabReps").onclick = () => setTab("reps");

let saveT;
function saveDraft() {
  clearTimeout(saveT);
  saveT = setTimeout(async () => {
    if (!S.cur) return; await draftPut(S.cur.id, S.draft);
    $("savedTxt").textContent = "Borrador guardado en este teléfono · " + new Date().toLocaleTimeString("es-CL", { hour: "2-digit", minute: "2-digit" });
  }, 300);
}

function renderSurvey() {
  const box = $("survey"); box.textContent = "";
  activeQuestions().forEach((q, i) => {
    const d = document.createElement("div"); d.className = "q"; d.dataset.k = q.id;
    const t = document.createElement("div"); t.className = "q-t"; t.innerHTML = "<span></span><div></div>";
    t.firstChild.textContent = (i + 1) + "."; t.lastChild.textContent = q.texto + (q.req === false ? " (opcional)" : ""); d.append(t);
    if (q.tipo === "text") {
      const ta = document.createElement("textarea"); ta.id = "q_" + q.id; ta.style.minHeight = "70px";
      ta.value = S.draft.respuestas[q.id] || "";
      ta.oninput = () => { S.draft.respuestas[q.id] = ta.value; d.classList.remove("missing"); saveDraft(); };
      d.append(ta);
    } else {
      const sc = document.createElement("div"); sc.className = "scale" + (q.tipo === "yn" ? " yn" : ""); sc.setAttribute("role", "radiogroup"); sc.setAttribute("aria-label", q.texto);
      const opts = q.tipo === "yn" ? [["si", "Sí"], ["no", "No"]] : [1, 2, 3, 4, 5].map(n => [String(n), String(n)]);
      for (const [v, l] of opts) {
        const lab = document.createElement("label"); const inp = document.createElement("input");
        inp.type = "radio"; inp.name = "q_" + q.id; inp.id = "q_" + q.id + "_" + v; inp.value = v;
        inp.checked = String(S.draft.respuestas[q.id] ?? "") === v;
        inp.onchange = () => { S.draft.respuestas[q.id] = q.tipo === "yn" ? v : Number(v); d.classList.remove("missing"); saveDraft(); };
        const s = document.createElement("span"); s.className = "opt"; s.textContent = l; lab.append(inp, s); sc.append(lab);
      }
      d.append(sc);
      if (q.tipo === "scale") {
        const e = document.createElement("div"); e.className = "ends"; e.innerHTML = "<span></span><span></span>";
        e.firstChild.textContent = "1 · " + (q.lo || SCALE_TXT[1]); e.lastChild.textContent = (q.hi || SCALE_TXT[5]) + " · 5"; d.append(e);
      }
    }
    box.append(d);
  });
}
$("obs").oninput = () => { S.draft.obs = $("obs").value; saveDraft(); };

function renderThumbs() {
  const box = $("thumbs"); box.textContent = "";
  S.draft.fotos.forEach((f, i) => {
    const d = document.createElement("div"); d.className = "thumb";
    const img = document.createElement("img"); img.src = f.data; img.alt = "Foto " + (i + 1);
    const x = document.createElement("button"); x.type = "button"; x.textContent = "×"; x.setAttribute("aria-label", "Quitar foto " + (i + 1));
    x.onclick = () => { S.draft.fotos.splice(i, 1); renderThumbs(); saveDraft(); };
    d.append(img, x); box.append(d);
  });
  const n = S.draft.fotos.length;
  $("photoHint").textContent = n ? `${n} de ${MAX_FOTOS} fotos` + (n >= MAX_FOTOS ? " · llegaste al máximo." : ".") : `Sin fotos todavía. Máximo ${MAX_FOTOS}.`;
  const full = n >= MAX_FOTOS;
  for (const id of ["cam", "gal"]) { $(id).disabled = full; $(id).nextElementSibling.style.opacity = full ? .45 : ""; }
}
async function addFiles(files) {
  const room = MAX_FOTOS - S.draft.fotos.length;
  if (room <= 0) { toast(`Máximo ${MAX_FOTOS} fotos por visita.`); return; }
  if (files.length > room) { toast(`Solo se agregaron ${room} foto${room > 1 ? "s" : ""}: el máximo es ${MAX_FOTOS}.`); files = files.slice(0, room); }
  for (const f of files) { try { S.draft.fotos.push(await compress(f)); } catch (e) { toast("No se pudo leer una de las fotos."); } }
  renderThumbs(); saveDraft();
}
$("cam").onchange = e => { addFiles([...e.target.files]); e.target.value = ""; };
$("gal").onchange = e => { addFiles([...e.target.files]); e.target.value = ""; };

/* ---------- finalizar y enviar ---------- */
function resetConfirm() { $("finish").hidden = false; $("confirm").hidden = true; }
$("finish").onclick = () => {
  const miss = activeQuestions().filter(q => q.req !== false && (S.draft.respuestas[q.id] == null || S.draft.respuestas[q.id] === ""));
  document.querySelectorAll(".q").forEach(el => el.classList.toggle("missing", miss.some(q => q.id === el.dataset.k)));
  if (miss.length) {
    toast(`Falta responder ${miss.length} pregunta${miss.length > 1 ? "s" : ""} del reporte.`);
    document.querySelector(".q.missing")?.scrollIntoView({ behavior: "smooth", block: "center" }); return;
  }
  $("finish").hidden = true; $("confirm").hidden = false;
};
$("cancelSend").onclick = resetConfirm;
$("doSend").onclick = async () => {
  const lot = S.cur, d = S.draft, id = uid();
  const preguntas = activeQuestions().map(q => ({ id: q.id, texto: q.texto, tipo: q.tipo, lo: q.lo || "", hi: q.hi || "" }));
  const respuestas = {}; for (const q of preguntas) if (d.respuestas[q.id] != null && d.respuestas[q.id] !== "") respuestas[q.id] = d.respuestas[q.id];
  const rep = {
    terrenoId: lot.id, terrenoNombre: lot.nombre, fecha: new Date().toISOString(), inicio: d.inicio,
    autorUid: S.uid, autorNombre: myName(), preguntas, respuestas, obs: (d.obs || "").trim(),
    nFotos: d.fotos.length
  };
  const b = writeBatch(db);
  b.set(doc(db, "reportes", id), rep);
  d.fotos.forEach((f, i) => b.set(doc(db, "reportes", id, "fotos", String(i).padStart(2, "0")), { i, data: f.data, w: f.w, h: f.h, autorUid: S.uid }));
  // Firestore guarda el envío en el teléfono y lo sube solo cuando hay señal, aunque se cierre la app.
  b.commit().then(() => toast("Reporte enviado: " + repName(rep))).catch(err => {
    console.error(err); toast("El servidor rechazó el reporte. Avisa al administrador.");
  });
  await draftDel(lot.id);
  S.draft = { respuestas: {}, obs: "", fotos: [], inicio: new Date().toISOString() };
  $("obs").value = ""; renderSurvey(); renderThumbs(); resetConfirm();
  $("savedTxt").textContent = "El borrador se guarda en este teléfono.";
  toast(navigator.onLine ? "Enviando reporte…" : "Sin señal: el reporte quedó guardado y se enviará solo.");
  setTab("reps");
};

/* ---------- reportes ---------- */
function refresh() {
  if (S.cur) { if (!$("pReps").hidden) renderReps(); updateTabCount(); }
  else if (!$("vHome").hidden) renderHome();
}
function updateTabCount() {
  if (!S.cur) return; const n = S.reportes.filter(r => r.terrenoId === S.cur.id).length;
  $("tabReps").textContent = n ? `Reportes (${n})` : "Reportes";
  const reps = S.reportes.filter(r => r.terrenoId === S.cur.id);
  $("lotSub").textContent = reps.length ? `Última visita: ${fmtDate(reps[0].fecha)}` : "Sin visitas registradas";
}
function renderReps() {
  const box = $("pReps"); box.textContent = ""; if (!S.cur) return;
  const reps = S.reportes.filter(r => r.terrenoId === S.cur.id);
  if (!reps.length) { const e = document.createElement("div"); e.className = "empty"; e.textContent = "Todavía no hay reportes de este terreno."; box.append(e); return; }
  for (const r of reps) {
    const c = document.createElement("div"); c.className = "card rep";
    c.innerHTML = '<div class="rep-h"><div><div class="rep-date"></div><div class="rep-by"></div></div></div><div class="chips"></div>';
    c.querySelector(".rep-date").textContent = repName(r);
    c.querySelector(".rep-by").textContent = [fmtDT(r.fecha), r.autorNombre && ("Por " + r.autorNombre), (r.nFotos || 0) + " foto" + (r.nFotos === 1 ? "" : "s")].filter(Boolean).join(" · ");
    const ch = c.querySelector(".chips");
    for (const q of (r.preguntas || [])) {
      if (q.tipo === "text") continue;
      const v = r.respuestas?.[q.id]; const s = document.createElement("span"); s.className = "chip";
      if ((q.tipo === "scale" && v <= 2) || (q.tipo === "yn" && v === "si")) s.classList.add("low");
      s.textContent = (q.texto.length > 26 ? q.texto.slice(0, 24) + "…" : q.texto) + ": ";
      const b = document.createElement("b"); b.textContent = ansTxt(q, v); s.append(b); ch.append(s);
    }
    if (r.obs) { const p = document.createElement("p"); p.className = "rep-obs"; p.textContent = r.obs; c.append(p); }
    if (r._pend) { const t = document.createElement("span"); t.className = "tag"; t.textContent = navigator.onLine ? "Enviando…" : "Pendiente de envío"; c.querySelector(".rep-h").append(t); }
    const row = document.createElement("div"); row.className = "share"; row.style.gridTemplateColumns = "1fr";
    const btn = document.createElement("button"); btn.className = "btn"; btn.textContent = "Compartir / descargar PDF";
    btn.onclick = () => prepareShare(r, btn); row.append(btn); c.append(row);
    box.append(c);
  }
}

/* ---------- PDF ---------- */
async function buildPdf(r) {
  if (!window.jspdf) throw new Error("jspdf");
  const { jsPDF } = window.jspdf; const doc_ = new jsPDF({ unit: "mm", format: "a4" });
  const W = 210, M = 16; let y;
  const navy = [27, 34, 61], grey = [103, 103, 103], gold = [181, 128, 106];
  const pdf = doc_;
  pdf.setFillColor(...navy); pdf.rect(0, 0, W, 30, "F"); pdf.setFillColor(...gold); pdf.rect(0, 30, W, 1.2, "F");
  try {
    const li = $("logoImg"), ratio = li.naturalHeight / li.naturalWidth, lc = document.createElement("canvas");
    lc.width = 300; lc.height = Math.round(300 * ratio); lc.getContext("2d").drawImage(li, 0, 0, lc.width, lc.height);
    const lw = 26, lh = lw * ratio; pdf.addImage(lc.toDataURL("image/png"), "PNG", W - M - lw, (30 - lh) / 2, lw, lh);
  } catch (e) { }
  pdf.setTextColor(255, 255, 255); pdf.setFont("helvetica", "bold"); pdf.setFontSize(9);
  pdf.text("H&C  ·  ADMINISTRACIÓN PREDIAL  ·  REPORTE DE VISITA", M, 11);
  pdf.setFontSize(20); pdf.text(r.terrenoNombre, M, 22);
  y = 40; pdf.setTextColor(...navy); pdf.setFontSize(10);
  const row = (k, v) => { pdf.setFont("helvetica", "bold"); pdf.setTextColor(...grey); pdf.text(k, M, y); pdf.setFont("helvetica", "normal"); pdf.setTextColor(...navy); pdf.text(v, M + 38, y); y += 6; };
  row("Fecha de envío", fmtDT(r.fecha)); row("Generado por", r.autorNombre || "—"); row("Fotos", String(r.nFotos || 0));
  y += 4; pdf.setDrawColor(217, 221, 211); pdf.line(M, y, W - M, y); y += 9;
  pdf.setFont("helvetica", "bold"); pdf.setFontSize(12); pdf.text("Reporte", M, y); y += 7; pdf.setFontSize(10);
  (r.preguntas || []).forEach((q, i) => {
    const v = r.respuestas?.[q.id];
    if (y > 270) { pdf.addPage(); y = M; }
    pdf.setFont("helvetica", "normal"); pdf.setTextColor(...navy);
    if (q.tipo === "text") {
      pdf.text(`${i + 1}. ${q.texto}`, M, y); y += 5; pdf.setFont("helvetica", "bold");
      for (const l of pdf.splitTextToSize(ansTxt(q, v), W - 2 * M - 6)) { pdf.text(l, M + 6, y); y += 5; }
      y -= 1;
    } else {
      const qt = pdf.splitTextToSize(`${i + 1}. ${q.texto}`, 120); pdf.text(qt, M, y);
      pdf.setFont("helvetica", "bold"); pdf.text(ansLong(q, v), W - M, y, { align: "right" }); y += 5 * (qt.length - 1);
    }
    y += 4; pdf.setDrawColor(235, 237, 231); pdf.line(M, y, W - M, y); y += 6;
  });
  y += 4; if (y > 260) { pdf.addPage(); y = M; }
  pdf.setFontSize(12); pdf.setFont("helvetica", "bold"); pdf.text("Observaciones", M, y); y += 7;
  pdf.setFont("helvetica", "normal"); pdf.setFontSize(10);
  for (const l of pdf.splitTextToSize(r.obs || "Sin observaciones.", W - 2 * M)) { if (y > 282) { pdf.addPage(); y = M; } pdf.text(l, M, y); y += 5; }
  y += 10; if (y > 255) { pdf.addPage(); y = M + 10; }
  pdf.setDrawColor(...gold); pdf.setLineWidth(0.4); pdf.line(M, y + 12, M + 70, y + 12); pdf.setLineWidth(0.2);
  pdf.setFont("helvetica", "bold"); pdf.setFontSize(10); pdf.setTextColor(...navy); pdf.text(r.autorNombre || "—", M, y + 17);
  pdf.setFont("helvetica", "normal"); pdf.setFontSize(8); pdf.setTextColor(...grey); pdf.text("Generado y enviado desde la app el " + fmtDT(r.fecha), M, y + 21);
  pdf.setTextColor(...navy);

  if (r.nFotos) {
    const fs = await getDocs(collection(db, "reportes", r.id, "fotos"));
    const fotos = fs.docs.map(d => d.data()).sort((a, b) => a.i - b.i);
    // Registro fotográfico: hoja aparte, todas del mismo tamaño (3:2), máximo 8 por página (2 x 4)
    const PER = 8, COLS = 2, GAP = 6, ROWS = 4, CH = (286 - (M + 13)) / ROWS, IH = CH - 9, CW = IH * 1.5, X0 = (W - (COLS * CW + GAP)) / 2;
    const crop = async data => {
      const im = await new Promise((ok, no) => { const i = new Image(); i.onload = () => ok(i); i.onerror = no; i.src = data; });
      const cw = 1200, ch = 800, c = document.createElement("canvas"); c.width = cw; c.height = ch;
      const sc = Math.max(cw / im.naturalWidth, ch / im.naturalHeight), w = im.naturalWidth * sc, h = im.naturalHeight * sc;
      c.getContext("2d").drawImage(im, (cw - w) / 2, (ch - h) / 2, w, h); return c.toDataURL("image/jpeg", 0.82);
    };
    for (let i = 0; i < fotos.length; i++) {
      if (i % PER === 0) {
        pdf.addPage(); y = M;
        pdf.setFont("helvetica", "bold"); pdf.setFontSize(12); pdf.setTextColor(...navy); pdf.text("Registro fotográfico", M, y + 4);
        pdf.setFont("helvetica", "normal"); pdf.setFontSize(9); pdf.setTextColor(...grey); pdf.text(`${r.terrenoNombre} · ${fmtDT(r.fecha)}`, W - M, y + 4, { align: "right" });
        pdf.setDrawColor(...gold); pdf.setLineWidth(0.4); pdf.line(M, y + 7, W - M, y + 7); pdf.setLineWidth(0.2); y += 13;
      }
      const k = i % PER, col = k % COLS, rowI = Math.floor(k / COLS), x = X0 + col * (CW + GAP), yy = y + rowI * CH;
      try { pdf.addImage(await crop(fotos[i].data), "JPEG", x, yy, CW, IH, undefined, "FAST"); }
      catch (e) { pdf.setFillColor(235, 236, 240); pdf.rect(x, yy, CW, IH, "F"); }
      pdf.setFontSize(8); pdf.setFont("helvetica", "normal"); pdf.setTextColor(...grey); pdf.text("Foto " + (i + 1), x, yy + IH + 4.5);
    }
  }
  const pages = pdf.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    pdf.setPage(p); pdf.setFontSize(8); pdf.setTextColor(...grey);
    pdf.text(`${r.terrenoNombre} · ${fmtDT(r.fecha)}`, M, 292); pdf.text(`${p}/${pages}`, W - M, 292, { align: "right" });
  }
  return pdf.output("blob");
}

/* ---------- compartir ---------- */
function shareText(r) {
  const L = ["*" + repName(r) + "*", "Reporte de visita a terreno · H&C Administración Predial", "", "Fecha: " + fmtDT(r.fecha)];
  if (r.autorNombre) L.push("Generado por: " + r.autorNombre);
  L.push("");
  (r.preguntas || []).forEach((q, i) => L.push(`${i + 1}. ${q.texto}: ${ansLong(q, r.respuestas?.[q.id])}`));
  L.push("", "Observaciones: " + (r.obs || "Sin observaciones."), "Fotos: " + (r.nFotos || 0));
  return L.join("\n");
}
async function prepareShare(r, btn) {
  const old = btn.textContent; btn.disabled = true; btn.textContent = "Generando PDF…";
  try {
    const blob = await buildPdf(r);
    const file = new File([blob], repName(r) + ".pdf", { type: "application/pdf" });
    S.share = { r, file, blob };
    const canShare = !!(navigator.canShare && navigator.canShare({ files: [file] }));
    $("shNative").hidden = !canShare;
    $("shT").textContent = repName(r);
    $("shSub").textContent = canShare ? "Elige cómo enviarlo. Al compartir, el PDF va adjunto." : "Este teléfono no permite adjuntar el PDF directamente: descárgalo y compártelo desde tus archivos.";
    $("shWa").href = "https://wa.me/?text=" + encodeURIComponent(shareText(r));
    $("shSheet").hidden = false;
  } catch (e) {
    console.error(e);
    toast(e && e.message === "jspdf" ? "No se pudo cargar el generador de PDF. Abre la app con señal una vez para que quede guardado." : "No se pudo generar el PDF. Si el reporte tiene fotos, se necesita señal la primera vez.");
  } finally { btn.disabled = false; btn.textContent = old; }
}
$("shNative").onclick = async () => {
  const s = S.share; if (!s) return;
  try { await navigator.share({ files: [s.file], title: repName(s.r), text: "Reporte de visita " + repName(s.r) }); $("shSheet").hidden = true; }
  catch (e) { if (e && e.name !== "AbortError") toast("No se pudo abrir el menú de compartir. Usa Descargar PDF."); }
};
$("shDownload").onclick = () => {
  const s = S.share; if (!s) return;
  const url = URL.createObjectURL(s.blob); const a = document.createElement("a");
  a.href = url; a.download = s.file.name; document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
};
$("shClose").onclick = () => { $("shSheet").hidden = true; };
$("shSheet").onclick = e => { if (e.target === $("shSheet")) $("shSheet").hidden = true; };

/* ---------- administración ---------- */
function showAdmin() {
  S.cur = null; setView("vAdmin"); $("brand").textContent = "Administración";
  if (!S.preguntasLive && isAdmin()) seedPreguntas();
  setAdminTab("q");
}
function setAdminTab(t) {
  $("tabQ").setAttribute("aria-selected", t === "q"); $("tabU").setAttribute("aria-selected", t === "u");
  $("pQ").hidden = t !== "q"; $("pU").hidden = t !== "u";
  if (t === "q") renderQAdmin();
  if (t === "u") {
    if (!S.unsubUsers) S.unsubUsers = onSnapshot(collection(db, "usuarios"), s => { S.usuarios = s.docs.map(d => ({ id: d.id, ...d.data() })); renderUAdmin(); }, () => { });
    renderUAdmin();
  }
}
$("tabQ").onclick = () => setAdminTab("q"); $("tabU").onclick = () => setAdminTab("u");

const TIPO_TXT = { scale: "Escala 1 a 5", yn: "Sí / No", text: "Texto libre" };
function renderQAdmin() {
  const box = $("qList"); box.textContent = "";
  S.preguntas.forEach((q, i) => {
    const r = document.createElement("div"); r.className = "qa-row" + (q.activa === false ? " qa-off" : "");
    r.innerHTML = '<div><div class="qa-t"></div><div class="qa-m"></div></div><div class="qa-btns"></div>';
    r.querySelector(".qa-t").textContent = q.texto;
    r.querySelector(".qa-m").textContent = [TIPO_TXT[q.tipo], q.tipo === "scale" ? `1 = ${q.lo || SCALE_TXT[1]} · 5 = ${q.hi || SCALE_TXT[5]}` : "", q.req === false ? "Opcional" : "Obligatoria", q.activa === false ? "Desactivada" : ""].filter(Boolean).join(" · ");
    const bs = r.querySelector(".qa-btns");
    const mk = (txt, label, fn, dis) => { const b = document.createElement("button"); b.className = "btn small quiet"; b.textContent = txt; b.setAttribute("aria-label", label); b.disabled = !!dis; b.onclick = fn; bs.append(b); };
    mk("↑", "Subir", () => moveQ(i, -1), i === 0);
    mk("↓", "Bajar", () => moveQ(i, 1), i === S.preguntas.length - 1);
    mk("Editar", "Editar pregunta", () => openQ(q));
    mk(q.activa === false ? "Activar" : "Desactivar", "Activar o desactivar", () => setDoc(doc(db, "preguntas", q.id), { activa: q.activa === false }, { merge: true }).catch(adminErr));
    box.append(r);
  });
}
function adminErr() { toast("No se pudo guardar. Solo un administrador puede editar preguntas."); }
async function moveQ(i, dir) {
  const a = S.preguntas[i], b = S.preguntas[i + dir]; if (!a || !b) return;
  const bt = writeBatch(db);
  bt.set(doc(db, "preguntas", a.id), { orden: b.orden }, { merge: true });
  bt.set(doc(db, "preguntas", b.id), { orden: a.orden }, { merge: true });
  bt.commit().catch(adminErr);
}
let editingQ = null;
function openQ(q) {
  editingQ = q || null;
  $("qSheetT").textContent = q ? "Editar pregunta" : "Nueva pregunta";
  $("qText").value = q?.texto || ""; $("qType").value = q?.tipo || "scale";
  $("qLo").value = q?.lo || ""; $("qHi").value = q?.hi || ""; $("qReq").checked = q ? q.req !== false : true;
  $("qEnds").hidden = $("qType").value !== "scale";
  $("qSheet").hidden = false; setTimeout(() => $("qText").focus(), 60);
}
$("qType").onchange = () => { $("qEnds").hidden = $("qType").value !== "scale"; };
$("addQ").onclick = () => openQ(null);
$("qCancel").onclick = () => { $("qSheet").hidden = true; };
$("qForm").onsubmit = e => {
  e.preventDefault();
  const data = { texto: $("qText").value.trim(), tipo: $("qType").value, lo: $("qLo").value.trim(), hi: $("qHi").value.trim(), req: $("qReq").checked };
  if (!data.texto) return;
  if (editingQ) setDoc(doc(db, "preguntas", editingQ.id), data, { merge: true }).catch(adminErr);
  else {
    const orden = Math.max(0, ...S.preguntas.map(q => q.orden || 0)) + 1;
    setDoc(doc(db, "preguntas", slug(data.texto).slice(0, 30) + "-" + uid().slice(-4)), { ...data, orden, activa: true }).catch(adminErr);
  }
  $("qSheet").hidden = true; toast("Pregunta guardada. Ya aparece en todos los teléfonos.");
};

function renderUAdmin() {
  const box = $("uList"); box.textContent = "";
  if (!S.usuarios.length) { box.textContent = "Cargando…"; return; }
  for (const u of S.usuarios.sort((a, b) => (a.nombre || "").localeCompare(b.nombre || "", "es"))) {
    const r = document.createElement("div"); r.className = "user-row";
    const left = document.createElement("div"); const n = document.createElement("div"); n.textContent = (u.nombre || "Sin nombre") + (u.id === S.uid ? " (tú)" : "");
    left.append(n);
    if (u.rol === "admin") { const bdg = document.createElement("span"); bdg.className = "badge"; bdg.textContent = "Administrador"; left.append(bdg); }
    r.append(left);
    if (u.id !== S.uid) {
      const b = document.createElement("button"); b.className = "btn small quiet";
      b.textContent = u.rol === "admin" ? "Quitar admin" : "Hacer admin";
      b.onclick = () => updateDoc(doc(db, "usuarios", u.id), { rol: u.rol === "admin" ? "usuario" : "admin" }).catch(adminErr);
      r.append(b);
    }
    box.append(r);
  }
}

/* ---------- instalación ---------- */
let deferredPrompt = null;
const standalone = matchMedia("(display-mode: standalone)").matches || navigator.standalone;
const isIOS = /iphone|ipad|ipod/i.test(navigator.userAgent);
let installDismissed = false; try { installDismissed = localStorage.getItem("hyc_install_ok") === "1"; } catch (e) { }
addEventListener("beforeinstallprompt", e => {
  e.preventDefault(); deferredPrompt = e;
  if (!standalone && !installDismissed) { $("installTxt").textContent = "Instala la app en este teléfono para abrirla desde su ícono."; $("installBtn").hidden = false; $("installNote").hidden = false; }
});
$("installBtn").onclick = async () => {
  if (!deferredPrompt) return; deferredPrompt.prompt(); await deferredPrompt.userChoice; deferredPrompt = null; $("installNote").hidden = true;
};
if (isIOS && !standalone && !installDismissed) {
  $("installTxt").textContent = "Para instalarla: toca el botón Compartir de Safari y luego “Agregar a inicio”.";
  $("installNote").hidden = false;
}
$("installHide").onclick = () => { $("installNote").hidden = true; try { localStorage.setItem("hyc_install_ok", "1"); } catch (e) { } };

/* ---------- service worker (funciona sin señal y se actualiza solo) ---------- */
if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("sw.js").then(reg => {
    reg.addEventListener("updatefound", () => {
      const nw = reg.installing;
      nw && nw.addEventListener("statechange", () => { if (nw.state === "installed" && navigator.serviceWorker.controller) toast("Hay una versión nueva. Se aplicará la próxima vez que abras la app."); });
    });
  }).catch(() => { });
}
