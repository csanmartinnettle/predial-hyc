/* Administración Predial - H&C
   App Android simple para reportes de visitas a terrenos.
   Sin cuentas ni servidores: todo queda guardado en el teléfono.
   Al finalizar una visita, el PDF se guarda en Documentos › Predial HyC y se abre el menú para compartirlo. */

import { jsPDF } from "jspdf";
import { Capacitor } from "@capacitor/core";
import { Share } from "@capacitor/share";
import { Filesystem, Directory } from "@capacitor/filesystem";
import { StatusBar, Style } from "@capacitor/status-bar";
import "./fonts.css";

const NATIVE = Capacitor.isNativePlatform();
if (NATIVE) {
  document.documentElement.classList.add("native");
  StatusBar.setBackgroundColor({ color: "#1b223d" }).catch(() => { });
  StatusBar.setStyle({ style: Style.Dark }).catch(() => { });
}

/* ---------- valores iniciales (se cambian actualizando la app) ---------- */
const MAX_FOTOS = 8;
const TERRENOS_INICIALES = ["Los Andes", "Suizos", "Reina Norte", "Batuco", "Lote G", "Talagante"];
// Para cambiar las preguntas de todos los teléfonos: editar esta lista y subir la VERSION_PREGUNTAS.
const VERSION_PREGUNTAS = 2;
const PREGUNTAS_INICIALES = [
  { id: "cierre", texto: "Estado del cierre perimetral", tipo: "scale", lo: "Muy malo", hi: "Muy bueno", req: true },
  { id: "personas", texto: "Presencia de personas ajenas al terreno", tipo: "yn", req: true },
  { id: "vandalismo", texto: "Vandalismo en terreno", tipo: "yn", nota: true, req: true },
  { id: "animales", texto: "Presencia de animales no deseados", tipo: "yn", nota: true, req: true },
  { id: "cuidador", texto: "Evaluación gestión cuidador", tipo: "scale", lo: "Muy mala", hi: "Muy buena", nota: true, notaLabel: "Comentarios", req: true }
];
const SCALE_TXT = { 1: "Muy malo", 2: "Malo", 3: "Regular", 4: "Bueno", 5: "Muy bueno" };
const CARPETA_PDF = "Predial HyC";

const $ = id => document.getElementById(id);
const S = { nombre: "", terrenos: [], preguntas: [], reportes: [], cur: null, draft: null, share: null };

/* ---------- utilidades ---------- */
function toast(t) { const el = $("toast"); el.textContent = t; el.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => el.hidden = true, 3800); }
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
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
const lower = s => String(s || "").trim().toLowerCase();
const activeQuestions = () => S.preguntas.filter(q => q.activa !== false);
const safeName = s => String(s).replace(/[\\/:*?"<>|#%]/g, "-").trim();
function blobToDataUrl(blob) { return new Promise((ok, no) => { const fr = new FileReader(); fr.onload = () => ok(String(fr.result)); fr.onerror = no; fr.readAsDataURL(blob); }); }
async function blobToBase64(blob) { return (await blobToDataUrl(blob)).split(",")[1]; }

async function compress(file, max = 1280, q = 0.72) {
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = url; });
    const sc = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
    const w = Math.round(img.naturalWidth * sc), h = Math.round(img.naturalHeight * sc);
    const c = document.createElement("canvas"); c.width = w; c.height = h; c.getContext("2d").drawImage(img, 0, 0, w, h);
    return { data: c.toDataURL("image/jpeg", q), w, h };
  } finally { URL.revokeObjectURL(url); }
}

/* ---------- almacenamiento en el teléfono ---------- */
const mem = { drafts: {}, reportes: {} };
let idbP = null;
function idb() {
  if (idbP) return idbP;
  idbP = new Promise(res => {
    try {
      const r = indexedDB.open("hyc-predial-local", 1);
      r.onupgradeneeded = () => { r.result.createObjectStore("drafts"); r.result.createObjectStore("reportes"); };
      r.onsuccess = () => res(r.result); r.onerror = () => res(null);
    } catch (e) { res(null); }
  });
  return idbP;
}
async function kvGet(st, k) {
  const d = await idb(); if (!d) return mem[st][k];
  return new Promise(res => { try { const r = d.transaction(st).objectStore(st).get(k); r.onsuccess = () => res(r.result); r.onerror = () => res(undefined); } catch (e) { res(undefined); } });
}
async function kvPut(st, k, v) {
  const d = await idb(); if (!d) { mem[st][k] = v; return; }
  return new Promise((res, rej) => { try { const tx = d.transaction(st, "readwrite"); tx.objectStore(st).put(v, k); tx.oncomplete = res; tx.onerror = () => rej(tx.error); } catch (e) { rej(e); } });
}
async function kvDel(st, k) {
  const d = await idb(); if (!d) { delete mem[st][k]; return; }
  return new Promise(res => { try { const tx = d.transaction(st, "readwrite"); tx.objectStore(st).delete(k); tx.oncomplete = res; tx.onerror = res; } catch (e) { res(); } });
}
async function kvAll(st) {
  const d = await idb(); if (!d) return Object.values(mem[st]);
  return new Promise(res => { try { const r = d.transaction(st).objectStore(st).getAll(); r.onsuccess = () => res(r.result || []); r.onerror = () => res([]); } catch (e) { res([]); } });
}
const ls = {
  get(k, def) { try { const v = localStorage.getItem(k); return v == null ? def : JSON.parse(v); } catch (e) { return def; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { } }
};
function savePrefs() { ls.set("hyc_nombre", S.nombre); ls.set("hyc_terrenos", S.terrenos); ls.set("hyc_preguntas", S.preguntas); }

/* ---------- arranque ---------- */
(async () => {
  S.nombre = ls.get("hyc_nombre", "");
  S.terrenos = ls.get("hyc_terrenos", null) || TERRENOS_INICIALES.map(n => ({ id: uid(), nombre: n }));
  // Las preguntas se reemplazan por las de la app cuando se instala una versión con preguntas nuevas.
  const pv = ls.get("hyc_preguntas_version", 0);
  S.preguntas = (pv === VERSION_PREGUNTAS && ls.get("hyc_preguntas", null)) || PREGUNTAS_INICIALES.map((q, i) => ({ ...q, orden: i + 1, activa: true }));
  ls.set("hyc_preguntas_version", VERSION_PREGUNTAS);
  savePrefs();
  S.reportes = (await kvAll("reportes")).sort((a, b) => b.fecha.localeCompare(a.fecha));
  if (!S.nombre) showOnboard();
  showWho(); route();
})();

/* ---------- nombre (firma) ---------- */
function showOnboard() { $("userName").value = S.nombre; $("nameErr").textContent = ""; $("onboard").hidden = false; setTimeout(() => $("userName").focus(), 60); }
$("nameForm").onsubmit = e => {
  e.preventDefault();
  const n = $("userName").value.trim().replace(/\s+/g, " ");
  if (n.split(" ").length < 2) { $("nameErr").textContent = "Escribe nombre y apellido."; return; }
  const first = !S.nombre; S.nombre = n; savePrefs(); showWho(); $("onboard").hidden = true;
  toast(first ? "Hola, " + n.split(" ")[0] + "." : "Nombre actualizado.");
};
$("changeName").onclick = showOnboard;
function showWho() { $("whoName").textContent = S.nombre || "—"; $("adminBtn").hidden = false; }

/* ---------- navegación ---------- */
function route() {
  const h = location.hash.replace("#", "");
  if (h === "admin") return showAdmin();
  const id = h.startsWith("t-") ? h.slice(2) : null;
  const lot = id && S.terrenos.find(t => String(t.id) === id);
  if (lot) { if (!S.cur || S.cur.id !== lot.id) openLot(lot); else render(); }
  else showHome();
}
addEventListener("hashchange", route);
$("back").onclick = () => { history.length > 1 ? history.back() : (location.hash = ""); };
$("adminBtn").onclick = () => { location.hash = "admin"; };

function setView(v) {
  for (const id of ["vHome", "vLot", "vAdmin"]) $(id).hidden = id !== v;
  $("back").hidden = v === "vHome"; $("logoImg").hidden = v !== "vHome";
  $("adminBtn").hidden = v === "vAdmin";
  $("actionbar").hidden = !(v === "vLot" && !$("pVisit").hidden);
}
function showHome() { S.cur = null; setView("vHome"); $("brand").textContent = "Administración Predial"; render(); }
function render() {
  if (!$("vHome").hidden) renderHome();
  if (S.cur) { updateTabCount(); if (!$("pReps").hidden) renderReps(); }
  if (!$("vAdmin").hidden) renderQAdmin();
}

/* ---------- inicio ---------- */
const repsOf = tid => S.reportes.filter(r => String(r.terrenoId) === String(tid));
function renderHome() {
  const box = $("lots"); box.textContent = "";
  for (const t of S.terrenos) {
    const reps = repsOf(t.id);
    const b = document.createElement("button"); b.className = "lot"; b.type = "button";
    b.innerHTML = '<span class="pin"></span><span><div class="lot-name"></div><div class="lot-meta"></div></span><span class="chev">›</span>';
    b.querySelector(".pin").textContent = t.nombre.trim().charAt(0).toUpperCase();
    b.querySelector(".lot-name").textContent = t.nombre;
    b.querySelector(".lot-meta").textContent = reps.length ? `Última visita ${fmtDate(reps[0].fecha)} · ${reps.length} reporte${reps.length > 1 ? "s" : ""}` : "Sin visitas registradas";
    b.onclick = () => { location.hash = "t-" + t.id; };
    box.append(b);
  }
}
$("addForm").onsubmit = e => {
  e.preventDefault(); const name = $("newLot").value.trim(); if (!name) return;
  if (S.terrenos.some(t => lower(t.nombre) === lower(name))) { toast("Ese terreno ya existe."); return; }
  S.terrenos.push({ id: uid(), nombre: name }); S.terrenos.sort((x, y) => x.nombre.localeCompare(y.nombre, "es"));
  savePrefs(); $("newLot").value = ""; toast("Terreno agregado: " + name); render();
};

/* ---------- terreno ---------- */
async function openLot(lot) {
  S.cur = lot; setView("vLot"); $("brand").textContent = lot.nombre; $("lotName").textContent = lot.nombre;
  S.draft = (await kvGet("drafts", String(lot.id))) || { respuestas: {}, obs: "", fotos: [], inicio: new Date().toISOString() };
  if (S.cur !== lot) return;
  $("obs").value = S.draft.obs || "";
  renderSurvey(); renderThumbs(); setTab("visit"); resetConfirm(); updateTabCount(); scrollTo(0, 0);
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
    if (!S.cur) return;
    try { await kvPut("drafts", String(S.cur.id), S.draft); } catch (e) { toast("No queda espacio en el teléfono para el borrador."); return; }
    $("savedTxt").textContent = "Borrador guardado en este teléfono · " + new Date().toLocaleTimeString("es-CL", { hour: "2-digit", minute: "2-digit" });
  }, 300);
}

function renderSurvey() {
  const box = $("survey"); box.textContent = "";
  activeQuestions().forEach((q, i) => {
    const k = String(q.id);
    const d = document.createElement("div"); d.className = "q"; d.dataset.k = k;
    const t = document.createElement("div"); t.className = "q-t"; t.innerHTML = "<span></span><div></div>";
    t.firstChild.textContent = (i + 1) + "."; t.lastChild.textContent = q.texto + (q.req === false ? " (opcional)" : ""); d.append(t);
    if (q.tipo === "text") {
      const ta = document.createElement("textarea"); ta.id = "q_" + k; ta.style.minHeight = "70px";
      ta.value = S.draft.respuestas[k] || "";
      ta.oninput = () => { S.draft.respuestas[k] = ta.value; d.classList.remove("missing"); saveDraft(); };
      d.append(ta);
    } else {
      const sc = document.createElement("div"); sc.className = "scale" + (q.tipo === "yn" ? " yn" : ""); sc.setAttribute("role", "radiogroup"); sc.setAttribute("aria-label", q.texto);
      const opts = q.tipo === "yn" ? [["si", "Sí"], ["no", "No"]] : [1, 2, 3, 4, 5].map(n => [String(n), String(n)]);
      for (const [v, l] of opts) {
        const lab = document.createElement("label"); const inp = document.createElement("input");
        inp.type = "radio"; inp.name = "q_" + k; inp.id = "q_" + k + "_" + v; inp.value = v;
        inp.checked = String(S.draft.respuestas[k] ?? "") === v;
        inp.onchange = () => { S.draft.respuestas[k] = q.tipo === "yn" ? v : Number(v); d.classList.remove("missing"); saveDraft(); };
        const s = document.createElement("span"); s.className = "opt"; s.textContent = l; lab.append(inp, s); sc.append(lab);
      }
      d.append(sc);
      if (q.tipo === "scale") {
        const e = document.createElement("div"); e.className = "ends"; e.innerHTML = "<span></span><span></span>";
        e.firstChild.textContent = "1 · " + (q.lo || SCALE_TXT[1]); e.lastChild.textContent = (q.hi || SCALE_TXT[5]) + " · 5"; d.append(e);
      }
      if (q.nota) {
        const ta = document.createElement("textarea"); ta.id = "q_" + k + "_nota"; ta.className = "q-nota";
        ta.placeholder = (q.notaLabel || "Observaciones") + " (opcional)"; ta.setAttribute("aria-label", (q.notaLabel || "Observaciones") + " sobre " + q.texto);
        ta.value = S.draft.respuestas[k + "__nota"] || "";
        ta.oninput = () => { S.draft.respuestas[k + "__nota"] = ta.value; saveDraft(); };
        d.append(ta);
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

/* ---------- finalizar ---------- */
function resetConfirm() { $("finish").hidden = false; $("confirm").hidden = true; }
$("finish").textContent = "Finalizar y guardar";
$("finish").onclick = () => {
  if (!S.nombre) { showOnboard(); return; }
  const miss = activeQuestions().filter(q => q.req !== false && (S.draft.respuestas[String(q.id)] == null || S.draft.respuestas[String(q.id)] === ""));
  document.querySelectorAll(".q").forEach(el => el.classList.toggle("missing", miss.some(q => String(q.id) === el.dataset.k)));
  if (miss.length) {
    toast(`Falta responder ${miss.length} pregunta${miss.length > 1 ? "s" : ""} del reporte.`);
    document.querySelector(".q.missing")?.scrollIntoView({ behavior: "smooth", block: "center" }); return;
  }
  $("finish").hidden = true; $("confirm").hidden = false;
};
$("cancelSend").onclick = resetConfirm;
$("doSend").textContent = "Confirmar";
$("doSend").onclick = async () => {
  const lot = S.cur, d = S.draft;
  const preguntas = activeQuestions().map(q => ({ id: String(q.id), texto: q.texto, tipo: q.tipo, lo: q.lo || "", hi: q.hi || "", nota: !!q.nota, notaLabel: q.notaLabel || "" }));
  const respuestas = {}; for (const q of preguntas) {
    if (d.respuestas[q.id] != null && d.respuestas[q.id] !== "") respuestas[q.id] = d.respuestas[q.id];
    const n = (d.respuestas[q.id + "__nota"] || "").trim(); if (q.nota && n) respuestas[q.id + "__nota"] = n;
  }
  const fecha = new Date().toISOString();
  const base = repName({ fecha, terrenoNombre: lot.nombre });
  const usados = new Set(S.reportes.map(r => r.nombre));
  let nombre = base; for (let n = 2; usados.has(nombre); n++) nombre = `${base} (${n})`;
  const rep = {
    id: uid(), nombre, terrenoId: String(lot.id), terrenoNombre: lot.nombre, fecha, inicio: d.inicio,
    autorNombre: S.nombre, preguntas, respuestas, obs: (d.obs || "").trim(),
    fotos: d.fotos.map(f => ({ data: f.data, w: f.w, h: f.h })), nFotos: d.fotos.length
  };
  try { await kvPut("reportes", rep.id, rep); }
  catch (e) { toast("No queda espacio en el teléfono para guardar el reporte. Elimina reportes antiguos ya compartidos."); return; }
  await kvDel("drafts", String(lot.id));
  S.reportes.unshift(rep);
  S.draft = { respuestas: {}, obs: "", fotos: [], inicio: new Date().toISOString() };
  $("obs").value = ""; renderSurvey(); renderThumbs(); resetConfirm();
  $("savedTxt").textContent = "El borrador se guarda en este teléfono.";
  setTab("reps");
  // Respaldo: el PDF queda también en Documentos › Predial HyC, aunque se desinstale la app.
  try {
    const blob = await buildPdf(rep, rep.fotos);
    if (NATIVE) await savePdfToDocuments(rep.nombre, blob);
    openShare(rep, blob, "Reporte guardado" + (NATIVE ? " · PDF en Documentos › " + CARPETA_PDF : ""));
  } catch (e) { console.error(e); toast("Reporte guardado. No se pudo generar el PDF ahora; inténtalo desde Reportes."); }
};

/* ---------- reportes ---------- */
function updateTabCount() {
  if (!S.cur) return; const reps = repsOf(S.cur.id);
  $("tabReps").textContent = reps.length ? `Reportes (${reps.length})` : "Reportes";
  $("lotSub").textContent = reps.length ? `Última visita: ${fmtDate(reps[0].fecha)}` : "Sin visitas registradas";
}
function renderReps() {
  const box = $("pReps"); box.textContent = ""; if (!S.cur) return;
  const reps = repsOf(S.cur.id);
  if (!reps.length) { const e = document.createElement("div"); e.className = "empty"; e.textContent = "Todavía no hay reportes de este terreno en este teléfono."; box.append(e); return; }
  for (const r of reps) {
    const c = document.createElement("div"); c.className = "card rep";
    c.innerHTML = '<div class="rep-h"><div><div class="rep-date"></div><div class="rep-by"></div></div></div><div class="chips"></div>';
    c.querySelector(".rep-date").textContent = r.nombre || repName(r);
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
    const row = document.createElement("div"); row.className = "share"; row.style.gridTemplateColumns = "1fr";
    const btn = document.createElement("button"); btn.className = "btn"; btn.textContent = "Compartir / guardar PDF";
    btn.onclick = async () => {
      const old = btn.textContent; btn.disabled = true; btn.textContent = "Generando PDF…";
      try { openShare(r, await buildPdf(r, r.fotos)); } catch (e) { console.error(e); toast("No se pudo generar el PDF."); }
      finally { btn.disabled = false; btn.textContent = old; }
    };
    row.append(btn); c.append(row);
    box.append(c);
  }
}

/* ---------- PDF ---------- */
async function buildPdf(r, fotos) {
  const doc_ = new jsPDF({ unit: "mm", format: "a4" });
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
      const nota = r.respuestas?.[q.id + "__nota"];
      if (nota) {
        pdf.setFont("helvetica", "italic"); pdf.setTextColor(...grey);
        for (const l of pdf.splitTextToSize((q.notaLabel || "Observaciones") + ": " + nota, W - 2 * M - 6)) { y += 5; if (y > 282) { pdf.addPage(); y = M; } pdf.text(l, M + 6, y); }
        pdf.setTextColor(...navy);
      }
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

  if (fotos && fotos.length) {
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
  const L = ["*" + (r.nombre || repName(r)) + "*", "Reporte de visita a terreno · H&C Administración Predial", "", "Fecha: " + fmtDT(r.fecha)];
  if (r.autorNombre) L.push("Generado por: " + r.autorNombre);
  L.push("");
  (r.preguntas || []).forEach((q, i) => { L.push(`${i + 1}. ${q.texto}: ${ansLong(q, r.respuestas?.[q.id])}`); const n = r.respuestas?.[q.id + "__nota"]; if (n) L.push("   " + (q.notaLabel || "Observaciones") + ": " + n); });
  L.push("", "Observaciones: " + (r.obs || "Sin observaciones."), "Fotos: " + (r.nFotos || 0));
  return L.join("\n");
}
async function savePdfToDocuments(name, blob) {
  await Filesystem.writeFile({ path: `${CARPETA_PDF}/${safeName(name)}.pdf`, data: await blobToBase64(blob), directory: Directory.Documents, recursive: true });
}
function openShare(r, blob, title) {
  const name = r.nombre || repName(r);
  const file = new File([blob], safeName(name) + ".pdf", { type: "application/pdf" });
  S.share = { r, file, blob, name };
  const canShare = NATIVE || !!(navigator.canShare && navigator.canShare({ files: [file] }));
  $("shNative").hidden = !canShare;
  $("shT").textContent = title || name;
  $("shSub").textContent = (title ? name + ". " : "") + (canShare ? "Elige cómo enviarlo: el PDF va adjunto." : "Descarga el PDF y compártelo desde tus archivos.");
  $("shDownload").textContent = NATIVE ? "Guardar PDF en Documentos" : "Descargar PDF";
  $("shWa").href = "https://wa.me/?text=" + encodeURIComponent(shareText(r));
  $("shDelete").textContent = "Eliminar reporte"; delete $("shDelete").dataset.armed;
  $("shSheet").hidden = false;
}
$("shNative").onclick = async () => {
  const s = S.share; if (!s) return;
  try {
    if (NATIVE) {
      const w = await Filesystem.writeFile({ path: s.file.name, data: await blobToBase64(s.blob), directory: Directory.Cache });
      await Share.share({ title: s.name, text: "Reporte de visita " + s.name, files: [w.uri], dialogTitle: "Compartir reporte" });
    } else {
      await navigator.share({ files: [s.file], title: s.name, text: "Reporte de visita " + s.name });
    }
    $("shSheet").hidden = true;
  } catch (e) {
    const msg = String(e && (e.message || e.name) || "");
    if (!/cancel|abort/i.test(msg)) toast("No se pudo abrir el menú de compartir.");
  }
};
$("shDownload").onclick = async () => {
  const s = S.share; if (!s) return;
  if (NATIVE) {
    try { await savePdfToDocuments(s.name, s.blob); toast("PDF guardado en Documentos › " + CARPETA_PDF + "."); }
    catch (e) { toast("No se pudo guardar. Usa Compartir y elige Guardar en Drive."); }
    return;
  }
  const url = URL.createObjectURL(s.blob); const a = document.createElement("a");
  a.href = url; a.download = s.file.name; document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
};
$("shDelete").onclick = async () => {
  const s = S.share; if (!s) return; const b = $("shDelete");
  if (!b.dataset.armed) { b.dataset.armed = "1"; b.textContent = "Toca de nuevo para eliminar de la app"; return; }
  await kvDel("reportes", s.r.id);
  S.reportes = S.reportes.filter(r => r.id !== s.r.id);
  $("shSheet").hidden = true; toast("Reporte eliminado de la app. El PDF en Documentos se mantiene."); render();
};
$("shClose").onclick = () => { $("shSheet").hidden = true; };
$("shSheet").onclick = e => { if (e.target === $("shSheet")) $("shSheet").hidden = true; };

/* ---------- preguntas (ajustes del teléfono) ---------- */
function showAdmin() { S.cur = null; setView("vAdmin"); $("brand").textContent = "Preguntas"; renderQAdmin(); }
function saveQ() { S.preguntas.forEach((q, i) => q.orden = i + 1); savePrefs(); renderQAdmin(); }
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
    mk("↑", "Subir", () => { [S.preguntas[i - 1], S.preguntas[i]] = [S.preguntas[i], S.preguntas[i - 1]]; saveQ(); }, i === 0);
    mk("↓", "Bajar", () => { [S.preguntas[i + 1], S.preguntas[i]] = [S.preguntas[i], S.preguntas[i + 1]]; saveQ(); }, i === S.preguntas.length - 1);
    mk("Editar", "Editar pregunta", () => openQ(q));
    mk(q.activa === false ? "Activar" : "Desactivar", "Activar o desactivar", () => { q.activa = q.activa === false; saveQ(); });
    box.append(r);
  });
}
let editingQ = null;
function openQ(q) {
  editingQ = q || null;
  $("qSheetT").textContent = q ? "Editar pregunta" : "Nueva pregunta";
  $("qText").value = q?.texto || ""; $("qType").value = q?.tipo || "scale";
  $("qLo").value = q?.lo || ""; $("qHi").value = q?.hi || ""; $("qReq").checked = q ? q.req !== false : true; $("qNota").checked = !!q?.nota;
  $("qEnds").hidden = $("qType").value !== "scale";
  $("qSheet").hidden = false; setTimeout(() => $("qText").focus(), 60);
}
$("qType").onchange = () => { $("qEnds").hidden = $("qType").value !== "scale"; };
$("addQ").onclick = () => openQ(null);
$("qCancel").onclick = () => { $("qSheet").hidden = true; };
$("qForm").onsubmit = e => {
  e.preventDefault();
  const f = { texto: $("qText").value.trim(), tipo: $("qType").value, lo: $("qLo").value.trim(), hi: $("qHi").value.trim(), req: $("qReq").checked, nota: $("qNota").checked };
  if (!f.texto) return;
  if (editingQ) Object.assign(editingQ, f); else S.preguntas.push({ id: uid(), ...f, activa: true });
  $("qSheet").hidden = true; saveQ(); toast("Pregunta guardada en este teléfono.");
};

/* ---------- versión web (sin APK) ---------- */
let deferredPrompt = null;
const standalone = matchMedia("(display-mode: standalone)").matches || navigator.standalone;
const isIOS = /iphone|ipad|ipod/i.test(navigator.userAgent);
let installDismissed = false; try { installDismissed = localStorage.getItem("hyc_install_ok") === "1"; } catch (e) { }
if (!NATIVE) addEventListener("beforeinstallprompt", e => {
  e.preventDefault(); deferredPrompt = e;
  if (!standalone && !installDismissed) { $("installTxt").textContent = "Instala la app en este teléfono para abrirla desde su ícono."; $("installBtn").hidden = false; $("installNote").hidden = false; }
});
$("installBtn").onclick = async () => { if (!deferredPrompt) return; deferredPrompt.prompt(); await deferredPrompt.userChoice; deferredPrompt = null; $("installNote").hidden = true; };
if (!NATIVE && isIOS && !standalone && !installDismissed) { $("installTxt").textContent = "Para instalarla: toca el botón Compartir de Safari y luego “Agregar a inicio”."; $("installNote").hidden = false; }
$("installHide").onclick = () => { $("installNote").hidden = true; try { localStorage.setItem("hyc_install_ok", "1"); } catch (e) { } };
if (!NATIVE && "serviceWorker" in navigator) navigator.serviceWorker.register("sw.js", { updateViaCache: "none" }).then(r => r.update()).catch(() => { });
