/* Administración Predial - H&C
   App Android (Capacitor) para reportes de visitas a terrenos.
   Datos en Microsoft 365: inicio de sesión con la cuenta de H&C (MSAL) y SharePoint (Microsoft Graph).
   - Listas del sitio: "Predial - Terrenos", "Predial - Preguntas", "Predial - Reportes", "Predial - Administradores".
   - Biblioteca "Documentos": <carpeta>/Reportes/<Terreno>/<AAAA MM DD TERRENO>/ con las fotos y el PDF. */

import { m365Config } from "./m365-config.js";
import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";
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

/* ---------- constantes ---------- */
const MAX_FOTOS = 8;
const SCOPES = ["User.Read", "Sites.ReadWrite.All", "Sites.Manage.All"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const TERRENOS_INICIALES = ["Los Andes", "Suizos", "Reina Norte", "Batuco", "Lote G", "Talagante"];
const PREGUNTAS_INICIALES = [
  { texto: "Estado del cierre perimetral", tipo: "scale", lo: "Muy malo", hi: "Muy bueno" },
  { texto: "Presencia de personas ajenas al terreno", tipo: "yn" },
  { texto: "Vandalismo en terreno", tipo: "scale", lo: "Grave", hi: "Sin vandalismo" },
  { texto: "Presencia de animales no deseados", tipo: "scale", lo: "Muchos", hi: "Ninguno" }
];
const LISTAS = {
  terrenos: { nombre: "Predial - Terrenos", cols: [] },
  preguntas: {
    nombre: "Predial - Preguntas", cols: [
      { name: "Tipo", text: {} }, { name: "Lo", text: {} }, { name: "Hi", text: {} },
      { name: "Req", boolean: {} }, { name: "Activa", boolean: {} }, { name: "Orden", number: {} }]
  },
  reportes: {
    nombre: "Predial - Reportes", cols: [
      { name: "TerrenoId", text: {} }, { name: "Terreno", text: {} }, { name: "Fecha", text: {} },
      { name: "Autor", text: {} }, { name: "AutorEmail", text: {} }, { name: "NFotos", number: {} },
      { name: "Carpeta", text: { allowMultipleLines: true, textType: "plain" } },
      { name: "Obs", text: { allowMultipleLines: true, textType: "plain" } },
      { name: "Datos", text: { allowMultipleLines: true, textType: "plain" } }]
  },
  admins: { nombre: "Predial - Administradores", cols: [] }
};
const SCALE_TXT = { 1: "Muy malo", 2: "Malo", 3: "Regular", 4: "Bueno", 5: "Muy bueno" };

const $ = id => document.getElementById(id);
const S = {
  me: null, site: null, lists: {}, terrenos: [], preguntas: [], reportes: [], admins: [],
  outbox: [], cur: null, draft: null, share: null, ready: false, syncing: false, needLogin: false
};

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
const isAdmin = () => !!S.me && S.admins.some(a => lower(a.email) === lower(S.me.email));
const activeQuestions = () => S.preguntas.filter(q => q.activa !== false);
const safeName = s => String(s).replace(/[\\/:*?"<>|#%]/g, "-").trim();
function dataUrlToBlob(d) { const [h, b] = d.split(","); const bin = atob(b); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return new Blob([u], { type: h.slice(5, h.indexOf(";")) }); }
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
const mem = { drafts: {}, outbox: {} };
let idbP = null;
function idb() {
  if (idbP) return idbP;
  idbP = new Promise(res => {
    try {
      const r = indexedDB.open("hyc-predial-m365", 1);
      r.onupgradeneeded = () => { r.result.createObjectStore("drafts"); r.result.createObjectStore("outbox"); };
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
  return new Promise(res => { try { const tx = d.transaction(st, "readwrite"); tx.objectStore(st).put(v, k); tx.oncomplete = res; tx.onerror = res; } catch (e) { res(); } });
}
async function kvDel(st, k) {
  const d = await idb(); if (!d) { delete mem[st][k]; return; }
  return new Promise(res => { try { const tx = d.transaction(st, "readwrite"); tx.objectStore(st).delete(k); tx.oncomplete = res; tx.onerror = res; } catch (e) { res(); } });
}
async function kvAll(st) {
  const d = await idb(); if (!d) return Object.values(mem[st]);
  return new Promise(res => { try { const r = d.transaction(st).objectStore(st).getAll(); r.onsuccess = () => res(r.result || []); r.onerror = () => res([]); } catch (e) { res([]); } });
}
function cacheSave() {
  try { localStorage.setItem("hyc_cache", JSON.stringify({ me: S.me, site: S.site, lists: S.lists, terrenos: S.terrenos, preguntas: S.preguntas, reportes: S.reportes, admins: S.admins })); } catch (e) { }
}
function cacheLoad() {
  try { const c = JSON.parse(localStorage.getItem("hyc_cache") || "null"); if (c) Object.assign(S, c); } catch (e) { }
}

/* ---------- configuración ---------- */
const CFG = m365Config || {};
if (!CFG.clientId || CFG.clientId.startsWith("PEGAR") || !CFG.siteUrl || CFG.siteUrl.startsWith("PEGAR")) {
  document.body.innerHTML = '<main class="wrap"><h1>Falta configurar</h1><p class="sub">Completa el archivo <b>m365-config.js</b> con el Id. de aplicación de Microsoft Entra y la dirección del sitio de SharePoint, como indica la guía de instalación.</p></main>';
  throw new Error("m365-config.js sin completar");
}

/* ---------- inicio de sesión (Microsoft) ---------- */
const pca = new PublicClientApplication({
  auth: {
    clientId: CFG.clientId,
    authority: "https://login.microsoftonline.com/" + CFG.tenant,
    redirectUri: location.origin + "/",
    postLogoutRedirectUri: location.origin + "/"
  },
  cache: { cacheLocation: "localStorage" }
});
const account = () => pca.getActiveAccount() || pca.getAllAccounts()[0] || null;

async function token() {
  const acc = account();
  if (!acc) { const e = new Error("login"); e.login = true; throw e; }
  try { return (await pca.acquireTokenSilent({ scopes: SCOPES, account: acc })).accessToken; }
  catch (e) {
    if (e instanceof InteractionRequiredAuthError) { S.needLogin = true; $("authNote").hidden = false; const x = new Error("login"); x.login = true; throw x; }
    throw e;
  }
}
async function login() {
  $("loginErr").textContent = "";
  try { await pca.loginRedirect({ scopes: SCOPES, prompt: "select_account", loginHint: "@" + CFG.tenant }); }
  catch (e) { $("loginErr").textContent = "No se pudo abrir el inicio de sesión de Microsoft. Revisa la conexión."; }
}
$("loginBtn").onclick = login;
$("reloginBtn").onclick = async () => { const acc = account(); try { await pca.acquireTokenRedirect({ scopes: SCOPES, account: acc || undefined }); } catch (e) { login(); } };
$("logoutBtn").onclick = async () => {
  const pend = (await kvAll("outbox")).length;
  if (pend && !confirmLogout()) return;
  try { await pca.logoutRedirect({ account: account(), onRedirectNavigate: () => false }); } catch (e) { }
  try { localStorage.removeItem("hyc_cache"); } catch (e) { }
  location.hash = ""; location.reload();
};
let logoutArmed = 0;
function confirmLogout() {
  if (Date.now() - logoutArmed < 5000) return true;
  logoutArmed = Date.now(); toast("Hay reportes sin enviar en este teléfono. Toca de nuevo Cerrar sesión para confirmar."); return false;
}

/* ---------- Microsoft Graph ---------- */
async function g(path, opts = {}) {
  const t = await token();
  const headers = { Authorization: "Bearer " + t, ...(opts.headers || {}) };
  let body = opts.body;
  if (opts.json !== undefined) { headers["Content-Type"] = "application/json"; body = JSON.stringify(opts.json); }
  const r = await fetch(path.startsWith("http") ? path : GRAPH + path, { method: opts.method || "GET", headers, body });
  if (!r.ok) {
    const e = new Error("Graph " + r.status); e.status = r.status;
    try { e.body = await r.json(); } catch (x) { }
    if (r.status === 401) { S.needLogin = true; $("authNote").hidden = false; }
    throw e;
  }
  if (r.status === 204) return null;
  return opts.raw ? r : r.json();
}
async function gAll(path) {
  const out = []; let url = path;
  while (url) { const j = await g(url); out.push(...(j.value || [])); url = j["@odata.nextLink"] || null; }
  return out;
}
const encPath = p => p.split("/").filter(Boolean).map(encodeURIComponent).join("/");
const drivePath = (...parts) => [CFG.carpeta || "Administración Predial", ...parts].map(safeName).join("/");

async function uploadFile(path, blob, type) {
  const enc = encPath(path);
  if (blob.size < 3.5e6) {
    return g(`/sites/${S.site.id}/drive/root:/${enc}:/content?@microsoft.graph.conflictBehavior=replace`, { method: "PUT", body: blob, headers: { "Content-Type": type } });
  }
  const s = await g(`/sites/${S.site.id}/drive/root:/${enc}:/createUploadSession`, { method: "POST", json: { item: { "@microsoft.graph.conflictBehavior": "replace" } } });
  const CH = 320 * 1024 * 16; let pos = 0, res = null;
  while (pos < blob.size) {
    const end = Math.min(pos + CH, blob.size);
    const r = await fetch(s.uploadUrl, { method: "PUT", headers: { "Content-Range": `bytes ${pos}-${end - 1}/${blob.size}` }, body: blob.slice(pos, end) });
    if (!r.ok) throw new Error("Subida " + r.status);
    if (r.status !== 202) res = await r.json();
    pos = end;
  }
  return res;
}
async function driveItem(path) {
  try { return await g(`/sites/${S.site.id}/drive/root:/${encPath(path)}`); }
  catch (e) { if (e.status === 404) return null; throw e; }
}

/* ---------- SharePoint: sitio y listas ---------- */
async function resolveSite() {
  if (S.site && S.site.url === CFG.siteUrl) return;
  const u = new URL(CFG.siteUrl);
  const path = decodeURIComponent(u.pathname).replace(/\/+$/, "");
  const s = await g(path ? `/sites/${u.hostname}:${path}` : `/sites/${u.hostname}`);
  S.site = { id: s.id, url: CFG.siteUrl, webUrl: s.webUrl }; S.lists = {};
}
async function ensureLists() {
  const need = Object.keys(LISTAS).filter(k => !S.lists[k]);
  if (!need.length) return;
  const all = await gAll(`/sites/${S.site.id}/lists?$select=id,displayName`);
  for (const k of need) { const l = all.find(x => x.displayName === LISTAS[k].nombre); if (l) S.lists[k] = l.id; }
  const missing = Object.keys(LISTAS).filter(k => !S.lists[k]);
  if (!missing.length) return;
  // Primera vez: crea las listas (requiere permiso para administrar el sitio) y carga los datos iniciales.
  for (const k of missing) {
    const l = await g(`/sites/${S.site.id}/lists`, { method: "POST", json: { displayName: LISTAS[k].nombre, columns: LISTAS[k].cols, list: { template: "genericList" } } });
    S.lists[k] = l.id;
    if (k === "terrenos") for (const n of TERRENOS_INICIALES) await addItem("terrenos", { Title: n });
    if (k === "preguntas") { let o = 1; for (const q of PREGUNTAS_INICIALES) await addItem("preguntas", { Title: q.texto, Tipo: q.tipo, Lo: q.lo || "", Hi: q.hi || "", Req: true, Activa: true, Orden: o++ }); }
    if (k === "admins") await addItem("admins", { Title: S.me.email });
  }
}
const addItem = (k, fields) => g(`/sites/${S.site.id}/lists/${S.lists[k]}/items`, { method: "POST", json: { fields } });
const setItem = (k, id, fields) => g(`/sites/${S.site.id}/lists/${S.lists[k]}/items/${id}/fields`, { method: "PATCH", json: fields });
const delItem = (k, id) => g(`/sites/${S.site.id}/lists/${S.lists[k]}/items/${id}`, { method: "DELETE" });
const listItems = k => gAll(`/sites/${S.site.id}/lists/${S.lists[k]}/items?expand=fields&$top=500`);

async function loadAll() {
  const [t, p, r, a] = await Promise.all(["terrenos", "preguntas", "reportes", "admins"].map(listItems));
  S.terrenos = t.map(i => ({ id: i.id, nombre: i.fields.Title || "" })).filter(x => x.nombre).sort((x, y) => x.nombre.localeCompare(y.nombre, "es"));
  S.preguntas = p.map(i => ({ id: i.id, texto: i.fields.Title || "", tipo: i.fields.Tipo || "scale", lo: i.fields.Lo || "", hi: i.fields.Hi || "", req: i.fields.Req !== false, activa: i.fields.Activa !== false, orden: Number(i.fields.Orden || 0) }))
    .sort((x, y) => x.orden - y.orden);
  S.reportes = r.map(i => {
    const f = i.fields; let d = {}; try { d = JSON.parse(f.Datos || "{}"); } catch (e) { }
    return { id: i.id, nombre: f.Title, terrenoId: f.TerrenoId, terrenoNombre: f.Terreno, fecha: f.Fecha, autorNombre: f.Autor, autorEmail: f.AutorEmail, nFotos: Number(f.NFotos || 0), carpeta: f.Carpeta, obs: f.Obs || "", preguntas: d.preguntas || [], respuestas: d.respuestas || {} };
  }).filter(x => x.fecha).sort((x, y) => y.fecha.localeCompare(x.fecha));
  S.admins = a.map(i => ({ id: i.id, email: i.fields.Title || "" })).filter(x => x.email);
  cacheSave();
}

/* ---------- sincronización ---------- */
async function sync(manual) {
  if (S.syncing) return; S.syncing = true;
  try {
    if (!navigator.onLine) { if (manual) toast("Sin señal. Se sincroniza solo al volver la conexión."); return; }
    if (!S.me) { const m = await g("/me?$select=displayName,mail,userPrincipalName"); S.me = { name: m.displayName, email: m.mail || m.userPrincipalName }; }
    await resolveSite();
    await ensureLists();
    await flushOutbox();
    await loadAll();
    $("errNote").hidden = true; S.needLogin = false; $("authNote").hidden = true;
  } catch (e) {
    console.error(e);
    if (e.login) { /* aviso ya visible */ }
    else if (e.status === 403) showErr("Tu cuenta no tiene permiso en el sitio de SharePoint de la app. Pide a un administrador que te agregue como miembro del sitio.");
    else if (e.status === 404) showErr("No se encontró el sitio de SharePoint configurado. Revisa la dirección en m365-config.js.");
    else if (manual) toast("No se pudo sincronizar. Se reintentará solo.");
  } finally { S.syncing = false; render(); }
}
function showErr(t) { $("errNote").textContent = t; $("errNote").hidden = false; }
addEventListener("online", () => sync());
setInterval(() => { if (navigator.onLine && S.ready) sync(); }, 60000);
document.addEventListener("visibilitychange", () => { if (!document.hidden && S.ready) sync(); });

async function flushOutbox() {
  const items = (await kvAll("outbox")).sort((a, b) => a.fecha.localeCompare(b.fecha));
  let sent = 0;
  for (const rep of items) {
    const base = repName(rep);
    if (!rep.carpeta) {
      for (let n = 1; n < 30; n++) {
        const name = n === 1 ? base : `${base} (${n})`;
        const path = drivePath("Reportes", rep.terrenoNombre, name);
        if (!(await driveItem(path))) { rep.carpeta = path; rep.nombre = name; break; }
      }
      await kvPut("outbox", rep.id, rep);
    }
    for (let i = 0; i < rep.fotos.length; i++) {
      if (rep.fotos[i].subida) continue;
      await uploadFile(`${rep.carpeta}/foto-${String(i + 1).padStart(2, "0")}.jpg`, dataUrlToBlob(rep.fotos[i].data), "image/jpeg");
      rep.fotos[i].subida = true; await kvPut("outbox", rep.id, rep);
    }
    if (!rep.pdfSubido) {
      const pdf = await buildPdf(rep, rep.fotos);
      await uploadFile(`${rep.carpeta}/${safeName(rep.nombre)}.pdf`, pdf, "application/pdf");
      rep.pdfSubido = true; await kvPut("outbox", rep.id, rep);
    }
    await addItem("reportes", {
      Title: rep.nombre, TerrenoId: String(rep.terrenoId), Terreno: rep.terrenoNombre, Fecha: rep.fecha,
      Autor: rep.autorNombre, AutorEmail: rep.autorEmail, NFotos: rep.fotos.length, Carpeta: rep.carpeta, Obs: rep.obs,
      Datos: JSON.stringify({ preguntas: rep.preguntas, respuestas: rep.respuestas, inicio: rep.inicio })
    });
    await kvDel("outbox", rep.id); sent++;
  }
  S.outbox = await kvAll("outbox");
  if (sent) toast(sent > 1 ? `${sent} reportes enviados a SharePoint.` : "Reporte enviado a SharePoint.");
}

/* ---------- arranque ---------- */
(async () => {
  cacheLoad();
  S.outbox = await kvAll("outbox");
  try {
    await pca.initialize();
    const res = await pca.handleRedirectPromise();
    if (res && res.account) pca.setActiveAccount(res.account);
  } catch (e) { console.error(e); $("loginErr").textContent = "No se pudo completar el inicio de sesión. Inténtalo de nuevo."; }
  const acc = account();
  if (!acc) { $("onboard").hidden = false; return; }
  pca.setActiveAccount(acc);
  if (!S.me || lower(S.me.email) !== lower(acc.username)) {
    S.me = { name: acc.name || acc.username, email: acc.username };
    if (S.site && S.reportes.length && !S.reportes.some(r => lower(r.autorEmail) === lower(acc.username))) { /* misma organización: se mantiene el caché */ }
  }
  S.ready = true; render(); route();
  sync();
})();

/* ---------- navegación ---------- */
function route() {
  if (!S.ready) return;
  const h = location.hash.replace("#", "");
  if (h === "admin" && isAdmin()) return showAdmin();
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
  $("actionbar").hidden = !(v === "vLot" && !$("pVisit").hidden);
}
function showHome() { S.cur = null; setView("vHome"); $("brand").textContent = "Administración Predial"; render(); }
function render() {
  $("whoName").textContent = (S.me && S.me.name) || "—";
  $("adminBtn").hidden = !isAdmin();
  if (!$("vHome").hidden) renderHome();
  if (S.cur) { updateTabCount(); if (!$("pReps").hidden) renderReps(); if (!$("pVisit").hidden && !$("survey").childElementCount) renderSurvey(); }
  if (!$("vAdmin").hidden) { renderQAdmin(); renderUAdmin(); }
}

/* ---------- inicio ---------- */
function repsOf(tid) {
  const pend = S.outbox.filter(o => String(o.terrenoId) === String(tid)).map(o => ({ ...o, _pend: true }));
  return [...pend, ...S.reportes.filter(r => String(r.terrenoId) === String(tid))].sort((a, b) => b.fecha.localeCompare(a.fecha));
}
function renderHome() {
  const box = $("lots"); box.textContent = "";
  if (!S.terrenos.length) { const d = document.createElement("div"); d.className = "empty"; d.textContent = navigator.onLine ? "Cargando terrenos desde SharePoint…" : "Sin señal. Abre la app con conexión una vez para cargar los terrenos."; box.append(d); }
  for (const t of S.terrenos) {
    const reps = repsOf(t.id); const pend = reps.filter(r => r._pend).length;
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
  const n = S.outbox.length;
  $("pendNote").hidden = !n;
  $("pendNote").textContent = n ? `${n} reporte${n > 1 ? "s" : ""} guardado${n > 1 ? "s" : ""} en este teléfono. Se enviará${n > 1 ? "n" : ""} a SharePoint cuando haya señal.` : "";
}
$("addForm").onsubmit = async e => {
  e.preventDefault(); const name = $("newLot").value.trim(); if (!name) return;
  if (S.terrenos.some(t => lower(t.nombre) === lower(name))) { toast("Ese terreno ya existe."); return; }
  if (!navigator.onLine) { toast("Necesitas señal para agregar un terreno."); return; }
  try { const it = await addItem("terrenos", { Title: name }); S.terrenos.push({ id: it.id, nombre: name }); S.terrenos.sort((x, y) => x.nombre.localeCompare(y.nombre, "es")); cacheSave(); $("newLot").value = ""; toast("Terreno agregado: " + name); render(); }
  catch (err) { toast("No se pudo agregar el terreno."); }
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
    if (!S.cur) return; await kvPut("drafts", String(S.cur.id), S.draft);
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
    }
    box.append(d);
  });
  if (!box.childElementCount) { const e = document.createElement("p"); e.className = "hint"; e.textContent = "Las preguntas se cargan desde SharePoint la primera vez que abres la app con señal."; box.append(e); }
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
  const miss = activeQuestions().filter(q => q.req !== false && (S.draft.respuestas[String(q.id)] == null || S.draft.respuestas[String(q.id)] === ""));
  document.querySelectorAll(".q").forEach(el => el.classList.toggle("missing", miss.some(q => String(q.id) === el.dataset.k)));
  if (miss.length) {
    toast(`Falta responder ${miss.length} pregunta${miss.length > 1 ? "s" : ""} del reporte.`);
    document.querySelector(".q.missing")?.scrollIntoView({ behavior: "smooth", block: "center" }); return;
  }
  $("finish").hidden = true; $("confirm").hidden = false;
};
$("cancelSend").onclick = resetConfirm;
$("doSend").onclick = async () => {
  const lot = S.cur, d = S.draft;
  const preguntas = activeQuestions().map(q => ({ id: String(q.id), texto: q.texto, tipo: q.tipo, lo: q.lo || "", hi: q.hi || "" }));
  const respuestas = {}; for (const q of preguntas) if (d.respuestas[q.id] != null && d.respuestas[q.id] !== "") respuestas[q.id] = d.respuestas[q.id];
  const rep = {
    id: uid(), terrenoId: String(lot.id), terrenoNombre: lot.nombre, fecha: new Date().toISOString(), inicio: d.inicio,
    autorNombre: (S.me && S.me.name) || "", autorEmail: (S.me && S.me.email) || "",
    preguntas, respuestas, obs: (d.obs || "").trim(), fotos: d.fotos.map(f => ({ data: f.data, w: f.w, h: f.h })), nFotos: d.fotos.length
  };
  await kvPut("outbox", rep.id, rep); await kvDel("drafts", String(lot.id));
  S.outbox = await kvAll("outbox");
  S.draft = { respuestas: {}, obs: "", fotos: [], inicio: new Date().toISOString() };
  $("obs").value = ""; renderSurvey(); renderThumbs(); resetConfirm();
  $("savedTxt").textContent = "El borrador se guarda en este teléfono.";
  toast(navigator.onLine ? "Enviando reporte a SharePoint…" : "Sin señal: el reporte quedó guardado y se enviará solo.");
  setTab("reps"); sync();
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
  if (!reps.length) { const e = document.createElement("div"); e.className = "empty"; e.textContent = "Todavía no hay reportes de este terreno."; box.append(e); return; }
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
    if (r._pend) { const t = document.createElement("span"); t.className = "tag"; t.textContent = navigator.onLine ? "Enviando…" : "Pendiente de envío"; c.querySelector(".rep-h").append(t); }
    const row = document.createElement("div"); row.className = "share"; row.style.gridTemplateColumns = "1fr";
    const btn = document.createElement("button"); btn.className = "btn"; btn.textContent = "Compartir / descargar PDF";
    btn.onclick = () => prepareShare(r, btn); row.append(btn); c.append(row);
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
  (r.preguntas || []).forEach((q, i) => L.push(`${i + 1}. ${q.texto}: ${ansLong(q, r.respuestas?.[q.id])}`));
  L.push("", "Observaciones: " + (r.obs || "Sin observaciones."), "Fotos: " + (r.nFotos || 0));
  return L.join("\n");
}
async function prepareShare(r, btn) {
  const old = btn.textContent; btn.disabled = true; btn.textContent = r._pend ? "Generando PDF…" : "Descargando PDF…";
  try {
    const name = r.nombre || repName(r);
    let blob, webUrl = null;
    if (r._pend) blob = await buildPdf(r, r.fotos);
    else {
      const it = await driveItem(`${r.carpeta}/${safeName(name)}.pdf`);
      if (!it) throw new Error("PDF no encontrado en SharePoint");
      blob = await (await fetch(it["@microsoft.graph.downloadUrl"])).blob();
      try { const folder = await driveItem(r.carpeta); webUrl = folder && folder.webUrl; } catch (e) { }
    }
    const file = new File([blob], name + ".pdf", { type: "application/pdf" });
    S.share = { r, file, blob, name };
    const canShare = NATIVE || !!(navigator.canShare && navigator.canShare({ files: [file] }));
    $("shNative").hidden = !canShare;
    $("shT").textContent = name;
    $("shSub").textContent = canShare ? "Elige cómo enviarlo. Al compartir, el PDF va adjunto." : "Este teléfono no permite adjuntar el PDF directamente: descárgalo y compártelo desde tus archivos.";
    $("shDownload").textContent = NATIVE ? "Guardar PDF en el teléfono" : "Descargar PDF";
    $("shWa").href = "https://wa.me/?text=" + encodeURIComponent(shareText(r));
    $("shSp").hidden = !webUrl; if (webUrl) $("shSp").href = webUrl;
    $("shSheet").hidden = false;
  } catch (e) {
    console.error(e);
    toast(navigator.onLine ? "No se pudo obtener el PDF desde SharePoint." : "Sin señal: el PDF de un reporte ya enviado se descarga desde SharePoint.");
  } finally { btn.disabled = false; btn.textContent = old; }
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
    if (!/cancel|abort/i.test(msg)) toast("No se pudo abrir el menú de compartir. Usa Guardar PDF.");
  }
};
$("shDownload").onclick = async () => {
  const s = S.share; if (!s) return;
  if (NATIVE) {
    try {
      await Filesystem.writeFile({ path: "Predial HyC/" + s.file.name, data: await blobToBase64(s.blob), directory: Directory.Documents, recursive: true });
      toast("PDF guardado en Documentos › Predial HyC.");
    } catch (e) { toast("No se pudo guardar en el teléfono. Usa Compartir y elige Guardar en Drive o OneDrive."); }
    return;
  }
  const url = URL.createObjectURL(s.blob); const a = document.createElement("a");
  a.href = url; a.download = s.file.name; document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
};
$("shClose").onclick = () => { $("shSheet").hidden = true; };
$("shSheet").onclick = e => { if (e.target === $("shSheet")) $("shSheet").hidden = true; };

/* ---------- administración ---------- */
function showAdmin() { S.cur = null; setView("vAdmin"); $("brand").textContent = "Administración"; setAdminTab("q"); }
function setAdminTab(t) {
  $("tabQ").setAttribute("aria-selected", t === "q"); $("tabU").setAttribute("aria-selected", t === "u");
  $("pQ").hidden = t !== "q"; $("pU").hidden = t !== "u";
  renderQAdmin(); renderUAdmin();
}
$("tabQ").onclick = () => setAdminTab("q"); $("tabU").onclick = () => setAdminTab("u");
function adminErr(e) { console.error(e); toast(navigator.onLine ? "No se pudo guardar en SharePoint." : "Necesitas señal para editar."); }
async function afterAdminChange(msg) { await loadAll().catch(() => { }); render(); if (msg) toast(msg); }

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
    mk(q.activa === false ? "Activar" : "Desactivar", "Activar o desactivar", () => setItem("preguntas", q.id, { Activa: q.activa === false }).then(() => afterAdminChange()).catch(adminErr));
    box.append(r);
  });
}
async function moveQ(i, dir) {
  const a = S.preguntas[i], b = S.preguntas[i + dir]; if (!a || !b) return;
  try {
    const oa = a.orden || i + 1, ob = b.orden || i + dir + 1;
    await setItem("preguntas", a.id, { Orden: ob === oa ? oa + dir : ob });
    await setItem("preguntas", b.id, { Orden: oa });
    await afterAdminChange();
  } catch (e) { adminErr(e); }
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
$("qForm").onsubmit = async e => {
  e.preventDefault();
  const f = { Title: $("qText").value.trim(), Tipo: $("qType").value, Lo: $("qLo").value.trim(), Hi: $("qHi").value.trim(), Req: $("qReq").checked };
  if (!f.Title) return;
  $("qSheet").hidden = true;
  try {
    if (editingQ) await setItem("preguntas", editingQ.id, f);
    else await addItem("preguntas", { ...f, Activa: true, Orden: Math.max(0, ...S.preguntas.map(q => q.orden || 0)) + 1 });
    await afterAdminChange("Pregunta guardada en SharePoint.");
  } catch (err) { adminErr(err); }
};

function renderUAdmin() {
  const box = $("uList"); box.textContent = "";
  for (const a of S.admins.slice().sort((x, y) => x.email.localeCompare(y.email))) {
    const r = document.createElement("div"); r.className = "user-row";
    const n = document.createElement("div"); n.textContent = a.email + (S.me && lower(a.email) === lower(S.me.email) ? " (tú)" : "");
    r.append(n);
    if (!(S.me && lower(a.email) === lower(S.me.email))) {
      const b = document.createElement("button"); b.className = "btn small quiet"; b.textContent = "Quitar";
      b.onclick = () => delItem("admins", a.id).then(() => afterAdminChange("Administrador quitado.")).catch(adminErr);
      r.append(b);
    }
    box.append(r);
  }
}
$("addAdmin").onsubmit = async e => {
  e.preventDefault(); const em = lower($("adminEmail").value);
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) { toast("Escribe un correo válido."); return; }
  if (S.admins.some(a => lower(a.email) === em)) { toast("Ya es administrador."); return; }
  try { await addItem("admins", { Title: em }); $("adminEmail").value = ""; await afterAdminChange("Administrador agregado."); } catch (err) { adminErr(err); }
};

/* ---------- versión web (sin APK): instalación y funcionamiento sin señal ---------- */
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
if (!NATIVE && "serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => { });
