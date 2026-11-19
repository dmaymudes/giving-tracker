"use strict";
/* Family Giving Tracker: a static page over a Google Sheet.
   Data model (tabs): Gifts, Organizations, Config. Rows are found by Gift ID, never by row number. */
const CFG = window.GT_CONFIG;
const API = "https://sheets.googleapis.com/v4/spreadsheets/" + CFG.spreadsheetId;
const DEMO = new URLSearchParams(location.search).has("demo");
const $ = (s, el) => (el || document).querySelector(s);
const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const money = n => (n ? "$" + Math.round(n).toLocaleString("en-US") : "–");
const pct = x => (x * 100).toFixed(1) + "%";
const num = v => (typeof v === "number" ? v : parseFloat(String(v || "").replace(/[$,]/g, "")) || 0);
const today = () => { const d = new Date(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); };
const colLetter = i => String.fromCharCode(65 + i);

const S = { token: null, gifts: [], orgs: {}, orgList: [], lists: {}, targets: {}, gcol: {}, year: new Date().getFullYear(), filters: { loc: "", veh: "", who: "", q: "" }, showAll: false, canEdit: true };

/* ---------- messages ---------- */
function say(text, err) { $("#msg").innerHTML = text ? `<div class="${err ? "err" : ""}">${esc(text)}</div>` : ""; }

/* ---------- auth ---------- */
let tokenClient;
function signIn() {
  if (DEMO) return start();
  if (!window.google || !google.accounts) return say("Google sign-in has not loaded yet. Try again in a moment.", true);
  tokenClient = tokenClient || google.accounts.oauth2.initTokenClient({
    client_id: CFG.clientId, scope: CFG.scope,
    callback: r => {
      if (r.error) return say("Sign-in failed: " + r.error, true);
      S.token = r.access_token;
      try { sessionStorage.setItem("gt_tok", JSON.stringify({ t: r.access_token, x: Date.now() + (r.expires_in - 60) * 1000 })); } catch (e) {}
      start();
    }
  });
  tokenClient.requestAccessToken({ prompt: "" });
}
function restoreToken() {
  try { const o = JSON.parse(sessionStorage.getItem("gt_tok") || "null"); if (o && o.x > Date.now()) { S.token = o.t; return true; } } catch (e) {}
  return false;
}
async function api(path, opts) {
  const r = await fetch(API + path, Object.assign({}, opts, { headers: { Authorization: "Bearer " + S.token, "Content-Type": "application/json" } }));
  if (r.status === 401) { S.token = null; try { sessionStorage.removeItem("gt_tok"); } catch (e) {} throw new Error("Your sign-in expired. Reload the page and sign in again."); }
  if (r.status === 403 || r.status === 404) { const e = new Error("This Google account can't do that on the sheet. View-only access can read but not change it."); e.status = r.status; throw e; }
  if (!r.ok) throw new Error("Sheets request failed (" + r.status + "): " + (await r.text()).slice(0, 200));
  return r.json();
}

/* ---------- load ---------- */
async function fetchTables() {
  if (DEMO) return demoTables();
  const q = ["Gifts!A1:P", "Organizations!A1:P", "Config!A1:Z"].map(r => "ranges=" + encodeURIComponent(r)).join("&");
  const j = await api("/values:batchGet?" + q + "&valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING");
  return j.valueRanges.map(v => v.values || []);
}
function ingest(tables) {
  const [g, o, c] = tables;
  const oh = o[0] || []; S.orgs = {}; S.orgList = [];
  o.slice(1).forEach(r => {
    const x = {}; oh.forEach((h, i) => (x[h] = r[i] == null ? "" : r[i]));
    if (x["Organization"]) { S.orgs[x["Organization"]] = x; S.orgList.push(x["Organization"]); }
  });
  S.orgList.sort((a, b) => a.localeCompare(b, "en", { sensitivity: "base" }));
  const gh = g[0] || []; S.gcol = {}; gh.forEach((h, i) => (S.gcol[h] = i));
  S.gifts = [];
  g.slice(1).forEach((r, i) => {
    const v = h => (r[S.gcol[h]] == null ? "" : r[S.gcol[h]]);
    if (!v("Organization")) return;
    const org = S.orgs[v("Organization")] || {};
    S.gifts.push({ id: String(v("Gift ID")), row: i + 2, year: num(v("Year")), org: v("Organization"),
      loc: org["Location"] || v("Location") || "", cat: org["Category"] || v("Category") || "", who: org["Directed by"] || v("Directed by") || "",
      proposed: num(v("Proposed ($)")), decision: v("Decision"), amount: num(v("Amount ($)")), vehicle: v("Vehicle"),
      pay: v("Payment status"), method: v("Method"), ref: String(v("Reference")), date: String(v("Date sent")), notes: v("Notes") });
  });
  S.nextRow = g.length + 1;
  for (let i = 1; i < g.length; i++) if (!(g[i][S.gcol["Organization"]])) { S.nextRow = i + 1; break; }
  const ch = c[0] || []; S.lists = {}; S.targets = {};
  const tcol = ch.indexOf("Soft target");
  ch.forEach((h, i) => { if (h && (tcol < 0 || i < tcol)) S.lists[h] = c.slice(1).map(r => r[i]).filter(x => x !== "" && x != null); });
  readTargets();
}
function readTargets() {
  const c = S.tables[2], ch = c[0] || [], ti = ch.indexOf("Soft target"); S.targets = {};
  if (ti >= 0) { const yi = ch.findIndex((h, i) => i > ti && String(h) === String(S.year)); const col = yi >= 0 ? yi : ti + 1;
    c.slice(1).forEach(r => { if (r[ti] && typeof r[col] === "number") S.targets[r[ti]] = r[col]; }); }
}
/* With the drive.file scope the page has no access to the sheet until the user grants it once through
   Google's picker. The picker is opened already showing only the giving sheet. */
function pickSheet() {
  return new Promise((resolve, reject) => {
    if (!window.gapi) return reject(new Error("Google's file picker has not loaded yet. Try again in a moment."));
    gapi.load("picker", () => {
      const view = new google.picker.DocsView(google.picker.ViewId.SPREADSHEETS).setFileIds(CFG.spreadsheetId).setMode(google.picker.DocsViewMode.LIST);
      new google.picker.PickerBuilder().addView(view).setOAuthToken(S.token).setDeveloperKey(CFG.apiKey).setAppId(CFG.appId)
        .setTitle("Select the Family Giving Tracker sheet to give this page access")
        .setCallback(d => {
          if (d.action === google.picker.Action.PICKED) return d.docs.some(x => x.id === CFG.spreadsheetId) ? resolve() : reject(new Error("That was a different file. Choose the Family Giving Tracker sheet."));
          if (d.action === google.picker.Action.CANCEL) reject(new Error("No sheet was selected. If the list was empty, this Google account has not been given access to the giving sheet; ask David to share it with you."));
        }).build().setVisible(true);
    });
  });
}
async function load() {
  say("Loading…");
  try {
    try { S.tables = await fetchTables(); }
    catch (e) { if (DEMO || !e.status) throw e; say("Choose the giving sheet in the Google window to continue."); await pickSheet(); say("Loading…"); S.tables = await fetchTables(); }
    ingest(S.tables); say("");
  } catch (e) {
    say(e.status ? "This Google account can't open the giving sheet. Ask David to share it with you." : e.message, true);
    if (!S.gifts.length) $("#main").innerHTML = `<div id="signin"><p>The giving sheet isn't open yet.</p><button id="retry" class="primary">Try again</button></div>`;
    throw e;
  }
}
async function start() {
  if ($("#signin")) $("#signin").hidden = true;
  try { await load(); } catch (e) { return; }
  const years = [...new Set(S.gifts.map(g => g.year))].sort((a, b) => b - a);
  if (!years.includes(S.year)) S.year = years[0];
  $("#year").innerHTML = years.map(y => `<option${y === S.year ? " selected" : ""}>${y}</option>`).join("");
  $("#nav").hidden = $("#controls").hidden = false;
  $("#sheetlink").href = "https://docs.google.com/spreadsheets/d/" + CFG.spreadsheetId + "/edit";
  if (!location.hash) location.hash = "#decide"; else render();
}

/* ---------- writes ---------- */
const FIELD = { proposed: "Proposed ($)", decision: "Decision", amount: "Amount ($)", vehicle: "Vehicle", pay: "Payment status", method: "Method", ref: "Reference", date: "Date sent", notes: "Notes" };
async function saveGifts(changes) { // [{gift, fields:{key:value}}]
  if (!changes.length) return;
  if (!DEMO) {
    const ids = (await api("/values/" + encodeURIComponent("Gifts!A1:A") + "?valueRenderOption=UNFORMATTED_VALUE")).values.map(r => String(r[0] || ""));
    const data = [];
    for (const ch of changes) {
      const idx = ids.indexOf(ch.gift.id);
      if (idx < 1) throw new Error("Gift " + ch.gift.id + " was not found in the sheet. Reload and try again.");
      ch.gift.row = idx + 1;
      for (const k in ch.fields) data.push({ range: "Gifts!" + colLetter(S.gcol[FIELD[k]]) + ch.gift.row, values: [[ch.fields[k] === null ? "" : ch.fields[k]]] });
    }
    await api("/values:batchUpdate", { method: "POST", body: JSON.stringify({ valueInputOption: "USER_ENTERED", data }) });
  }
  changes.forEach(ch => Object.assign(ch.gift, Object.fromEntries(Object.entries(ch.fields).map(([k, v]) => [k, v === null ? (["proposed", "amount"].includes(k) ? 0 : "") : v]))));
}
async function act(fn) { try { say("Saving…"); await fn(); say(""); render(); } catch (e) { say(e.message, true); } }
async function addGift(org, amount, vehicle) {
  const maxId = S.gifts.reduce((m, g) => Math.max(m, parseInt(g.id.replace(/\D/g, ""), 10) || 0), 0);
  const id = "G" + String(maxId + 1).padStart(4, "0");
  const g = { id, year: S.year, org, loc: (S.orgs[org] || {})["Location"] || "", cat: (S.orgs[org] || {})["Category"] || "", who: (S.orgs[org] || {})["Directed by"] || "",
    proposed: amount, decision: "Proposed", amount: 0, vehicle, pay: "", method: "", ref: "", date: "", notes: "" };
  if (!DEMO) {
    const col = (await api("/values/" + encodeURIComponent("Gifts!C1:C") + "?valueRenderOption=UNFORMATTED_VALUE")).values || [];
    let row = col.length + 1;
    for (let i = 1; i < col.length; i++) if (!col[i][0]) { row = i + 1; break; }
    const c = h => colLetter(S.gcol[h]);
    const data = [["Gift ID", id], ["Year", S.year], ["Organization", org], ["Proposed ($)", amount], ["Decision", "Proposed"], ["Vehicle", vehicle]]
      .map(([h, v]) => ({ range: "Gifts!" + c(h) + row, values: [[v]] }));
    await api("/values:batchUpdate", { method: "POST", body: JSON.stringify({ valueInputOption: "USER_ENTERED", data }) });
    g.row = row;
  }
  S.gifts.push(g);
}

/* ---------- derived ---------- */
const yearGifts = () => S.gifts.filter(g => g.year === S.year);
const lastYearTotal = org => S.gifts.filter(g => g.org === org && g.year === S.year - 1 && g.decision === "Approved").reduce((s, g) => s + g.amount, 0);
const pipeline = g => (g.decision === "Approved" ? g.amount : g.decision === "Proposed" ? g.proposed : 0);
const isSent = g => g.pay === "Sent" || g.pay === "Confirmed";
function passes(g) { const f = S.filters; return (!f.loc || g.loc === f.loc) && (!f.veh || g.vehicle === f.veh) && (!f.who || g.who === f.who) && (!f.q || g.org.toLowerCase().includes(f.q.toLowerCase())); }
function filterBar(extra) {
  const f = S.filters, sel = (k, label, opts) => `<select data-f="${k}"><option value="">${label}: all</option>${opts.map(o => `<option${f[k] === o ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
  return `<div class="bar">${sel("loc", "Location", S.lists["Location"] || [])}${sel("veh", "Vehicle", S.lists["Vehicle"] || [])}${sel("who", "Directed by", S.lists["Directed by"] || [])}
    <input data-f="q" placeholder="Search organization" value="${esc(f.q)}"><span class="sp"></span>${extra || ""}</div>`;
}
const vehOpts = cur => (S.lists["Vehicle"] || []).map(v => `<option${v === cur ? " selected" : ""}>${esc(v)}</option>`).join("");
const orgLink = n => `<a class="orglink" data-org="${esc(n)}">${esc(n)}</a>`;

/* ---------- views ---------- */
function viewDecide() {
  const all = yearGifts().filter(passes).filter(g => g.loc !== "Overhead");
  const rows = all.filter(g => S.showAll || g.decision === "Proposed").sort((a, b) => a.loc.localeCompare(b.loc) || b.proposed - a.proposed || a.org.localeCompare(b.org));
  const open = all.filter(g => g.decision === "Proposed");
  const sum = a => a.reduce((s, g) => s + pipeline(g), 0);
  let h = filterBar(`<label><input type="checkbox" id="showall"${S.showAll ? " checked" : ""}> Show decided too</label>
    <button id="approveall" class="ok"${open.length ? "" : " disabled"}>Approve all ${open.length} shown at proposed amount</button>`);
  h += `<div class="card"><h2>Decisions to make <small>${open.length} undecided, ${money(sum(open))} · approved so far ${money(sum(all.filter(g => g.decision === "Approved")))}</small></h2>`;
  if (!rows.length) h += `<div class="empty">Nothing to decide with these filters.</div>`;
  else {
    h += `<table><tr><th>Organization</th><th>Location</th><th>Directed by</th><th class="n">${S.year - 1} total</th><th class="n">Proposed</th><th>Vehicle</th><th>Decision</th><th></th></tr>`;
    rows.forEach(g => {
      const done = g.decision !== "Proposed";
      h += `<tr data-id="${esc(g.id)}"><td class="org">${orgLink(g.org)}</td><td>${esc(g.loc)}</td><td>${esc(g.who)}</td><td class="n">${money(lastYearTotal(g.org))}</td>
        <td class="n">${done ? money(g.decision === "Approved" ? g.amount : g.proposed) : `<input class="amt" data-a="proposed" value="${g.proposed || ""}" inputmode="decimal">`}</td>
        <td>${done ? esc(g.vehicle) : `<select data-a="vehicle">${vehOpts(g.vehicle)}</select>`}</td>
        <td><span class="tag ${esc(g.decision)}">${esc(g.decision)}</span></td>
        <td>${done ? `<button data-a="reopen">Reopen</button>` : `<button data-a="approve" class="ok">Approve</button> <button data-a="defer">Defer</button> <button data-a="decline" class="bad">Decline</button>`}</td></tr>`;
    });
    h += `</table>`;
  }
  h += `</div><div class="card"><h2>Add a gift for ${S.year}</h2><div class="bar" style="padding:12px;margin:0">
    <input id="neworg" list="orglist" placeholder="Organization" size="36"><datalist id="orglist">${S.orgList.map(o => `<option value="${esc(o)}">`).join("")}</datalist>
    <input id="newamt" class="amt" placeholder="Amount" inputmode="decimal"><select id="newveh">${vehOpts("")}</select><button id="addgift" class="primary">Add as proposed</button>
    <small style="color:var(--mute)">New organizations are added on the Organizations tab of the sheet first.</small></div></div>`;
  return h;
}
function viewExecute() {
  const all = yearGifts().filter(passes);
  const todo = all.filter(g => g.decision === "Approved" && !isSent(g)), sent = all.filter(g => g.decision === "Approved" && g.pay === "Sent");
  let h = filterBar();
  if (!todo.length) h += `<div class="card"><div class="empty">No approved gifts are waiting to be sent.</div></div>`;
  (S.lists["Vehicle"] || []).concat([""]).forEach(v => {
    const rows = todo.filter(g => g.vehicle === v).sort((a, b) => a.org.localeCompare(b.org));
    if (!rows.length) return;
    h += `<div class="card"><h2>${esc(v || "No vehicle set")} <small>${rows.length} to send, ${money(rows.reduce((s, g) => s + g.amount, 0))}</small></h2>
      <table><tr><th>Organization</th><th class="n">Amount</th><th>Address</th><th>Method</th><th>Reference</th><th>Date sent</th><th></th></tr>`;
    rows.forEach(g => {
      const o = S.orgs[g.org] || {}, addr = [o["Address line 1"], o["Address line 2"], [o["City"], o["State"], o["ZIP"]].filter(Boolean).join(" ")].filter(Boolean).join(", ");
      h += `<tr data-id="${esc(g.id)}"><td class="org">${orgLink(g.org)}</td><td class="n">${money(g.amount)}</td><td class="addr">${esc(addr)}${o["EIN"] ? " · EIN " + esc(o["EIN"]) : ""}</td>
        <td><select data-x="method"><option></option>${(S.lists["Method"] || []).map(m => `<option${m === g.method ? " selected" : ""}>${esc(m)}</option>`).join("")}</select></td>
        <td><input class="ref" data-x="ref" value="${esc(g.ref)}" placeholder="check no."></td><td><input type="date" data-x="date" value="${esc(g.date || today())}"></td>
        <td><button data-a="sent" class="ok">Mark sent</button></td></tr>`;
    });
    h += `</table></div>`;
  });
  if (sent.length) {
    h += `<div class="card"><h2>Sent, not yet confirmed <small>${sent.length} gifts, ${money(sent.reduce((s, g) => s + g.amount, 0))}</small></h2><table><tr><th>Organization</th><th class="n">Amount</th><th>Vehicle</th><th>Method</th><th>Reference</th><th>Date sent</th><th></th></tr>`;
    sent.sort((a, b) => a.date.localeCompare(b.date)).forEach(g => (h += `<tr data-id="${esc(g.id)}"><td class="org">${orgLink(g.org)}</td><td class="n">${money(g.amount)}</td><td>${esc(g.vehicle)}</td><td>${esc(g.method)}</td><td>${esc(g.ref)}</td><td>${esc(g.date)}</td>
      <td><button data-a="confirm" class="ok">Confirm</button> <button data-a="unsend">Undo</button></td></tr>`));
    h += `</table></div>`;
  }
  return h;
}
function viewBudget() {
  const gs = yearGifts(), giving = gs.filter(g => g.loc !== "Overhead");
  const agg = (rows) => ({ prop: rows.filter(g => g.decision === "Proposed").reduce((s, g) => s + g.proposed, 0), appr: rows.filter(g => g.decision === "Approved").reduce((s, g) => s + g.amount, 0),
    sent: rows.filter(g => g.decision === "Approved" && isSent(g)).reduce((s, g) => s + g.amount, 0), n: rows.filter(g => g.decision === "Approved" || g.decision === "Proposed").length });
  const T = agg(giving), total = T.prop + T.appr, t = S.targets;
  const tTotal = t["Total giving, excluding overhead ($)"], tSea = t["Seattle, maximum share of total"], tCape = t["Cape Cod, minimum share of total"], tSeaN = t["Seattle, maximum number of gifts"], tSmall = t["Seattle, maximum gifts of $2,500 or less"];
  const table = (title, keys, field, withTargets) => {
    let h = `<div class="card"><h2>${title}</h2><table><tr><th></th><th class="n">Still proposed</th><th class="n">Approved</th><th class="n">Sent</th><th class="n">Approved + proposed</th><th class="n">Gifts</th><th class="n">Share</th>${withTargets ? "<th>Soft target</th>" : ""}</tr>`;
    keys.forEach(k => {
      const a = agg(giving.filter(g => g[field] === k)), tot = a.prop + a.appr, share = total ? tot / total : 0; if (!tot && !a.n) return;
      let tg = "";
      if (withTargets && k === "Seattle" && tSea != null) tg = `<span class="${share > tSea ? "over" : "fine"}">at most ${pct(tSea)}</span>${tSeaN != null ? ` · <span class="${a.n > tSeaN ? "over" : "fine"}">at most ${tSeaN} gifts</span>` : ""}`;
      if (withTargets && k === "Cape Cod" && tCape != null) tg = `<span class="${share < tCape ? "under" : "fine"}">at least ${pct(tCape)}</span>`;
      h += `<tr><td>${esc(k || "(none)")}</td><td class="n">${money(a.prop)}</td><td class="n">${money(a.appr)}</td><td class="n">${money(a.sent)}</td><td class="n">${money(tot)}</td><td class="n">${a.n}</td><td class="n">${pct(share)}</td>${withTargets ? `<td>${tg}</td>` : ""}</tr>`;
    });
    h += `<tr class="tot"><td>Total</td><td class="n">${money(T.prop)}</td><td class="n">${money(T.appr)}</td><td class="n">${money(T.sent)}</td><td class="n">${money(total)}</td><td class="n">${T.n}</td><td></td>${withTargets ? `<td>${tTotal ? `${money(tTotal)} target, ${money(Math.abs(tTotal - total))} ${total > tTotal ? "over" : "under"}` : ""}</td>` : ""}</tr></table></div>`;
    return h;
  };
  const live = giving.filter(g => (g.decision === "Approved" || g.decision === "Proposed") && pipeline(g) > 0);
  const tierOf = a => CFG.tiers.find(x => a >= x);
  let pyr = `<div class="card"><h2>Giving pyramid <small>approved + proposed, ${S.year} against ${S.year - 1}</small></h2><table><tr><th>Gift size</th><th class="n">Gifts</th><th class="n">Total</th><th></th><th class="n">${S.year - 1} gifts</th><th class="n">${S.year - 1} total</th></tr>`;
  const prev = S.gifts.filter(g => g.year === S.year - 1 && g.decision === "Approved" && g.loc !== "Overhead" && g.amount > 0);
  const maxTier = Math.max(1, ...CFG.tiers.map(x => live.filter(g => tierOf(pipeline(g)) === x).reduce((s, g) => s + pipeline(g), 0)));
  CFG.tiers.forEach((x, i) => {
    const a = live.filter(g => tierOf(pipeline(g)) === x), p = prev.filter(g => tierOf(g.amount) === x), sum = a.reduce((s, g) => s + pipeline(g), 0);
    const label = i === 0 ? money(x) + " and up" : x === 0 ? "Under " + money(CFG.tiers[i - 1]) : money(x) + " to " + money(CFG.tiers[i - 1] - 1);
    pyr += `<tr><td>${label}</td><td class="n">${a.length}</td><td class="n">${money(sum)}</td><td><div class="meter"><i style="width:${(100 * sum / maxTier).toFixed(0)}%"></i></div></td><td class="n">${p.length}</td><td class="n">${money(p.reduce((s, g) => s + g.amount, 0))}</td></tr>`;
  });
  pyr += `</table></div>`;
  const smallSea = live.filter(g => g.loc === "Seattle" && pipeline(g) <= 2500).length;
  const oh = gs.filter(g => g.loc === "Overhead" && g.decision === "Approved").reduce((s, g) => s + g.amount, 0);
  const note = `<div class="bar"><span>Seattle gifts of $2,500 or less: <b class="${tSmall != null && smallSea > tSmall ? "over" : "fine"}">${smallSea}</b>${tSmall != null ? " (soft cap " + tSmall + ")" : ""}</span><span>Overhead paid: <b>${money(oh)}</b></span><span style="color:var(--mute)">Targets are soft and are edited on the Config tab of the sheet.</span></div>`;
  return note + table("By location", S.lists["Location"] || [], "loc", true) + `<div class="grid2">` + table("By vehicle", S.lists["Vehicle"] || [], "vehicle") + table("By who directs", S.lists["Directed by"] || [], "who") + `</div>` + pyr;
}
function viewHistory() {
  const years = [...new Set(S.gifts.map(g => g.year))].sort(), f = S.filters, by = {};
  S.gifts.filter(g => g.decision === "Approved" && (!f.loc || g.loc === f.loc) && (!f.who || g.who === f.who) && (!f.veh || g.vehicle === f.veh) && (!f.q || g.org.toLowerCase().includes(f.q.toLowerCase())))
    .forEach(g => { const o = (by[g.org] = by[g.org] || { loc: g.loc, tot: 0 }); o[g.year] = (o[g.year] || 0) + g.amount; o.tot += g.amount; });
  const names = Object.keys(by).sort((a, b) => by[a].loc.localeCompare(by[b].loc) || by[b].tot - by[a].tot);
  let h = filterBar() + `<div class="card"><h2>Gifts by organization and year <small>${names.length} organizations</small></h2><table><tr><th>Organization</th><th>Location</th>${years.map(y => `<th class="n">${y}</th>`).join("")}<th class="n">Total</th></tr>`;
  names.forEach(n => (h += `<tr><td class="org">${orgLink(n)}</td><td>${esc(by[n].loc)}</td>${years.map(y => `<td class="n">${money(by[n][y])}</td>`).join("")}<td class="n"><b>${money(by[n].tot)}</b></td></tr>`));
  h += `<tr class="tot"><td>Total</td><td></td>${years.map(y => `<td class="n">${money(names.reduce((s, n) => s + (by[n][y] || 0), 0))}</td>`).join("")}<td class="n">${money(names.reduce((s, n) => s + by[n].tot, 0))}</td></tr></table></div>`;
  return h;
}
function viewCharts() {
  const live = yearGifts().filter(g => g.loc !== "Overhead" && (g.decision === "Approved" || g.decision === "Proposed"));
  const total = live.reduce((s, g) => s + pipeline(g), 0);
  const tip = t => ` data-tip="${esc(t)}"`;
  const breakdown = (title, field) => {
    const m = {}; live.forEach(g => { const k = g[field] || "(none)"; (m[k] = m[k] || { v: 0, n: 0 }); m[k].v += pipeline(g); m[k].n++; });
    const keys = Object.keys(m).filter(k => m[k].v > 0).sort((a, b) => m[b].v - m[a].v), max = Math.max(1, ...keys.map(k => m[k].v));
    return `<div class="card"><h2>${title}</h2><div class="hb">` + keys.map(k => `<div class="hbrow"${tip(`${k}: ${money(m[k].v)}, ${pct(total ? m[k].v / total : 0)} of giving, ${m[k].n} gifts`)}>
      <span class="hbl">${esc(k)}</span><span class="hbt"><i style="width:${(100 * m[k].v / max).toFixed(1)}%"></i></span><span class="hbv">${money(m[k].v)}</span><span class="hbp">${pct(total ? m[k].v / total : 0)}</span></div>`).join("") + `</div></div>`;
  };
  const years = [...new Set(S.gifts.map(g => g.year))].sort();
  const counts = g => g.loc !== "Overhead" && (g.decision === "Approved" || (g.year >= new Date().getFullYear() && g.decision === "Proposed"));
  // stacked columns by year; colour follows the entity (its position in keys), never its rank
  const stacked = (keys, field, label) => {
    const val = (y, k) => S.gifts.filter(g => g.year === y && counts(g) && (k === "Other" && !keys.includes(g[field]) || g[field] === k)).reduce((s, g) => s + pipeline(g), 0);
    const tot = years.map(y => keys.reduce((s, k) => s + val(y, k), 0)), top = Math.max(1, ...tot), step = top > 400000 ? 100000 : top > 150000 ? 50000 : 25000, ymax = Math.ceil(top / step) * step;
    const W = 760, H = 300, L = 56, B = 26, T = 18, pw = W - L - 10, ph = H - B - T, bw = Math.min(64, pw / years.length * 0.6);
    const col = (k, j) => (k === "Other" ? "#98a2b3" : SERIES[j % SERIES.length]);
    let svg = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(label)}">`;
    for (let v = 0; v <= ymax; v += step) { const y = T + ph - ph * v / ymax; svg += `<line x1="${L}" x2="${W - 10}" y1="${y}" y2="${y}" stroke="#e6e9ef"/><text x="${L - 6}" y="${y + 4}" text-anchor="end" class="ax">${v ? "$" + v / 1000 + "K" : "0"}</text>`; }
    years.forEach((y, i) => {
      const cx = L + pw * (i + 0.5) / years.length; let base = 0;
      keys.forEach((k, j) => { const v = val(y, k); if (!v) return; const h = ph * v / ymax, y0 = T + ph - ph * base / ymax - h;
        svg += `<rect x="${cx - bw / 2}" y="${y0 + 1}" width="${bw}" height="${Math.max(0, h - 2)}" rx="2" fill="${col(k, j)}"${tip(`${y} ${k}: ${money(v)}, ${pct(v / tot[i])} of the year`)}/>`; base += v; });
      svg += `<text x="${cx}" y="${T + ph - ph * tot[i] / ymax - 5}" text-anchor="middle" class="ax b">${money(tot[i])}</text><text x="${cx}" y="${H - 8}" text-anchor="middle" class="ax">${y}${y >= new Date().getFullYear() ? " plan" : ""}</text>`;
    });
    return `<div class="legend">${keys.map((k, j) => `<span><i style="background:${col(k, j)}"></i>${esc(k)}</span>`).join("")}</div>${svg}</svg>`;
  };
  const locs = (S.lists["Location"] || []).filter(l => l !== "Overhead");
  const catTot = {}; S.gifts.filter(g => g.loc !== "Overhead" && g.decision === "Approved").forEach(g => (catTot[g.cat] = (catTot[g.cat] || 0) + g.amount));
  const topCats = Object.keys(catTot).sort((a, b) => catTot[b] - catTot[a]).slice(0, 7);
  const cats = (S.lists["Category"] || []).filter(c => topCats.includes(c)).concat(Object.keys(catTot).length > 7 ? ["Other"] : []);
  // change against last year, per organization
  const cur = {}, prev = {};
  live.forEach(g => (cur[g.org] = (cur[g.org] || 0) + pipeline(g)));
  S.gifts.filter(g => g.year === S.year - 1 && g.decision === "Approved" && g.loc !== "Overhead").forEach(g => (prev[g.org] = (prev[g.org] || 0) + g.amount));
  const ch = [...new Set(Object.keys(cur).concat(Object.keys(prev)))].map(o => ({ o, c: cur[o] || 0, p: prev[o] || 0, d: (cur[o] || 0) - (prev[o] || 0) })).filter(x => x.c || x.p);
  const kind = x => (!x.p ? "New" : !x.c ? "Dropped" : x.d > 0 ? "Increased" : x.d < 0 ? "Decreased" : "Unchanged");
  const kc = {}; ch.forEach(x => { const k = kind(x); (kc[k] = kc[k] || { n: 0, d: 0 }); kc[k].n++; kc[k].d += x.d; });
  const moved = ch.filter(x => x.d).sort((a, b) => b.d - a.d), dmax = Math.max(1, ...moved.map(x => Math.abs(x.d)));
  const changes = `<div class="tiles">${["New", "Increased", "Unchanged", "Decreased", "Dropped"].map(k => `<div class="tile"><b>${(kc[k] || { n: 0 }).n}</b><span>${k}</span><small>${kc[k] && kc[k].d ? (kc[k].d > 0 ? "+" : "−") + money(Math.abs(kc[k].d)) : "&nbsp;"}</small></div>`).join("")}</div>` +
    (moved.length ? `<div class="hb dv">${moved.map(x => `<div class="dvrow"${tip(`${x.o}: ${money(x.p)} in ${S.year - 1}, ${money(x.c)} in ${S.year} (${kind(x).toLowerCase()})`)}><span class="hbl">${orgLink(x.o)}</span>
      <span class="dvt"><i class="${x.d > 0 ? "up" : "dn"}" style="width:${(50 * Math.abs(x.d) / dmax).toFixed(1)}%"></i></span><span class="hbv">${x.d > 0 ? "+" : "−"}${money(Math.abs(x.d))}</span><span class="hbp">${kind(x)}</span></div>`).join("")}</div>` : `<div class="empty">No organization's amount differs from ${S.year - 1}.</div>`);
  // cumulative sent by month, this year against last year and the target
  const cum = y => { const m = Array(12).fill(0); S.gifts.filter(g => g.year === y && g.decision === "Approved" && g.loc !== "Overhead" && isSent(g) && /^\d{4}-\d{2}/.test(g.date)).forEach(g => (m[+g.date.slice(5, 7) - 1] += g.amount)); let t = 0; return m.map(v => (t += v)); };
  const cy = cum(S.year), py = cum(S.year - 1), target = S.targets["Total giving, excluding overhead ($)"] || 0, apprTot = live.filter(g => g.decision === "Approved").reduce((s, g) => s + g.amount, 0);
  const nowM = S.year < new Date().getFullYear() ? 11 : S.year > new Date().getFullYear() ? -1 : new Date().getMonth();
  const progress = (() => {
    const W = 760, H = 280, L = 56, R = 120, B = 26, T = 14, pw = W - L - R, ph = H - B - T, top = Math.max(1, target, apprTot, cy[11], py[11]), step = top > 400000 ? 100000 : top > 150000 ? 50000 : 25000, ymax = Math.ceil(top / step) * step;
    const X = i => L + pw * i / 11, Y = v => T + ph - ph * v / ymax;
    let g = `<div class="legend"><span><i style="background:#2a78d6"></i>${S.year}</span><span><i style="background:#98a2b3"></i>${S.year - 1}</span></div><svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Cumulative giving sent by month">`;
    for (let v = 0; v <= ymax; v += step) g += `<line x1="${L}" x2="${W - R}" y1="${Y(v)}" y2="${Y(v)}" stroke="#e6e9ef"/><text x="${L - 6}" y="${Y(v) + 4}" text-anchor="end" class="ax">${v ? "$" + v / 1000 + "K" : "0"}</text>`;
    MN.forEach((m, i) => (g += `<text x="${X(i)}" y="${H - 8}" text-anchor="middle" class="ax">${m}</text>`));
    const ref = (v, label, dash) => (v ? `<line x1="${L}" x2="${W - R}" y1="${Y(v)}" y2="${Y(v)}" stroke="#667085" stroke-width="1.5" stroke-dasharray="${dash}"/><text x="${W - R + 6}" y="${Y(v) + 4}" class="ax b">${label} ${money(v)}</text>` : "");
    g += ref(target, "Target", "6 4");
    if (apprTot) g += ref(apprTot, "Approved", "2 3").replace(`y="${Y(apprTot) + 4}" class="ax b"`, `y="${Y(apprTot) + 4 + (target && Math.abs(Y(apprTot) - Y(target)) < 14 ? (apprTot > target ? -14 : 14) : 0)}" class="ax b"`);
    const line = (arr, n, color, name, yr) => { if (n < 0) return ""; const pts = arr.slice(0, n + 1);
      return `<polyline fill="none" stroke="${color}" stroke-width="2" stroke-linejoin="round" points="${pts.map((v, i) => X(i) + "," + Y(v)).join(" ")}"/>` +
        pts.map((v, i) => `<circle cx="${X(i)}" cy="${Y(v)}" r="9" fill="transparent"${tip(`${MN[i]} ${yr}: ${money(v)} sent so far`)}/><circle cx="${X(i)}" cy="${Y(v)}" r="3" fill="${color}" stroke="#fff" stroke-width="1.5" pointer-events="none"/>`).join("") +
        ""; };
    return g + line(py, 11, "#98a2b3", String(S.year - 1), S.year - 1) + line(cy, nowM, "#2a78d6", String(S.year), S.year) + `</svg>`;
  })();
  // largest recipients over all years in the sheet
  const allTot = {}; S.gifts.filter(g => g.decision === "Approved" && g.loc !== "Overhead").forEach(g => (allTot[g.org] = (allTot[g.org] || 0) + g.amount));
  const grand = Object.values(allTot).reduce((s, v) => s + v, 0), top15 = Object.keys(allTot).sort((a, b) => allTot[b] - allTot[a]).slice(0, 15), tmax = Math.max(1, ...top15.map(o => allTot[o]));
  const largest = `<div class="hb wide">${top15.map(o => `<div class="hbrow"${tip(`${o}: ${money(allTot[o])} over ${years[0]} to ${years[years.length - 1]}, ${pct(allTot[o] / grand)} of all giving`)}><span class="hbl">${orgLink(o)}</span><span class="hbt"><i style="width:${(100 * allTot[o] / tmax).toFixed(1)}%"></i></span><span class="hbv">${money(allTot[o])}</span><span class="hbp">${pct(allTot[o] / grand)}</span></div>`).join("")}</div>`;
  // when the money goes out, for the selected year
  const sentM = Array(12).fill(0), cntM = Array(12).fill(0);
  yearGifts().filter(g => g.decision === "Approved" && g.loc !== "Overhead" && /^\d{4}-\d{2}/.test(g.date)).forEach(g => { const m = +g.date.slice(5, 7) - 1; sentM[m] += g.amount; cntM[m]++; });
  const mmax = Math.max(1, ...sentM);
  const months = `<div class="cols">${sentM.map((v, i) => `<div class="col"${tip(`${MN[i]} ${S.year}: ${money(v)} in ${cntM[i]} gifts`)}><b>${v ? "$" + Math.round(v / 1000) + "K" : ""}</b><i style="height:${(100 * v / mmax).toFixed(1)}%"></i><span>${MN[i]}</span></div>`).join("")}</div>`;
  return `<div class="bar"><span>${S.year}: <b>${money(total)}</b> approved or proposed across ${live.length} gifts, excluding overhead.</span></div>
    <div class="grid2">${breakdown("By category", "cat")}${breakdown("By location", "loc")}${breakdown("By vehicle", "vehicle")}${breakdown("By who directs", "who")}</div>
    <div class="card"><h2>Changes against ${S.year - 1} <small>by organization, approved + proposed</small></h2>${changes}</div>
    <div class="card"><h2>Progress through ${S.year} <small>cumulative dollars sent, by date sent</small></h2>${progress}</div>
    <div class="card"><h2>Giving by year and location <small>approved gifts; years marked "plan" also count proposals</small></h2>${stacked(locs, "loc", "Giving by year and location")}</div>
    <div class="card"><h2>Giving by year and category <small>the seven largest categories, with the rest as Other</small></h2>${stacked(cats, "cat", "Giving by year and category")}</div>
    <div class="card"><h2>When gifts were sent in ${S.year} <small>by date sent</small></h2>${months}</div>
    <div class="card"><h2>Largest recipients, ${years[0]} to ${years[years.length - 1]} <small>approved gifts, share of all giving</small></h2>${largest}</div>`;
}
const SERIES = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6"];
const MN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function viewOrg(name) {
  const o = S.orgs[name] || {}, rows = S.gifts.filter(g => g.org === name).sort((a, b) => b.year - a.year || b.date.localeCompare(a.date));
  let h = `<div class="bar"><button onclick="history.back()">Back</button></div><div class="card"><h2>${esc(name)} <small>${esc([o["Location"], o["Category"], o["Directed by"] && "directed by " + o["Directed by"], o["Org type"]].filter(Boolean).join(" · "))}</small></h2>
    <table><tr><th>Year</th><th>Decision</th><th class="n">Proposed</th><th class="n">Amount</th><th>Vehicle</th><th>Payment</th><th>Method</th><th>Reference</th><th>Date sent</th><th>Notes</th></tr>`;
  rows.forEach(g => (h += `<tr><td>${g.year}</td><td><span class="tag ${esc(g.decision)}">${esc(g.decision)}</span></td><td class="n">${money(g.proposed)}</td><td class="n">${money(g.amount)}</td><td>${esc(g.vehicle)}</td><td>${esc(g.pay)}</td><td>${esc(g.method)}</td><td>${esc(g.ref)}</td><td>${esc(g.date)}</td><td>${esc(g.notes)}</td></tr>`));
  return h + `<tr class="tot"><td>Total given</td><td></td><td></td><td class="n">${money(rows.filter(g => g.decision === "Approved").reduce((s, g) => s + g.amount, 0))}</td><td colspan="6"></td></tr></table></div>`;
}

/* ---------- render and events ---------- */
function render() {
  if (!S.gifts.length && !S.tables) return;
  const hash = location.hash || "#decide", view = hash.split("/")[0];
  document.querySelectorAll("nav a").forEach(a => a.classList.toggle("on", a.getAttribute("href") === view));
  const m = $("#main");
  if (view === "#org") m.innerHTML = viewOrg(decodeURIComponent(hash.slice(5)));
  else m.innerHTML = ({ "#decide": viewDecide, "#execute": viewExecute, "#budget": viewBudget, "#charts": viewCharts, "#history": viewHistory }[view] || viewDecide)();
  applySorts();
}
/* ---------- click a column header to sort; the choice is kept per screen and table ---------- */
S.sort = {};
const sortKey = t => (location.hash || "#decide").split("/")[0] + ":" + [...document.querySelectorAll("#main table")].indexOf(t);
function cellVal(td) {
  if (!td) return "";
  const inp = td.querySelector("input:not([type=checkbox]),select"), t = (inp ? inp.value : td.textContent).trim();
  if (td.classList.contains("n") || /^[−+-]?\$?[\d,]+(\.\d+)?%?$/.test(t)) { const n = parseFloat(t.replace(/−/g, "-").replace(/[$,%+]/g, "")); return isNaN(n) ? 0 : n; }
  return t.toLowerCase();
}
function applySorts() {
  document.querySelectorAll("#main table").forEach(t => {
    const st = S.sort[sortKey(t)], head = t.querySelector("tr"); if (!head || !head.querySelector("th")) return;
    [...head.children].forEach((th, i) => { th.classList.add("sortable"); th.dataset.dir = st && st.col === i ? st.dir : ""; });
    if (!st) return;
    const body = head.parentNode, rows = [...body.children].filter(r => r !== head && !r.classList.contains("tot")), tot = [...body.children].filter(r => r.classList.contains("tot"));
    rows.sort((x, y) => { const a = cellVal(x.children[st.col]), b = cellVal(y.children[st.col]); return (typeof a === "number" && typeof b === "number" ? a - b : String(a).localeCompare(String(b))) * (st.dir === "asc" ? 1 : -1); });
    rows.concat(tot).forEach(r => body.appendChild(r));
  });
}
document.addEventListener("click", e => {
  const th = e.target.closest && e.target.closest("#main th"); if (!th) return;
  const t = th.closest("table"), k = sortKey(t), col = [...th.parentNode.children].indexOf(th), cur = S.sort[k];
  const numeric = th.classList.contains("n");
  S.sort[k] = { col, dir: cur && cur.col === col ? (cur.dir === "asc" ? "desc" : "asc") : numeric ? "desc" : "asc" };
  applySorts();
});
const giftOf = el => { const tr = el.closest("tr[data-id]"); return tr && S.gifts.find(g => g.id === tr.dataset.id); };
document.addEventListener("click", e => {
  const t = e.target;
  if (t.id === "signinbtn") return signIn();
  if (t.id === "retry") return S.token ? start() : signIn();
  if (t.id === "reload") return load().then(render).catch(() => {});
  if (t.dataset.org) return (location.hash = "#org/" + encodeURIComponent(t.dataset.org));
  if (t.id === "approveall") {
    const open = yearGifts().filter(passes).filter(g => g.decision === "Proposed" && g.loc !== "Overhead");
    if (!open.length || !confirm("Approve " + open.length + " gifts at their proposed amounts?")) return;
    return act(() => saveGifts(open.map(g => ({ gift: g, fields: { decision: "Approved", amount: g.proposed, pay: "To send" } }))));
  }
  if (t.id === "addgift") {
    const org = $("#neworg").value.trim(), amt = num($("#newamt").value), veh = $("#newveh").value;
    if (!S.orgs[org]) return say("Pick an organization from the list. Add new ones on the Organizations tab of the sheet first, then press Reload.", true);
    return act(() => addGift(org, amt, veh));
  }
  const a = t.dataset.a, g = a && giftOf(t); if (!g) return;
  const tr = t.closest("tr"), x = k => { const el = $(`[data-x="${k}"]`, tr); return el ? el.value : ""; };
  const F = { approve: { decision: "Approved", amount: g.proposed, pay: "To send" }, defer: { decision: "Deferred" }, decline: { decision: "Declined" },
    reopen: { decision: "Proposed", amount: null, pay: null, proposed: g.proposed || g.amount }, sent: { pay: "Sent", method: x("method"), ref: x("ref"), date: x("date") }, confirm: { pay: "Confirmed" }, unsend: { pay: "To send" } }[a];
  if (a === "reopen" && isSent(g)) return say("This gift is already marked sent. Change it in the sheet if that is wrong.", true);
  if (F) act(() => saveGifts([{ gift: g, fields: F }]));
});
document.addEventListener("change", e => {
  const t = e.target;
  if (t.id === "year") { S.year = +t.value; readTargets(); return render(); }
  if (t.id === "showall") { S.showAll = t.checked; return render(); }
  if (t.dataset.f) { S.filters[t.dataset.f] = t.value; return render(); }
  const g = t.dataset.a && giftOf(t); if (!g) return;
  if (t.dataset.a === "proposed") act(() => saveGifts([{ gift: g, fields: { proposed: num(t.value) } }]));
  if (t.dataset.a === "vehicle") act(() => saveGifts([{ gift: g, fields: { vehicle: t.value } }]));
});
window.addEventListener("hashchange", render);
window.addEventListener("load", () => { if (DEMO || restoreToken()) start(); });

/* ---------- demo data (?demo) : invented organizations, nothing is saved ---------- */
function demoTables() {
  if (window.GT_DEMO_TABLES) return window.GT_DEMO_TABLES; // test hook
  const orgs = [["Org ID", "Organization", "Location", "Category", "Directed by", "Org type", "EIN", "Usual vehicle", "Country", "Address line 1", "Address line 2", "City", "State", "ZIP"],
    ["O001", "Example Food Bank", "Seattle", "Basic Needs", "D&E", "501c3", "", "Foundation", "USA", "1 Main St", "", "Seattle", "WA", "98101"],
    ["O002", "Sample Land Trust", "Cape Cod", "Conservation/Climate", "All", "501c3", "", "Foundation"], ["O003", "Demo Library Friends", "Cape Cod", "Community Building", "Sue", "501c3", "", "Fidelity"],
    ["O004", "Placeholder Scholars", "Seattle", "Educational Equity", "Emily", "501c3", "", "DAF"], ["O005", "Test Voting Project", "National", "Democracy", "David", "501c3", "", "DAF"]];
  const H = ["Gift ID", "Year", "Organization", "Location", "Category", "Directed by", "Org total last year ($)", "Proposed ($)", "Decision", "Amount ($)", "Vehicle", "Payment status", "Method", "Reference", "Date sent", "Notes"];
  const gifts = [H]; let n = 1; const y = new Date().getFullYear();
  orgs.slice(1).forEach((o, i) => { const amt = [10000, 15000, 2500, 5000, 1000][i];
    gifts.push(["G" + String(n++).padStart(4, "0"), y - 1, o[1], "", "", "", "", "", "Approved", amt, o[7], "Confirmed", "Check", "10" + i, (y - 1) + "-12-10", ""]);
    gifts.push(["G" + String(n++).padStart(4, "0"), y, o[1], "", "", "", "", amt, i === 0 ? "Approved" : "Proposed", i === 0 ? amt : "", o[7], i === 0 ? "To send" : "", "", "", "", ""]); });
  const cfg = [["Location", "Vehicle", "Category", "Directed by", "Org type", "Decision", "Payment status", "Method", "", "Soft target", String(y)],
    ["Seattle", "Foundation", "Basic Needs", "David", "501c3", "Proposed", "To send", "Check", "", "Total giving, excluding overhead ($)", 40000],
    ["Cape Cod", "DAF", "Democracy", "Emily", "501c4", "Approved", "Sent", "eCheck", "", "Seattle, maximum share of total", 0.4],
    ["National", "Fidelity", "", "D&E", "", "Declined", "Confirmed", "Card", "", "Cape Cod, minimum share of total", 0.4],
    ["Overhead", "Cape Cod Foundation", "", "Sue", "", "Deferred", "", "Online grant", "", "Seattle, maximum number of gifts", 35], ["", "", "", "All", "", "", "", "", "", "Seattle, maximum gifts of $2,500 or less", 15]];
  return [gifts, orgs, cfg];
}

/* ---------- hover tooltip for charts ---------- */
const tipEl = document.createElement("div"); tipEl.id = "tip"; tipEl.hidden = true; document.body.appendChild(tipEl);
document.addEventListener("mousemove", e => {
  const t = e.target.closest && e.target.closest("[data-tip]");
  if (!t) return (tipEl.hidden = true);
  tipEl.textContent = t.dataset.tip; tipEl.hidden = false;
  tipEl.style.left = Math.min(e.clientX + 12, innerWidth - tipEl.offsetWidth - 8) + "px"; tipEl.style.top = e.clientY + 14 + "px";
});
