/*
 * Kiser Invoice add-in — Input Form (LDP / LCP / Change Order / Payment Request)
 * Ported from the Office Script "Run Input Form" to Office.js (Excel.run).
 * The form lives in this task pane (not in a sheet) so multiple users never collide.
 */

/* global console, document, Excel, Office */

const FMT_ACCT = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';
const FMT_USD = '"$"#,##0.00';

interface Item { desc: string; amt: number; isHdr: boolean; }
interface InputFormData {
  dateVal: number | "";
  secType: string;
  colName: string;
  secHeader: string;
  vtTotal: number | "";
  vtCost: number | "";
  vtHours: number | "";
  vtPO: string;
  items: Item[];
}

const CL = (i: number): string =>
  i < 26
    ? String.fromCharCode(65 + i)
    : String.fromCharCode(64 + Math.floor(i / 26)) + String.fromCharCode(65 + (i % 26));

// Legacy-webview-safe helpers (older Excel webviews lack NodeList.forEach and
// classList.toggle(force))
function qsa(sel: string): Element[] {
  return Array.prototype.slice.call(document.querySelectorAll(sel));
}
function toggleCls(el: Element, cls: string, on: boolean) {
  if (on) el.classList.add(cls);
  else el.classList.remove(cls);
}

// ───────────────────────── Form (task pane) ─────────────────────────

let poLoaded = false;
// Tracks whether the line-items table was auto-filled by getPR() (Payment
// Request flow). If the user then switches Section Type to LDP / LCP /
// Change Order, those auto-loaded rows must be cleared — otherwise stale
// PR line items would be saved against the wrong section type.
let prItemsAutoLoaded = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    (document.getElementById("sideload-msg") as HTMLElement).style.display = "none";
    // 'flex' (not 'block') so the column layout in CSS engages — pins the top bar
    // and tabs above a scrollable form area.
    (document.getElementById("app-body") as HTMLElement).style.display = "flex";

    for (let i = 0; i < 4; i++) addItemRow();
    for (let i = 0; i < 4; i++) addPORow();

    qsa(".tab").forEach((t) => {
      (t as HTMLElement).onclick = () => switchTab((t as HTMLElement).getAttribute("data-tab") as string);
    });

    (document.getElementById("addRow") as HTMLElement).onclick = () => addItemRow();
    (document.getElementById("generate") as HTMLElement).onclick = generate;
    (document.getElementById("invoiceGen") as HTMLElement).onclick = invoiceGen;

    // Help overlay — opens fullscreen Help & Reference panel
    (document.getElementById("helpBtn") as HTMLElement).onclick = () => {
      (document.getElementById("helpOverlay") as HTMLElement).style.display = "";
    };
    (document.getElementById("helpClose") as HTMLElement).onclick = () => {
      (document.getElementById("helpOverlay") as HTMLElement).style.display = "none";
    };
    // Collapsible help sections
    qsa(".hsec-head").forEach((h) => {
      (h as HTMLElement).onclick = () => {
        const tgt = (h as HTMLElement).getAttribute("data-target");
        if (!tgt) return;
        const body = document.getElementById(tgt) as HTMLElement | null;
        const arrow = (h as HTMLElement).querySelector(".hsec-arrow") as HTMLElement | null;
        if (!body) return;
        const open = body.style.display !== "none";
        body.style.display = open ? "none" : "";
        if (arrow) arrow.textContent = open ? "▶" : "▼";
      };
    });

    // Auto-load PR lines when user picks "Payment Request" in Section Type.
    // When switching to anything else, hide the "Latest PR" hint AND clear the
    // auto-loaded line items so they don't bleed into an LDP / LCP / CO save.
    (document.getElementById("secType") as HTMLSelectElement).onchange = (ev) => {
      const v = (ev.target as HTMLSelectElement).value;
      if (v === "Payment Request") {
        void getPR();
      } else {
        const hint = document.getElementById("latestPRHint") as HTMLElement;
        if (hint) hint.style.display = "none";
        if (prItemsAutoLoaded) {
          const body = document.getElementById("itemsBody") as HTMLTableSectionElement;
          body.innerHTML = "";
          for (let i = 0; i < 4; i++) addItemRow();
          prItemsAutoLoaded = false;
        }
      }
    };

    (document.getElementById("addPORow") as HTMLElement).onclick = () => addPORow();
    (document.getElementById("loadEst") as HTMLElement).onclick = loadEst;
    (document.getElementById("updatePo") as HTMLElement).onclick = updatePo;
    (document.getElementById("movePo") as HTMLElement).onclick = updatePo;
    (document.getElementById("loadMove") as HTMLElement).onclick = loadMove;
    (document.getElementById("contractAdj") as HTMLElement).onclick = contractAdjust;

    // Accordions (expand / collapse)
    qsa(".acc-head").forEach((h) => {
      (h as HTMLElement).onclick = () => {
        const acc = h.parentElement as HTMLElement;
        toggleCls(acc, "open", acc.className.indexOf("open") === -1);
      };
    });

    // Section Header toggle — show/hide the optional input
    const sht = document.getElementById("secHeaderTgl") as HTMLInputElement;
    const shi = document.getElementById("secHeader") as HTMLInputElement;
    const syncSH = () => { shi.style.display = sht.checked ? "block" : "none"; if (!sht.checked) shi.value = ""; };
    sht.onchange = syncSH;
    syncSH();
  }
});

function switchTab(name: string) {
  qsa(".tab").forEach((t) => toggleCls(t, "active", (t as HTMLElement).getAttribute("data-tab") === name));
  qsa(".tabpane").forEach((p) => toggleCls(p, "active", p.id === "tab-" + name));
  hideStatus();
  if (name === "po" && !poLoaded) { poLoaded = true; loadEst(); }
}

function hideStatus() {
  const el = document.getElementById("status") as HTMLElement;
  el.className = "status";
  el.textContent = "";
}

async function loadMove() {
  const source = (document.getElementById("poSource") as HTMLSelectElement).value.trim();
  if (!source) { setStatus("Select a Source Contract first.", "err"); return; }
  setStatus("Loading POs to move…", "busy"); setBusy(true);
  try {
    let rows: { vendor: string; poNum: string; amount: number; adj: number }[] = [];
    await Excel.run(async (context) => { rows = await runLoadMove(context, source); });
    const body = document.getElementById("poBody") as HTMLTableSectionElement;
    body.innerHTML = "";
    for (const r of rows) {
      addPORow();
      const tr = body.rows[body.rows.length - 1];
      (tr.querySelector(".po-vendor") as HTMLInputElement).value = r.vendor;
      (tr.querySelector(".po-num") as HTMLInputElement).value = r.poNum;
      (tr.querySelector(".po-amt") as HTMLInputElement).value = r.amount ? String(r.amount) : "";
      (tr.querySelector(".po-adj") as HTMLInputElement).value = r.adj ? String(r.adj) : "";
    }
    for (let i = 0; i < 2; i++) addPORow();
    setStatus(`Loaded ${rows.length} PO(s). Tick the ones to move, choose a Target, then Run Move Po.`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("ERROR: " + errMsg(e), "err");
  } finally { setBusy(false); }
}

async function contractAdjust() {
  const poNumber = (document.getElementById("poSource") as HTMLSelectElement).value.trim();
  const amtRaw = (document.getElementById("poAdjAmt") as HTMLInputElement).value.trim();
  const desc = (document.getElementById("poAdjDesc") as HTMLInputElement).value.trim();
  if (!poNumber) { setStatus("Select the Contract / PO Number (Source) first.", "err"); return; }
  if (desc === "") { setStatus("Enter a Contract Description.", "err"); return; }
  if (amtRaw === "" || isNaN(Number(amtRaw))) { setStatus("Enter a valid Contract Adjust amount.", "err"); return; }
  setStatus("Applying contract adjustment…", "busy"); setBusy(true);
  try {
    await Excel.run(async (context) => { await runContractAdjust(context, poNumber, Number(amtRaw), desc); });
    setStatus("Contract adjustment applied (saved as negative).", "ok");
    (document.getElementById("poAdjAmt") as HTMLInputElement).value = "";
    (document.getElementById("poAdjDesc") as HTMLInputElement).value = "";
  } catch (e) {
    console.error(e);
    setStatus("ERROR: " + errMsg(e), "err");
  } finally { setBusy(false); }
}

function addItemRow() {
  const body = document.getElementById("itemsBody") as HTMLTableSectionElement;
  const n = body.rows.length + 1;
  const tr = body.insertRow();
  tr.innerHTML =
    `<td class="numcol">${n}</td>` +
    `<td><input class="desc" type="text" /></td>` +
    `<td><input class="amt" type="number" step="any" /></td>`;
  wireAutoGrow(tr, body, addItemRow);
}

// When the user types in the last row, append a fresh empty row automatically.
function wireAutoGrow(tr: HTMLTableRowElement, body: HTMLTableSectionElement, addFn: () => void) {
  const fields = Array.prototype.slice.call(tr.querySelectorAll("input[type=text], input[type=number]"));
  for (let i = 0; i < fields.length; i++) {
    (fields[i] as HTMLInputElement).oninput = () => {
      if (tr !== body.rows[body.rows.length - 1]) return;
      const has = fields.some((f: HTMLInputElement) => f.value.trim() !== "");
      if (has) addFn();
    };
  }
}

function readForm(): InputFormData {
  const val = (id: string) => (document.getElementById(id) as HTMLInputElement).value.trim();
  const num = (id: string): number | "" => {
    const v = val(id);
    return v === "" || isNaN(Number(v)) ? "" : Number(v);
  };

  const dateStr = val("date");
  let dateVal: number | "" = "";
  if (dateStr !== "") dateVal = toExcelSerial(dateStr);

  const items: Item[] = [];
  qsa("#itemsBody tr").forEach((tr) => {
    const desc = (tr.querySelector(".desc") as HTMLInputElement).value.trim();
    if (desc === "") return;
    const amtRaw = (tr.querySelector(".amt") as HTMLInputElement).value.trim();
    const amt = amtRaw === "" || isNaN(Number(amtRaw)) ? 0 : Number(amtRaw);
    items.push({ desc, amt, isHdr: amt === 0 });
  });

  return {
    dateVal,
    secType: val("secType"),
    colName: val("colName"),
    secHeader: val("secHeader"),
    vtTotal: num("vtTotal"),
    vtCost: num("vtCost"),
    vtHours: num("vtHours"),
    vtPO: val("vtPO"),
    items,
  };
}

function toExcelSerial(iso: string): number {
  const d = new Date(iso + "T00:00:00Z");
  const epoch = Date.UTC(1899, 11, 30);
  return Math.round((d.getTime() - epoch) / 86400000);
}

function setStatus(msg: string, kind: "ok" | "err" | "busy") {
  const el = document.getElementById("status") as HTMLElement;
  el.className = "status " + kind;
  if (kind === "busy") {
    el.innerHTML = "";
    const s = document.createElement("span");
    s.className = "spin";
    el.appendChild(s);
    el.appendChild(document.createTextNode(msg));
  } else {
    el.textContent = msg;
  }
}

// Disable/enable all run buttons while a script is running
function setBusy(on: boolean) {
  const btns = qsa(".btn");
  for (let i = 0; i < btns.length; i++) (btns[i] as HTMLButtonElement).disabled = on;
}

function errMsg(e: unknown): string {
  return e && (e as Error).message ? (e as Error).message : String(e);
}

function clearFormUI() {
  ["date", "colName", "secHeader", "vtTotal", "vtCost", "vtHours", "vtPO"].forEach(
    (id) => ((document.getElementById(id) as HTMLInputElement).value = "")
  );
  (document.getElementById("secType") as HTMLSelectElement).value = "";
  (document.getElementById("itemsBody") as HTMLTableSectionElement).innerHTML = "";
  const hint = document.getElementById("latestPRHint") as HTMLElement;
  if (hint) hint.style.display = "none";
  prItemsAutoLoaded = false;
  for (let i = 0; i < 4; i++) addItemRow();
}

async function generate() {
  const form = readForm();
  if (!form.secType) { setStatus("Select a Section Type.", "err"); return; }
  if (!form.colName) { setStatus("Enter a Column Name.", "err"); return; }
  if (form.items.length === 0) { setStatus("Add at least one line item.", "err"); return; }

  const isPR = form.secType === "Payment Request";
  setStatus(isPR ? "Saving PR + updating Vendor Tracking…" : "Saving to Workbook…", "busy");
  setBusy(true);
  try {
    let summary = "Done.";
    await Excel.run(async (context) => {
      summary = await runInputForm(context, form); // creates/fills the column in the Invoice Worksheet
      if (isPR) await runUpdatePR(context);        // mirrors PR columns into Vendor Tracking
    });
    setStatus(isPR ? `${summary} Vendor Tracking linked.` : summary, "ok");
    clearFormUI();
    poLoaded = false; // refresh PO contracts dropdown on next PO tab visit (new contract added)
  } catch (e) {
    console.error(e);
    setStatus("ERROR: " + errMsg(e), "err");
  } finally { setBusy(false); }
}

async function invoiceGen() {
  setStatus("Generating invoice…", "busy"); setBusy(true);
  try {
    let summary = "Invoice generated.";
    await Excel.run(async (context) => { summary = await runInvoiceGenerate(context); });
    setStatus(summary, "ok");
  } catch (e) {
    console.error(e);
    setStatus("ERROR: " + errMsg(e), "err");
  } finally { setBusy(false); }
}

async function getPR() {
  setStatus("Loading lines…", "busy"); setBusy(true);
  try {
    let loaded: { desc: string; amt: number | "" }[] = [];
    let latestPR = "";
    await Excel.run(async (context) => {
      loaded = await runGetPR(context);
      latestPR = await runGetLatestPRName(context);
    });

    // Show the "Latest Payment Request" hint above Column Name
    const hint = document.getElementById("latestPRHint") as HTMLElement;
    const nm = document.getElementById("latestPRName") as HTMLElement;
    if (hint && nm) {
      nm.textContent = latestPR || "(none yet)";
      hint.style.display = "";
    }

    (document.getElementById("secType") as HTMLSelectElement).value = "Payment Request";
    const body = document.getElementById("itemsBody") as HTMLTableSectionElement;
    body.innerHTML = "";
    for (const it of loaded) {
      addItemRow();
      const tr = body.rows[body.rows.length - 1];
      (tr.querySelector(".desc") as HTMLInputElement).value = it.desc;
      (tr.querySelector(".amt") as HTMLInputElement).value = it.amt === "" ? "" : String(it.amt);
    }
    for (let i = 0; i < 3; i++) addItemRow();
    prItemsAutoLoaded = true; // remember so we can clear if user switches type
    setStatus(`Loaded ${loaded.length} item(s). Edit amounts, set Column Name, then Save to Workbook.`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("ERROR: " + errMsg(e), "err");
  } finally { setBusy(false); }
}

// ───────────────────────── PO Management (panel) ─────────────────────────

function addPORow() {
  const body = document.getElementById("poBody") as HTMLTableSectionElement;
  const tr = body.insertRow();
  tr.innerHTML =
    `<td class="chk"><input class="po-mark" type="checkbox" /></td>` +
    `<td><input class="po-vendor" type="text" /></td>` +
    `<td><input class="po-num" type="text" /></td>` +
    `<td><input class="po-amt" type="number" step="any" /></td>` +
    `<td><input class="po-adj" type="number" step="any" /></td>`;
  wireAutoGrow(tr, body, addPORow);
}

function readPOForm(): { source: string; target: string; formData: (string | number)[][] } {
  const source = (document.getElementById("poSource") as HTMLSelectElement).value.trim();
  const target = (document.getElementById("poTarget") as HTMLSelectElement).value.trim();
  const formData: (string | number)[][] = [];
  qsa("#poBody tr").forEach((tr) => {
    const mark = (tr.querySelector(".po-mark") as HTMLInputElement).checked ? "Y" : "";
    const vendor = (tr.querySelector(".po-vendor") as HTMLInputElement).value.trim();
    const num = (tr.querySelector(".po-num") as HTMLInputElement).value.trim();
    const amt = (tr.querySelector(".po-amt") as HTMLInputElement).value.trim();
    const adj = (tr.querySelector(".po-adj") as HTMLInputElement).value.trim();
    formData.push([mark, vendor, num, amt === "" ? "" : Number(amt), adj === "" ? "" : Number(adj)]);
  });
  return { source, target, formData };
}

function fillContractSelect(id: string, contracts: string[], blankLabel: string) {
  const sel = document.getElementById(id) as HTMLSelectElement;
  const prev = sel.value;
  sel.innerHTML = "";
  const o0 = document.createElement("option");
  o0.value = ""; o0.textContent = blankLabel; sel.appendChild(o0);
  for (const c of contracts) {
    const o = document.createElement("option");
    o.value = c; o.textContent = c; sel.appendChild(o);
  }
  if (prev && contracts.includes(prev)) sel.value = prev;
}

async function loadEst() {
  setStatus("Loading contracts…", "busy"); setBusy(true);
  try {
    let contracts: string[] = [];
    await Excel.run(async (context) => { contracts = await runLoadEst(context); });
    fillContractSelect("poSource", contracts, "-- select contract --");
    fillContractSelect("poTarget", contracts, "-- none --");
    setStatus(`Loaded ${contracts.length} contract(s).`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("ERROR: " + errMsg(e), "err");
  } finally { setBusy(false); }
}

// In-page Yes/No modal for duplicate POs. `window.confirm` is silently
// ignored inside the Office Add-in webview, so we render our own.
function showDupeModal(dupes: string[]): Promise<boolean> {
  return new Promise((resolve) => {
    const modal = document.getElementById("dupeModal") as HTMLElement;
    const list = document.getElementById("dupeList") as HTMLElement;
    const yes = document.getElementById("dupeYes") as HTMLElement;
    const no = document.getElementById("dupeNo") as HTMLElement;
    list.innerHTML = "";
    for (const p of dupes) {
      const li = document.createElement("li");
      li.textContent = p;
      list.appendChild(li);
    }
    const cleanup = () => {
      modal.style.display = "none";
      yes.onclick = null;
      no.onclick = null;
    };
    yes.onclick = () => { cleanup(); resolve(true); };
    no.onclick = () => { cleanup(); resolve(false); };
    modal.style.display = "";
  });
}

async function updatePo() {
  const form = readPOForm();
  if (!form.source) { setStatus("Select a Source Contract (run Load Est first).", "err"); return; }

  // Detect action — only a pure ADD should be guarded for duplicate PO numbers.
  // UPDATE/MOVE intentionally re-touch existing rows (and "Load Po to Move" pre-fills the
  // form with existing PO numbers, so duplicate-checking that flow would always fire).
  let hasYMarks = false;
  for (const r of form.formData) {
    if (r[0] && String(r[0]).trim().toUpperCase() === "Y") { hasYMarks = true; break; }
  }
  const isMoveIntent = form.target !== "" && form.target !== form.source;
  const isAdd = !hasYMarks && !isMoveIntent;

  if (isAdd) {
    setStatus("Checking for duplicate PO numbers…", "busy"); setBusy(true);
    let dupes: string[] = [];
    try {
      await Excel.run(async (context) => { dupes = await runFindDuplicatePOs(context, form.formData); });
    } catch (e) {
      console.error(e);
      setStatus("ERROR: " + errMsg(e), "err");
      setBusy(false);
      return;
    }
    setBusy(false);
    hideStatus();

    if (dupes.length > 0) {
      const proceed = await showDupeModal(dupes);
      if (!proceed) {
        // User chose "No, Skip" — drop only the duplicate rows, keep unique ones.
        const dupeSet: { [k: string]: boolean } = {};
        for (const d of dupes) dupeSet[d.toUpperCase()] = true;
        const beforeCount = form.formData.length;
        form.formData = form.formData.filter((r) => {
          const pn = r[2] ? String(r[2]).trim().toUpperCase() : "";
          return !pn || !dupeSet[pn];
        });
        const skipped = beforeCount - form.formData.length;
        // Check if any non-empty rows remain
        let remaining = 0;
        for (const r of form.formData) {
          const vendor = r[1] ? String(r[1]).trim() : "";
          if (vendor !== "") remaining++;
        }
        if (remaining === 0) {
          setStatus(`All ${skipped} row(s) were duplicates — nothing added.`, "err");
          return;
        }
        setStatus(`Skipped ${skipped} duplicate row(s). Saving the rest…`, "busy");
      }
    }
  }

  setStatus("Saving POs…", "busy"); setBusy(true);
  try {
    let summary = "POs saved.";
    await Excel.run(async (context) => { summary = await runUpdatePo(context, form); });
    setStatus(summary, "ok");
    (document.getElementById("poBody") as HTMLTableSectionElement).innerHTML = "";
    for (let i = 0; i < 4; i++) addPORow();
    (document.getElementById("poTarget") as HTMLSelectElement).value = "";
  } catch (e) {
    console.error(e);
    setStatus("ERROR: " + errMsg(e), "err");
  } finally { setBusy(false); }
}

// ───────────────────────── helper ─────────────────────────

async function readValues(context: Excel.RequestContext, range: Excel.Range): Promise<(string | number | boolean)[][]> {
  range.load("values");
  await context.sync();
  return range.values as (string | number | boolean)[][];
}

// ───────────────────────── Main port (Run Input Form) ─────────────────────────

async function runInputForm(context: Excel.RequestContext, form: InputFormData) {
  const { dateVal, secType, colName, secHeader, vtTotal, vtCost, vtHours, vtPO, items } = form;

  const wsInv = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Invoice Worksheet");
  wsInv.load("name");
  await context.sync();
  if (wsInv.isNullObject) throw new Error("'LDP & LCP - Invoice Worksheet' not found.");

  let hdr2 = (await readValues(context, wsInv.getRange("A2").getResizedRange(0, 51)))[0];
  let hdr3 = (await readValues(context, wsInv.getRange("A3").getResizedRange(0, 51)))[0];

  let tdIdx = -1;
  for (let c = 0; c < hdr2.length; c++) {
    if (hdr2[c] && String(hdr2[c]).trim() === "Total To date") { tdIdx = c; break; }
  }
  if (tdIdx === -1) throw new Error("'Total To date' not found.");

  let targetIdx = -1;
  const isPR = secType === "Payment Request";

  if (isPR) {
    let paidS = -1;
    for (let c = tdIdx + 1; c < hdr2.length; c++) {
      if (hdr2[c] && String(hdr2[c]).trim() === "Paid to Date") { paidS = c; break; }
    }
    let insIdx = paidS !== -1 ? paidS : tdIdx + 1;
    if (paidS !== -1) {
      const prev = hdr3[paidS - 1] ? String(hdr3[paidS - 1]).trim().toUpperCase() : "";
      if (prev.includes("TBB")) insIdx = paidS - 1;
    }
    wsInv.getRange(`${CL(insIdx)}:${CL(insIdx)}`).insert(Excel.InsertShiftDirection.right);
    await context.sync();
    // After insert, PR#TBB shifts right by one column. Copy its column formatting
    // (borders, fills, number formats) onto the new PR column so it matches.
    const tbbColIdx = insIdx + 1;
    wsInv.getRange(`${CL(insIdx)}1:${CL(insIdx)}200`).copyFrom(
      wsInv.getRange(`${CL(tbbColIdx)}1:${CL(tbbColIdx)}200`),
      Excel.RangeCopyType.formats
    );
    wsInv.getRange(`${CL(insIdx)}2`).values = [["Payment Request"]];
    wsInv.getRange(`${CL(insIdx)}3`).values = [[colName]];
    if (dateVal !== "") {
      wsInv.getRange(`${CL(insIdx)}4`).values = [[dateVal]];
      wsInv.getRange(`${CL(insIdx)}4`).numberFormat = [["m/d/yyyy"]];
    }
    await context.sync();
    targetIdx = insIdx;
  } else {
    let found = -1;
    for (let c = 2; c < tdIdx; c++) {
      if (hdr3[c] && String(hdr3[c]).trim() === colName) { found = c; break; }
    }
    if (found !== -1) {
      targetIdx = found;
      // Existing column re-used — still write the date if the user supplied one
      // (the previous logic only wrote the date when a fresh column was inserted,
      // so updating an existing contract's date silently did nothing).
      if (dateVal !== "") {
        wsInv.getRange(`${CL(found)}4`).values = [[dateVal]];
        wsInv.getRange(`${CL(found)}4`).numberFormat = [["m/d/yyyy"]];
        await context.sync();
      }
    } else {
      let insIdx = tdIdx;
      const prev = hdr3[tdIdx - 1] ? String(hdr3[tdIdx - 1]).trim().toUpperCase() : "";
      const hadTBBNeighbour = prev.includes("TBB");
      if (hadTBBNeighbour) insIdx = tdIdx - 1;
      wsInv.getRange(`${CL(insIdx)}:${CL(insIdx)}`).insert(Excel.InsertShiftDirection.right);
      await context.sync();
      // After insert, CO#TBB (the previous neighbour at hdr3[tdIdx-1]) has shifted
      // right by one. Copy its column formatting onto the new column so the new
      // Change Order / LDP / LCP column matches the established style.
      if (hadTBBNeighbour) {
        const tbbColIdx = insIdx + 1;
        wsInv.getRange(`${CL(insIdx)}1:${CL(insIdx)}200`).copyFrom(
          wsInv.getRange(`${CL(tbbColIdx)}1:${CL(tbbColIdx)}200`),
          Excel.RangeCopyType.formats
        );
      }
      // Row 2 = section type. Previously hard-coded to "Change Order" which mis-labeled
      // brand-new LDP / LCP contracts.
      wsInv.getRange(`${CL(insIdx)}2`).values = [[secType]];
      wsInv.getRange(`${CL(insIdx)}3`).values = [[colName]];
      if (dateVal !== "") {
        wsInv.getRange(`${CL(insIdx)}4`).values = [[dateVal]];
        wsInv.getRange(`${CL(insIdx)}4`).numberFormat = [["m/d/yyyy"]];
      }
      await context.sync();
      targetIdx = insIdx;
      tdIdx++;
    }
  }

  hdr2 = (await readValues(context, wsInv.getRange("A2").getResizedRange(0, 51)))[0];
  tdIdx = -1;
  let paidIdx = -1, compIdx = -1, pctIdx = -1, balIdx = -1, firstPRIdx = -1, lastPRIdx = -1;
  for (let c = 0; c < hdr2.length; c++) {
    const h = hdr2[c] ? String(hdr2[c]).trim() : "";
    if (h === "Total To date" && tdIdx === -1) tdIdx = c;
    if (h === "Paid to Date" && paidIdx === -1) paidIdx = c;
    if (h === "Completed to Date" && compIdx === -1) compIdx = c;
    if (h.includes("% Completed") && pctIdx === -1) pctIdx = c;
    if (h === "Balance to Finish" && balIdx === -1) balIdx = c;
  }
  for (let c = tdIdx + 1; c < (paidIdx !== -1 ? paidIdx : hdr2.length); c++) {
    const h = hdr2[c] ? String(hdr2[c]).trim() : "";
    if (h === "Payment Request" || h === "Deposit - LDP/LCP") {
      if (firstPRIdx === -1) firstPRIdx = c;
      lastPRIdx = c;
    }
  }
  const tgtCol = CL(targetIdx);

  let colA = await readValues(context, wsInv.getRange("A1:A100"));
  let grandRow = -1, ldpSubRow = -1, lcpSubRow = -1, lcpHdrRow = -1;
  for (let r = 4; r < 100; r++) {
    const a = colA[r][0] ? String(colA[r][0]).trim() : "";
    if (a === "Grand - TOTALS") grandRow = r + 1;
    if (a === "SUB - TOTALS - LDP") ldpSubRow = r + 1;
    if (a === "SUB - TOTALS - LCP") lcpSubRow = r + 1;
    if (a === "LCP" && lcpHdrRow === -1) lcpHdrRow = r + 1;
  }
  if (grandRow === -1) throw new Error("Grand Totals not found.");

  if (isPR) {
    for (const item of items) {
      if (item.isHdr || item.amt === 0) continue;
      let matched = false;
      for (let r = 4; r < grandRow - 1; r++) {
        if (colA[r][0] && String(colA[r][0]).trim() === item.desc) {
          wsInv.getRange(`${tgtCol}${r + 1}`).values = [[item.amt]];
          wsInv.getRange(`${tgtCol}${r + 1}`).numberFormat = [[FMT_ACCT]];
          matched = true;
          break;
        }
      }
      if (!matched) {
        const insertRow = grandRow;
        wsInv.getRange(`${insertRow}:${insertRow}`).insert(Excel.InsertShiftDirection.down);
        wsInv.getRange(`A${insertRow}`).values = [[item.desc]];
        wsInv.getRange(`A${insertRow}`).format.font.bold = false;
        wsInv.getRange(`${tgtCol}${insertRow}`).values = [[item.amt]];
        wsInv.getRange(`${tgtCol}${insertRow}`).numberFormat = [[FMT_ACCT]];
        grandRow++;
        colA = await readValues(context, wsInv.getRange("A1:A100"));
      }
    }
    await updateAllFormulas(context, wsInv, grandRow, tdIdx, firstPRIdx, lastPRIdx, paidIdx, compIdx, pctIdx, balIdx, lcpHdrRow, ldpSubRow, lcpSubRow);
    wsInv.activate();
    await context.sync();
    return `Payment Request '${colName}' applied (${items.length} item(s)).`;
  }

  interface BlockRow { text: string; bold: boolean; amt: number; }
  const block: BlockRow[] = [];
  const needLCPHdr = secType === "LCP" && lcpHdrRow === -1;

  let insertAt: number;
  if (secType === "LDP") {
    const boundary = lcpHdrRow !== -1 ? lcpHdrRow : grandRow;
    insertAt = 6;
    for (let r = 5; r < boundary - 1; r++) {
      const a = colA[r][0] ? String(colA[r][0]).trim() : "";
      if (a !== "") insertAt = r + 2;
    }
    if (insertAt > boundary) insertAt = boundary;
  } else {
    insertAt = grandRow;
  }

  // One blank separator row between groups: add a blank only if the row above isn't already blank
  if (insertAt > 6) {
    const prevRowIdx = insertAt - 2;
    const prevVal = prevRowIdx >= 0 && prevRowIdx < colA.length
      ? (colA[prevRowIdx][0] ? String(colA[prevRowIdx][0]).trim() : "")
      : "";
    if (prevVal !== "") block.push({ text: "", bold: false, amt: 0 });
  }

  if (needLCPHdr) block.push({ text: "LCP", bold: true, amt: 0 });
  if (secHeader !== "") block.push({ text: secHeader, bold: true, amt: 0 });
  for (const item of items) block.push({ text: item.desc, bold: item.isHdr, amt: item.amt });

  // Always insert fresh rows at insertAt so existing data is NEVER overwritten
  // (pushes the rest — including the LCP section / totals — down).
  const rowsNeeded = block.length;
  if (rowsNeeded > 0) {
    wsInv.getRange(`${insertAt}:${insertAt + rowsNeeded - 1}`).insert(Excel.InsertShiftDirection.down);
    grandRow += rowsNeeded;
    if (ldpSubRow !== -1) ldpSubRow += rowsNeeded;
    if (lcpSubRow !== -1) lcpSubRow += rowsNeeded;
    if (lcpHdrRow !== -1 && lcpHdrRow >= insertAt) lcpHdrRow += rowsNeeded;
    await context.sync();
  }

  for (let i = 0; i < block.length; i++) {
    const row = insertAt + i;
    const b = block[i];
    if (b.text === "") continue;
    wsInv.getRange(`A${row}`).values = [[b.text]];
    wsInv.getRange(`A${row}`).format.font.bold = b.bold;
    if (b.amt !== 0) {
      wsInv.getRange(`${tgtCol}${row}`).values = [[b.amt]];
      wsInv.getRange(`${tgtCol}${row}`).numberFormat = [[FMT_ACCT]];
    }
    if (b.text === "LCP" && b.bold && needLCPHdr) lcpHdrRow = row;
  }
  await context.sync();

  // Save line-item descriptions into PR#TBB (LDP / LCP only) — descriptions only, appended
  if (secType === "LDP" || secType === "LCP") {
    await addDescriptionsToPRTBB(context, items);
  }

  const wsVT = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Vendor Tracking");
  wsVT.load("name");
  await context.sync();
  if (!wsVT.isNullObject) {
    await addVendorTrackingRow(context, wsVT, secType, secHeader, colName, vtTotal, vtCost, vtHours, vtPO);
    await updateVTFormulas(context, wsVT);
    await repairAllVendorTrackingFormulas(context, wsVT);
    await loadPOContracts(context);
  }

  await updateAllFormulas(context, wsInv, grandRow, tdIdx, firstPRIdx, lastPRIdx, paidIdx, compIdx, pctIdx, balIdx, lcpHdrRow, ldpSubRow, lcpSubRow);
  wsInv.activate();
  await context.sync();
  return `${secType} added → column ${tgtCol} (${colName}), ${items.length} item(s).`;
}

// ───────────────────────── PR#TBB descriptions ─────────────────────────

async function addDescriptionsToPRTBB(context: Excel.RequestContext, items: Item[]) {
  const wsPR = context.workbook.worksheets.getItemOrNullObject("PR#TBB");
  wsPR.load("name");
  await context.sync();
  if (wsPR.isNullObject) return;

  let colA = await readValues(context, wsPR.getRange("A1:A60"));
  let subTotalRow = -1;
  for (let r = 0; r < colA.length; r++) {
    const a = colA[r][0] ? String(colA[r][0]).trim().toUpperCase().replace(/\s+/g, "") : "";
    if (a.startsWith("SUB-TOTAL")) { subTotalRow = r + 1; break; }
  }
  if (subTotalRow === -1) return;

  const descStart = 7;
  for (const item of items) {
    const desc = item.desc ? String(item.desc).trim() : "";
    if (desc === "") continue;

    colA = await readValues(context, wsPR.getRange("A1:A60"));
    let targetRow = -1;
    for (let r = descStart - 1; r < subTotalRow - 1; r++) {
      const v = colA[r] ? colA[r][0] : "";
      if (v === "" || v === null) { targetRow = r + 1; break; }
    }

    if (targetRow === -1) {
      targetRow = subTotalRow;
      wsPR.getRange(`${targetRow}:${targetRow}`).insert(Excel.InsertShiftDirection.down);
      await context.sync();
      // Inherit the template row's formatting (borders, fills, merge) so the
      // newly inserted row looks identical to the existing description rows.
      const templateRow = targetRow - 1;
      if (templateRow >= descStart) {
        wsPR.getRange(`A${targetRow}:M${targetRow}`).copyFrom(
          wsPR.getRange(`A${templateRow}:M${templateRow}`),
          Excel.RangeCopyType.formats
        );
        await context.sync();
      }
      // Set Aptos 9pt on the WHOLE A:D range BEFORE merging so the right-most
      // cell's font doesn't win after merge.
      wsPR.getRange(`A${targetRow}:D${targetRow}`).format.font.name = "Aptos";
      wsPR.getRange(`A${targetRow}:D${targetRow}`).format.font.size = 9;
      wsPR.getRange(`A${targetRow}:D${targetRow}`).merge(false);
      subTotalRow++;
    }
    wsPR.getRange(`A${targetRow}`).values = [[desc]];
    // Re-assert Aptos 9pt + normal weight on the merged anchor.
    wsPR.getRange(`A${targetRow}`).format.font.name = "Aptos";
    wsPR.getRange(`A${targetRow}`).format.font.size = 9;
    // Column A descriptions always render in normal weight (not bold), even for
    // sub-header items.
    wsPR.getRange(`A${targetRow}`).format.font.bold = false;
    await context.sync();
  }
}

// ───────────────────────── Vendor Tracking ─────────────────────────

async function addVendorTrackingRow(
  context: Excel.RequestContext, ws: Excel.Worksheet,
  secType: string, secHeader: string, colName: string,
  vtTotal: number | "", vtCost: number | "", vtHours: number | "", vtPO: string
) {
  if (secType === "Payment Request") return;
  const contractName = secHeader !== "" ? secHeader : colName;

  const vtAE = await readValues(context, ws.getRange("A5:E80")); // A..E for rows 5..80
  let lastDataRow = 6, clientTotalRow = -1, lcpAnalysisRow = -1, ldpMarkerRow = -1, lcpMarkerRow = -1;

  for (let r = 0; r < vtAE.length; r++) {
    const a = vtAE[r][0] ? String(vtAE[r][0]).trim() : "";
    const rowNum = r + 5;
    if (a === "LDP" && ldpMarkerRow === -1) ldpMarkerRow = rowNum;
    if (a === "LCP" && lcpMarkerRow === -1) lcpMarkerRow = rowNum;
    if (a.includes("LCP Analysis")) lcpAnalysisRow = rowNum;
    if (a.includes("Client total") || a.includes("Client Total")) { clientTotalRow = rowNum; break; }
    if (a !== "" && a !== "Contract" && !a.includes("Project Management") && !a.includes("Sub-Contractor")
      && !a.includes("Percentage") && !a.includes("Gross Margin") && !a.includes("LCP Analysis")) {
      lastDataRow = rowNum;
    }
  }

  const rowEmpty = (rowNum: number): boolean => {
    const r = rowNum - 5;
    if (r < 0 || r >= vtAE.length) return true;
    return vtAE[r].slice(0, 5).every((v) => v === "" || v === null);
  };

  let vtRow = -1;

  if (secType === "LDP" || secType === "LCP" || secType === "Change Order") {
    let sectionStart: number, sectionEnd: number;
    if (secType === "LDP") {
      sectionStart = ldpMarkerRow !== -1 ? ldpMarkerRow + 1 : 7;
      sectionEnd = lcpMarkerRow !== -1 ? lcpMarkerRow
        : (lcpAnalysisRow !== -1 ? lcpAnalysisRow : (clientTotalRow !== -1 ? clientTotalRow : 81));
    } else {
      sectionStart = lcpMarkerRow !== -1 ? lcpMarkerRow + 1 : 7;
      sectionEnd = lcpAnalysisRow !== -1 ? lcpAnalysisRow : (clientTotalRow !== -1 ? clientTotalRow : 81);
    }

    // Scan for the LAST non-empty row in this section (not the first empty one).
    // Previously the code looked for the first gap and inserted there, which dropped
    // new contracts into the middle of the section when separators existed between
    // earlier entries. New contracts should always land at the BOTTOM of the section.
    let lastUsed = -1;
    for (let row = sectionStart; row < sectionEnd; row++) {
      if (!rowEmpty(row)) lastUsed = row;
    }

    if (lastUsed === -1) {
      // Section is empty → first contract goes right after the section marker
      vtRow = sectionStart;
    } else {
      // Always place the new contract just before sectionEnd (the LCP marker / LCP
      // Analysis / Client total row). Keep a blank separator above it.
      const separatorNeeded = !rowEmpty(lastUsed + 1);
      if (separatorNeeded) {
        // No existing gap → insert 2 rows; first becomes the separator, second the contract
        ws.getRange(`${sectionEnd}:${sectionEnd + 1}`).insert(Excel.InsertShiftDirection.down);
        await context.sync();
        vtRow = sectionEnd + 1;
      } else {
        // There's already an empty row right after lastUsed acting as a separator;
        // we just need one new row right before sectionEnd for the contract itself.
        ws.getRange(`${sectionEnd}:${sectionEnd}`).insert(Excel.InsertShiftDirection.down);
        await context.sync();
        vtRow = sectionEnd;
      }
    }
  } else {
    const stopRow = lcpAnalysisRow !== -1 ? lcpAnalysisRow : (clientTotalRow !== -1 ? clientTotalRow : 81);
    for (let row = 7; row < stopRow; row++) {
      if (rowEmpty(row)) { vtRow = row; break; }
    }
    if (vtRow === -1) {
      if (lcpAnalysisRow !== -1) {
        vtRow = lcpAnalysisRow;
        ws.getRange(`${lcpAnalysisRow}:${lcpAnalysisRow + 1}`).insert(Excel.InsertShiftDirection.down);
      } else if (clientTotalRow !== -1) {
        vtRow = clientTotalRow;
        ws.getRange(`${clientTotalRow}:${clientTotalRow + 1}`).insert(Excel.InsertShiftDirection.down);
      } else {
        vtRow = lastDataRow + 2;
      }
      await context.sync();
    }
  }

  ws.getRange(`A${vtRow}`).values = [[contractName]];
  ws.getRange(`A${vtRow}:E${vtRow}`).format.font.bold = true;
  if (vtTotal !== "") { ws.getRange(`B${vtRow}`).values = [[vtTotal]]; ws.getRange(`B${vtRow}`).numberFormat = [[FMT_ACCT]]; }
  if (vtCost !== "") { ws.getRange(`C${vtRow}`).values = [[vtCost]]; ws.getRange(`C${vtRow}`).numberFormat = [[FMT_ACCT]]; }
  if (vtHours !== "") { ws.getRange(`D${vtRow}`).values = [[vtHours]]; ws.getRange(`D${vtRow}`).numberFormat = [["0.00"]]; }
  if (vtPO !== "") { ws.getRange(`E${vtRow}`).values = [[vtPO]]; ws.getRange(`E${vtRow}`).numberFormat = [["@"]]; }
  await context.sync();

  const next = await readValues(context, ws.getRange(`A${vtRow + 1}:E${vtRow + 1}`));
  const isNextRowEmpty = next[0].every((v) => v === "" || v === null);
  if (!isNextRowEmpty) {
    ws.getRange(`${vtRow + 1}:${vtRow + 1}`).insert(Excel.InsertShiftDirection.down);
    await context.sync();
  }
}

// ───── Vendor Tracking client-total helpers ─────
// The workbook can host three independent analysis blocks, each with its own
// "Client total Contract/Cost" row:
//   • LCP Analysis  → totals span only the LCP rows
//   • LDP Analysis  → totals span only the LDP rows
//   • Consolidated  → totals = LCP analysis row + LDP analysis row

// Scan vtA (rows 5..N) for markers and the three analysis headers, returning the
// row numbers and the immediately-following Client total row for each section.
interface VtSections {
  ldpMarker: number; lcpMarker: number;
  lcpAnalysisRow: number; ldpAnalysisRow: number; consAnalysisRow: number;
  lcpClientRow: number; ldpClientRow: number; consClientRow: number;
}

function scanVtSections(vtA: (string | number | boolean)[][]): VtSections {
  const s: VtSections = {
    ldpMarker: -1, lcpMarker: -1,
    lcpAnalysisRow: -1, ldpAnalysisRow: -1, consAnalysisRow: -1,
    lcpClientRow: -1, ldpClientRow: -1, consClientRow: -1,
  };
  for (let r = 0; r < vtA.length; r++) {
    const a = vtA[r][0] ? String(vtA[r][0]).trim() : "";
    const rowNum = r + 5;
    if (a === "LDP" && s.ldpMarker === -1) s.ldpMarker = rowNum;
    if (a === "LCP" && s.lcpMarker === -1) s.lcpMarker = rowNum;
    if (a.indexOf("LCP Analysis") !== -1 && s.lcpAnalysisRow === -1) s.lcpAnalysisRow = rowNum;
    if (a.indexOf("LDP Analysis") !== -1 && s.ldpAnalysisRow === -1) s.ldpAnalysisRow = rowNum;
    if (a.indexOf("Consolidated") !== -1 && s.consAnalysisRow === -1) s.consAnalysisRow = rowNum;
  }
  const findClientAfter = (anchor: number): number => {
    if (anchor === -1) return -1;
    for (let r = anchor - 5 + 1; r < vtA.length; r++) {
      const a = vtA[r][0] ? String(vtA[r][0]).trim() : "";
      if (a.indexOf("Client total") !== -1 || a.indexOf("Client Total") !== -1) return r + 5;
      // Stop scanning if we hit the next analysis header (no client total in this block)
      if (a.indexOf("Analysis") !== -1 && r + 5 !== anchor) return -1;
    }
    return -1;
  };
  s.lcpClientRow = findClientAfter(s.lcpAnalysisRow);
  s.ldpClientRow = findClientAfter(s.ldpAnalysisRow);
  s.consClientRow = findClientAfter(s.consAnalysisRow);
  return s;
}

// Write the per-section formulas onto each Client total row.
function writeClientTotalFormulas(ws: Excel.Worksheet, s: VtSections): void {
  const setRow = (clientRow: number, build: (col: string) => string) => {
    if (clientRow === -1) return;
    const fB = build("B"), fC = build("C"), fD = build("D");
    if (fB) { ws.getRange(`B${clientRow}`).formulas = [[fB]]; ws.getRange(`B${clientRow}`).numberFormat = [[FMT_ACCT]]; }
    if (fC) { ws.getRange(`C${clientRow}`).formulas = [[fC]]; ws.getRange(`C${clientRow}`).numberFormat = [[FMT_ACCT]]; }
    if (fD) { ws.getRange(`D${clientRow}`).formulas = [[fD]]; ws.getRange(`D${clientRow}`).numberFormat = [["0.00"]]; }
  };

  // LCP Analysis: sum rows just after the LCP marker, just before the LCP Analysis header
  if (s.lcpMarker !== -1 && s.lcpAnalysisRow !== -1 && s.lcpAnalysisRow > s.lcpMarker + 1) {
    const start = s.lcpMarker + 1, end = s.lcpAnalysisRow - 1;
    setRow(s.lcpClientRow, (col) => `=SUM(${col}${start}:${col}${end})`);
  }
  // LDP Analysis: sum rows just after the LDP marker, just before the LCP marker
  if (s.ldpMarker !== -1 && s.lcpMarker !== -1 && s.lcpMarker > s.ldpMarker + 1) {
    const start = s.ldpMarker + 1, end = s.lcpMarker - 1;
    setRow(s.ldpAnalysisRow !== -1 ? s.ldpClientRow : -1, (col) => `=SUM(${col}${start}:${col}${end})`);
  }
  // Consolidated: =LCPClientRow + LDPClientRow (auto-updates if either total changes)
  if (s.consClientRow !== -1) {
    setRow(s.consClientRow, (col) => {
      const refs: string[] = [];
      if (s.lcpClientRow !== -1) refs.push(`${col}${s.lcpClientRow}`);
      if (s.ldpClientRow !== -1) refs.push(`${col}${s.ldpClientRow}`);
      return refs.length === 0 ? "" : "=" + refs.join("+");
    });
  }
}

async function updateVTFormulas(context: Excel.RequestContext, ws: Excel.Worksheet) {
  const vtA = await readValues(context, ws.getRange("A5:A80"));
  const s = scanVtSections(vtA);
  writeClientTotalFormulas(ws, s);
  await context.sync();
}

async function repairAllVendorTrackingFormulas(context: Excel.RequestContext, wsVT: Excel.Worksheet) {
  const headers = (await readValues(context, wsVT.getRange("A5:AD5")))[0];
  let poTotalIdx = -1, ldpTotalIdx = -1, lcpTotalIdx = -1;
  for (let c = 0; c < headers.length; c++) {
    const h = headers[c] ? String(headers[c]).trim() : "";
    if (h === "PO Total") poTotalIdx = c;
    if (h === "LDP Total") ldpTotalIdx = c;
    if (h === "LCP Total") lcpTotalIdx = c;
  }
  // LCP total column may simply be labeled "Total" (sits under "LCP Project", after LDP Total)
  if (lcpTotalIdx === -1) {
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c] ? String(headers[c]).trim() : "";
      if (h === "Total" && c > ldpTotalIdx) { lcpTotalIdx = c; break; }
    }
  }
  if (poTotalIdx === -1) return;

  const poTotalCol = CL(poTotalIdx);
  const ldpTotalCol = ldpTotalIdx !== -1 ? CL(ldpTotalIdx) : "";
  const lcpTotalCol = lcpTotalIdx !== -1 ? CL(lcpTotalIdx) : "";

  const rows = await readValues(context, wsVT.getRange("A6:E80")); // A..E
  let section = "LDP"; // driven by bold "LDP"/"LCP" marker rows in column A
  for (let i = 0; i < rows.length; i++) {
    const row = i + 6;
    const name = rows[i][0] ? String(rows[i][0]).trim() : "";

    if (name === "LDP") { section = "LDP"; continue; }
    if (name === "LCP") { section = "LCP"; continue; }
    if (name.includes("Analysis") || name.includes("Client total") || name.includes("Client Total")) break;
    if (name === "" || name.includes("Project Management") || name.includes("Sub-Contractor")
      || name.includes("Percentage") || name.includes("Gross Margin")) continue;

    const b = rows[i][1];
    const e = rows[i][4];
    const isMainContract = b !== "" && b !== null && Number(b) > 0;
    const isVendorRow = e !== "" && e !== null && !isMainContract;
    if (!isVendorRow) continue;

    wsVT.getRange(`${poTotalCol}${row}`).formulas = [[`=SUM(F${row}:G${row})`]];
    wsVT.getRange(`${poTotalCol}${row}`).numberFormat = [[FMT_ACCT]];

    if (section === "LCP" && lcpTotalCol && lcpTotalIdx > ldpTotalIdx) {
      const firstLcpPrCol = CL(ldpTotalIdx + 1);
      const lastLcpPrCol = CL(lcpTotalIdx - 1);
      wsVT.getRange(`${lcpTotalCol}${row}`).formulas = [[`=${poTotalCol}${row}+SUM(${firstLcpPrCol}${row}:${lastLcpPrCol}${row})`]];
      wsVT.getRange(`${lcpTotalCol}${row}`).numberFormat = [[FMT_ACCT]];
      if (ldpTotalCol) wsVT.getRange(`${ldpTotalCol}${row}`).values = [[""]];
    } else if (ldpTotalCol && ldpTotalIdx > poTotalIdx) {
      const firstLdpPrCol = CL(poTotalIdx + 1);
      const lastLdpPrCol = CL(ldpTotalIdx - 1);
      wsVT.getRange(`${ldpTotalCol}${row}`).formulas = [[`=${poTotalCol}${row}+SUM(${firstLdpPrCol}${row}:${lastLdpPrCol}${row})`]];
      wsVT.getRange(`${ldpTotalCol}${row}`).numberFormat = [[FMT_ACCT]];
      if (lcpTotalCol) wsVT.getRange(`${lcpTotalCol}${row}`).values = [[""]];
    }
  }
  await context.sync();
}

async function loadPOContracts(context: Excel.RequestContext) {
  const wsVT = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Vendor Tracking");
  const wsPO = context.workbook.worksheets.getItemOrNullObject("PO Input Form");
  wsVT.load("name"); wsPO.load("name");
  await context.sync();
  if (wsVT.isNullObject || wsPO.isNullObject) return;

  const vt = await readValues(context, wsVT.getRange("A7:E50")); // A,B,...,E
  const contracts: string[] = [];
  for (let r = 0; r < vt.length; r++) {
    const a = vt[r][0] ? String(vt[r][0]).trim() : "";
    const b = vt[r][1];
    const e = vt[r][4] ? String(vt[r][4]).trim() : "";
    if (a === "" || a.includes("Client total") || a.includes("Client Total") || a.includes("Project Management")
      || a.includes("Sub-Contractor") || a.includes("Percentage") || a.includes("Gross Margin") || a === "Without Estimate") continue;
    if (e !== "" && b !== "" && b !== null && Number(b) > 0) contracts.push(e);
  }
  contracts.push("Without Estimate");
  const list = contracts.join(",");

  for (const cell of ["B4", "C4"]) {
    const dv = wsPO.getRange(cell).dataValidation;
    dv.clear();
    dv.rule = { list: { inCellDropDown: true, source: list } };
  }
  await context.sync();
}

// ───────────────────────── Invoice Worksheet formulas ─────────────────────────

async function updateAllFormulas(
  context: Excel.RequestContext, ws: Excel.Worksheet, grandRow: number,
  tdIdx: number, firstPRIdx: number, lastPRIdx: number,
  paidIdx: number, compIdx: number, pctIdx: number, balIdx: number,
  lcpHdrRow: number, ldpSubRow: number, lcpSubRow: number
) {
  const tdCol = CL(tdIdx), lastCO = CL(tdIdx - 1);
  const prS = firstPRIdx !== -1 ? CL(firstPRIdx) : "";
  const prE = lastPRIdx !== -1 ? CL(lastPRIdx) : "";
  const hCol = paidIdx !== -1 ? CL(paidIdx) : "";
  const cCol = compIdx !== -1 ? CL(compIdx) : "";
  const pCol = pctIdx !== -1 ? CL(pctIdx) : "";
  const bCol = balIdx !== -1 ? CL(balIdx) : "";

  const colA = await readValues(context, ws.getRange("A1:A100"));

  for (let r = 5; r <= grandRow - 1; r++) {
    const a = colA[r - 1][0] ? String(colA[r - 1][0]).trim() : "";
    if (a === "" || a === "Grand - TOTALS" || a.startsWith("SUB - TOTALS")) continue;

    ws.getRange(`${tdCol}${r}`).formulas = [[`=IF(SUMPRODUCT(--(B${r}:${lastCO}${r}<>""))=0,"",SUM(B${r}:${lastCO}${r}))`]];
    ws.getRange(`${tdCol}${r}`).numberFormat = [[FMT_USD]];

    if (hCol && prS) {
      // Paid to Date: SUM of all PR amounts minus PR#TBB (last PR col)
      ws.getRange(`${hCol}${r}`).formulas = [[`=IF(${tdCol}${r}="","",SUM(${prS}${r}:${prE}${r})-${prE}${r})`]];
      ws.getRange(`${hCol}${r}`).numberFormat = [[FMT_ACCT]];
    }
    if (cCol && prS) {
      // Completed to Date: SUM of all PR amounts including PR#TBB
      ws.getRange(`${cCol}${r}`).formulas = [[`=IF(${tdCol}${r}="","",SUM(${prS}${r}:${prE}${r}))`]];
      ws.getRange(`${cCol}${r}`).numberFormat = [[FMT_ACCT]];
    }
    if (pCol) {
      ws.getRange(`${pCol}${r}`).formulas = [[`=IFERROR(SUM(${cCol}${r}/${tdCol}${r}),"")`]];
      ws.getRange(`${pCol}${r}`).numberFormat = [["0%"]];
    }
    if (bCol) {
      // Balance to Finish: TotalToDate - Completed (only guards against blank TD)
      ws.getRange(`${bCol}${r}`).formulas = [[`=IF(${tdCol}${r}="","",(${tdCol}${r}-${cCol}${r}))`]];
      ws.getRange(`${bCol}${r}`).numberFormat = [[FMT_USD]];
    }
  }

  const dEnd = grandRow - 1;
  let ldpEnd = dEnd;
  if (lcpHdrRow !== -1) ldpEnd = lcpHdrRow - 1;

  const hdr2 = (await readValues(context, ws.getRange("A2").getResizedRange(0, 51)))[0];
  setTotalsRow(ws, hdr2, grandRow, 5, dEnd, tdIdx, compIdx, balIdx);
  if (ldpSubRow !== -1) setTotalsRow(ws, hdr2, ldpSubRow, 5, ldpEnd, tdIdx, compIdx, balIdx);
  if (lcpSubRow !== -1 && ldpSubRow !== -1) setLCPRow(ws, hdr2, lcpSubRow, grandRow, ldpSubRow, tdIdx, compIdx, balIdx);
  await context.sync();
}

function setTotalsRow(ws: Excel.Worksheet, hdr2: (string | number | boolean)[], row: number, s: number, e: number,
  tdIdx: number, compIdx: number, balIdx: number) {
  const end = balIdx !== -1 ? balIdx : tdIdx;
  for (let c = 1; c <= end; c++) {
    const col = CL(c);
    const hs = hdr2[c] ? String(hdr2[c]).trim() : "";
    if (hs === "Total To date") {
      ws.getRange(`${col}${row}`).formulas = [[`=SUM(B${row}:${CL(tdIdx - 1)}${row})`]];
      ws.getRange(`${col}${row}`).numberFormat = [[FMT_USD]];
    } else if (hs.includes("% Completed")) {
      ws.getRange(`${col}${row}`).formulas = [[`=IFERROR(SUM(${CL(compIdx)}${row}/${CL(tdIdx)}${row}),0)`]];
      ws.getRange(`${col}${row}`).numberFormat = [["0%"]];
    } else if (hs === "Balance to Finish") {
      ws.getRange(`${col}${row}`).formulas = [[`=IF(OR(${CL(tdIdx)}${row}="",${CL(compIdx)}${row}=""),"",${CL(tdIdx)}${row}-${CL(compIdx)}${row})`]];
      ws.getRange(`${col}${row}`).numberFormat = [[FMT_USD]];
    } else {
      ws.getRange(`${col}${row}`).formulas = [[`=SUM(${col}${s}:${col}${e})`]];
    }
  }
}

function setLCPRow(ws: Excel.Worksheet, hdr2: (string | number | boolean)[], row: number, gr: number, lr: number,
  tdIdx: number, compIdx: number, balIdx: number) {
  const end = balIdx !== -1 ? balIdx : tdIdx;
  for (let c = 1; c <= end; c++) {
    const col = CL(c);
    const hs = hdr2[c] ? String(hdr2[c]).trim() : "";
    if (hs === "Total To date") {
      ws.getRange(`${col}${row}`).formulas = [[`=SUM(B${row}:${CL(tdIdx - 1)}${row})`]];
      ws.getRange(`${col}${row}`).numberFormat = [[FMT_USD]];
    } else if (hs.includes("% Completed")) {
      ws.getRange(`${col}${row}`).formulas = [[`=IFERROR(SUM(${CL(compIdx)}${row}/${CL(tdIdx)}${row}),0)`]];
      ws.getRange(`${col}${row}`).numberFormat = [["0%"]];
    } else if (hs === "Balance to Finish") {
      ws.getRange(`${col}${row}`).formulas = [[`=IF(OR(${CL(tdIdx)}${row}="",${CL(compIdx)}${row}=""),"",${CL(tdIdx)}${row}-${CL(compIdx)}${row})`]];
      ws.getRange(`${col}${row}`).numberFormat = [[FMT_USD]];
    } else {
      ws.getRange(`${col}${row}`).formulas = [[`=${col}${gr}-${col}${lr}`]];
    }
  }
}

// ───────────────────────── Run Update pr (Vendor Tracking) ─────────────────────────

async function runUpdatePR(context: Excel.RequestContext) {
  const wsVT = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Vendor Tracking");
  const wsInv = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Invoice Worksheet");
  wsVT.load("name"); wsInv.load("name");
  await context.sync();
  if (wsVT.isNullObject || wsInv.isNullObject) throw new Error("Sheet not found.");

  const invH2 = (await readValues(context, wsInv.getRange("A2").getResizedRange(0, 30)))[0];
  const invH3 = (await readValues(context, wsInv.getRange("A3").getResizedRange(0, 30)))[0];
  let tdIdx = -1, paidIdx = -1;
  for (let c = 0; c < invH2.length; c++) {
    const h = invH2[c] ? String(invH2[c]).trim() : "";
    if (h === "Total To date" && tdIdx === -1) tdIdx = c;
    if (h === "Paid to Date" && paidIdx === -1) paidIdx = c;
  }
  if (tdIdx === -1) throw new Error("Total To date not found.");

  const invA = await readValues(context, wsInv.getRange("A1:A100"));
  let grandRow = -1, lcpHdrRow = -1;
  for (let r = 4; r < 100; r++) {
    const a = invA[r][0] ? String(invA[r][0]).trim() : "";
    if (a === "LCP" && lcpHdrRow === -1) lcpHdrRow = r + 1;
    if (a === "Grand - TOTALS") { grandRow = r + 1; break; }
  }
  if (grandRow === -1) throw new Error("Grand Totals not found.");

  const prEnd = paidIdx !== -1 ? paidIdx : 30;
  const invPRs: { name: string; col: number; isLCP: boolean }[] = [];
  if (prEnd > tdIdx + 1) {
    const prBlock = await readValues(context, wsInv.getRange(`${CL(tdIdx + 1)}5:${CL(prEnd - 1)}${grandRow - 1}`));
    for (let c = tdIdx + 1; c < prEnd; c++) {
      const name = invH3[c] ? String(invH3[c]).trim() : "";
      if (name === "") continue;
      let hasLDP = false, hasLCP = false;
      const colOff = c - (tdIdx + 1);
      for (let r = 0; r < prBlock.length; r++) {
        const v = prBlock[r][colOff];
        if (v !== null && v !== "" && v !== 0) {
          if (lcpHdrRow !== -1 && (r + 5) >= lcpHdrRow) hasLCP = true;
          else hasLDP = true;
        }
      }
      // Mirror EVERY PR that has any data (LDP, LCP, or mixed) into Vendor Tracking.
      // The earlier "LCP-only" gating was too strict and caused PRs with any LDP line
      // to silently skip the VT mirror.
      invPRs.push({ name, col: c, isLCP: hasLCP || hasLDP });
    }
  }

  const vtH = (await readValues(context, wsVT.getRange("A5").getResizedRange(0, 30)))[0];
  let vtNotes = -1;
  for (let c = 0; c < vtH.length; c++) { if (vtH[c] && String(vtH[c]).trim() === "Notes") vtNotes = c; }
  const validEnd = vtNotes !== -1 ? vtNotes : 30;
  const vtExist = new Set<string>();
  for (let c = 8; c < validEnd; c++) {
    const h = vtH[c] ? String(vtH[c]).trim() : "";
    if (h !== "" && h !== "LDP Total" && h !== "LCP Total") vtExist.add(h);
  }
  // Mirror every PR that has any data into Vendor Tracking — see classification above.
  // .reverse() = create right-to-left.
  const missing = invPRs.filter((p) => p.isLCP && !vtExist.has(p.name)).reverse();

  const vtA = await readValues(context, wsVT.getRange("A5:A60"));
  let clientRow = -1, subContRow = -1;
  for (let r = 0; r < vtA.length; r++) {
    const a = vtA[r][0] ? String(vtA[r][0]).trim() : "";
    if (a.includes("Client total") || a.includes("Client Total")) clientRow = r + 5;
    if (a.includes("Sub-Contractor Total")) subContRow = r + 5;
  }

  for (const pr of missing) {
    const cur = (await readValues(context, wsVT.getRange("A5").getResizedRange(0, 30)))[0];
    let curLDPTot = -1, curTBB = -1, curLCPTot = -1;
    for (let c = 0; c < cur.length; c++) {
      const h = cur[c] ? String(cur[c]).trim() : "";
      if (h === "LDP Total") curLDPTot = c;
      if (h.toUpperCase().includes("TBB")) curTBB = c;
      if (h === "LCP Total") curLCPTot = c;
    }
    // Payment Requests always belong in the LCP / PR#TBB area — insert before PR#TBB (never before LDP Total)
    let insertAt = -1;
    if (curTBB !== -1) insertAt = curTBB;
    else if (curLCPTot !== -1) insertAt = curLCPTot;
    else if (curLDPTot !== -1) insertAt = curLDPTot;
    if (insertAt === -1) continue;
    const col = CL(insertAt);
    wsVT.getRange(`${col}:${col}`).insert(Excel.InsertShiftDirection.right);
    await context.sync();
    // Drop the bad inherited stuff (formulas / values / yellow fill / conditional formats)
    // but KEEP borders / font / number-format so the column still looks bordered.
    const newCol = wsVT.getRange(`${col}1:${col}100`);
    newCol.clear(Excel.ClearApplyTo.contents);
    newCol.format.fill.clear();
    newCol.conditionalFormats.clearAll();
    wsVT.getRange(`${col}5`).values = [[pr.name]];
    wsVT.getRange(`${col}6`).formulas = [[`='LDP & LCP - Invoice Worksheet'!${CL(pr.col)}${grandRow}`]];
    wsVT.getRange(`${col}6`).numberFormat = [[FMT_ACCT]];
    await context.sync();

    // Mirror PR#TBB's bottom-row formulas (column totals + Project Indicator rows like
    // Total Cost LCP / Cost Percentage / Total Left to Pay Vendors) into the new PR column.
    // Excel's formulas-only paste-special adjusts relative column references automatically,
    // so SUM/ratio formulas retarget to the new column.
    const tbbColIdx = insertAt + 1; // PR#TBB shifted right by 1 after the insert above
    let indicRow = -1;
    const scanVT = await readValues(context, wsVT.getRange("A1:Z60"));
    for (let r = 0; r < scanVT.length && indicRow === -1; r++) {
      for (let cc = 0; cc < scanVT[r].length; cc++) {
        const v = scanVT[r][cc] ? String(scanVT[r][cc]).trim().toLowerCase() : "";
        if (v.indexOf("project indicators") !== -1) { indicRow = r + 1; break; }
      }
    }
    if (indicRow === -1) indicRow = 23; // safe fallback
    const lastRow = indicRow + 25;
    wsVT.getRange(`${col}${indicRow}:${col}${lastRow}`).copyFrom(
      wsVT.getRange(`${CL(tbbColIdx)}${indicRow}:${CL(tbbColIdx)}${lastRow}`),
      Excel.RangeCopyType.formulas
    );
    await context.sync();
  }

  const fin = (await readValues(context, wsVT.getRange("A5").getResizedRange(0, 30)))[0];

  for (const pr of invPRs) {
    if (!pr.isLCP) continue; // LDP PRs are not mirrored into Vendor Tracking
    for (let c = 8; c < fin.length; c++) {
      const h = fin[c] ? String(fin[c]).trim() : "";
      if (h === pr.name) {
        wsVT.getRange(`${CL(c)}6`).formulas = [[`='LDP & LCP - Invoice Worksheet'!${CL(pr.col)}${grandRow}`]];
        wsVT.getRange(`${CL(c)}6`).numberFormat = [[FMT_ACCT]];
        break;
      }
    }
  }

  let nLDPTot = -1, nLCPTot = -1, nTBB = -1;
  const ldp: number[] = [], lcp: number[] = [];
  for (let c = 0; c < fin.length; c++) {
    const h = fin[c] ? String(fin[c]).trim() : "";
    if (h === "LDP Total") nLDPTot = c;
    if (h === "LCP Total") nLCPTot = c;
    if (h.toUpperCase().includes("TBB")) nTBB = c;
  }
  // LCP total column may simply be labelled "Total" (under the "LCP Project" header)
  if (nLCPTot === -1) {
    for (let c = nLDPTot + 1; c < fin.length; c++) {
      const h = fin[c] ? String(fin[c]).trim() : "";
      if (h === "Total") { nLCPTot = c; break; }
    }
  }
  for (let c = 8; c < fin.length; c++) {
    const h = fin[c] ? String(fin[c]).trim() : "";
    if (h === "" || h === "Notes") break;
    if (c === nLDPTot || c === nLCPTot || c === nTBB) continue;
    if (nLDPTot !== -1 && c < nLDPTot) ldp.push(c);
    else if (nLDPTot !== -1 && c > nLDPTot) lcp.push(c);
  }

  const ldpFirst = ldp.length > 0 ? CL(ldp[0]) : "";
  const ldpLast = nLDPTot !== -1 ? CL(nLDPTot - 1) : "";
  const lcpFirst = nLDPTot !== -1 ? CL(nLDPTot + 1) : "";
  const lcpLast = nLCPTot !== -1 ? CL(nLCPTot - 1) : "";

  if (nLDPTot !== -1 && ldpFirst !== "" && ldpLast !== "") {
    wsVT.getRange(`${CL(nLDPTot)}6`).formulas = [[`=SUM(${ldpFirst}6:${ldpLast}6)`]];
    wsVT.getRange(`${CL(nLDPTot)}6`).numberFormat = [[FMT_ACCT]];
  }
  if (nLCPTot !== -1 && lcpFirst !== "" && lcpLast !== "") {
    wsVT.getRange(`${CL(nLCPTot)}6`).formulas = [[`=SUM(${lcpFirst}6:${lcpLast}6)`]];
    wsVT.getRange(`${CL(nLCPTot)}6`).numberFormat = [[FMT_ACCT]];
  }

  if (clientRow !== -1) {
    for (const c of ldp) { wsVT.getRange(`${CL(c)}${clientRow}`).formulas = [[`=${CL(c)}6`]]; wsVT.getRange(`${CL(c)}${clientRow}`).numberFormat = [[FMT_ACCT]]; }
    for (const c of lcp) { wsVT.getRange(`${CL(c)}${clientRow}`).formulas = [[`=${CL(c)}6`]]; wsVT.getRange(`${CL(c)}${clientRow}`).numberFormat = [[FMT_ACCT]]; }
    if (nTBB !== -1) { wsVT.getRange(`${CL(nTBB)}${clientRow}`).formulas = [[`=${CL(nTBB)}6`]]; wsVT.getRange(`${CL(nTBB)}${clientRow}`).numberFormat = [[FMT_ACCT]]; }
    if (nLDPTot !== -1 && ldpFirst !== "" && ldpLast !== "") {
      wsVT.getRange(`${CL(nLDPTot)}${clientRow}`).formulas = [[`=SUM(${ldpFirst}${clientRow}:${ldpLast}${clientRow})`]];
      wsVT.getRange(`${CL(nLDPTot)}${clientRow}`).numberFormat = [[FMT_ACCT]];
    }
    if (nLCPTot !== -1 && lcpFirst !== "" && lcpLast !== "") {
      wsVT.getRange(`${CL(nLCPTot)}${clientRow}`).formulas = [[`=SUM(${lcpFirst}${clientRow}:${lcpLast}${clientRow})`]];
      wsVT.getRange(`${CL(nLCPTot)}${clientRow}`).numberFormat = [[FMT_ACCT]];
    }
  }

  if (subContRow !== -1) {
    const end = clientRow !== -1 ? clientRow - 1 : subContRow - 1;
    const all = [...ldp, ...lcp];
    if (nLDPTot !== -1) all.push(nLDPTot);
    if (nTBB !== -1) all.push(nTBB);
    if (nLCPTot !== -1) all.push(nLCPTot);
    for (const c of all) {
      wsVT.getRange(`${CL(c)}${subContRow}`).formulas = [[`=SUM(${CL(c)}7:${CL(c)}${end})`]];
      wsVT.getRange(`${CL(c)}${subContRow}`).numberFormat = [[FMT_ACCT]];
    }
  }
  await context.sync();

  // Vendor LDP/LCP totals
  let nPO = -1;
  for (let c = 0; c < fin.length; c++) { if (fin[c] && String(fin[c]).trim() === "PO Total") nPO = c; }
  if (nPO !== -1) {
    const po = CL(nPO);
    const vtAr = await readValues(context, wsVT.getRange("A5:A60"));
    const vtB = await readValues(context, wsVT.getRange("B5:B60"));
    // Classify contracts by the bold "LDP"/"LCP" marker rows (NOT by contract name —
    // the LCP contract name may not contain "LCP")
    const ldpContracts: number[] = [], lcpContracts: number[] = [];
    let section = "LDP";
    for (let r = 0; r < vtAr.length; r++) {
      const a = vtAr[r][0] ? String(vtAr[r][0]).trim() : "";
      const b = vtB[r][0];
      if (a === "LDP") { section = "LDP"; continue; }
      if (a === "LCP") { section = "LCP"; continue; }
      if (a.includes("Client total") || a.includes("Client Total") || a.includes("Analysis")) break;
      if (b !== null && b !== "" && Number(b) > 0) {
        if (section === "LCP") lcpContracts.push(r + 5);
        else ldpContracts.push(r + 5);
      }
    }
    const lastRow = clientRow !== -1 ? clientRow - 1 : 30;
    const ldpRangeEnd = nLDPTot !== -1 ? CL(nLDPTot - 1) : "";
    const lcpRangeStart = nLDPTot !== -1 ? CL(nLDPTot + 1) : "";
    const lcpRangeEnd = nLCPTot !== -1 ? CL(nLCPTot - 1) : "";
    const firstLDPCol = ldp.length > 0 ? CL(ldp[0]) : (nLDPTot !== -1 ? CL(nLDPTot - 1) : "");

    const poVals = await readValues(context, wsVT.getRange(`${po}7:${po}${lastRow}`));
    for (let row = 7; row <= lastRow; row++) {
      const poVal = poVals[row - 7] ? poVals[row - 7][0] : "";
      if (poVal === null || poVal === "" || poVal === 0) continue;
      let nearestContract = -1;
      for (const cr of [...ldpContracts, ...lcpContracts].sort((a2, b2) => a2 - b2)) { if (cr <= row) nearestContract = cr; }
      const isLCP = lcpContracts.includes(nearestContract);
      if (!isLCP && nLDPTot !== -1 && firstLDPCol !== "" && ldpRangeEnd !== "") {
        wsVT.getRange(`${CL(nLDPTot)}${row}`).formulas = [[`=${po}${row}+SUM(${firstLDPCol}${row}:${ldpRangeEnd}${row})`]];
        wsVT.getRange(`${CL(nLDPTot)}${row}`).numberFormat = [[FMT_ACCT]];
      }
      if (isLCP && nLCPTot !== -1 && lcpRangeStart !== "" && lcpRangeEnd !== "") {
        wsVT.getRange(`${CL(nLCPTot)}${row}`).formulas = [[`=${po}${row}+SUM(${lcpRangeStart}${row}:${lcpRangeEnd}${row})`]];
        wsVT.getRange(`${CL(nLCPTot)}${row}`).numberFormat = [[FMT_ACCT]];
      }
    }
    await context.sync();
  }

  // Header colours / bold
  let ldpClr = "#C6EFCE", lcpClr = "#FCE4D6";
  let fLDP: Excel.RangeFill | undefined, fLCP: Excel.RangeFill | undefined;
  if (nLDPTot !== -1) { fLDP = wsVT.getRange(`${CL(nLDPTot)}5`).format.fill; fLDP.load("color"); }
  if (nLCPTot !== -1) { fLCP = wsVT.getRange(`${CL(nLCPTot)}5`).format.fill; fLCP.load("color"); }
  await context.sync();
  if (fLDP && fLDP.color) ldpClr = fLDP.color;
  if (fLCP && fLCP.color) lcpClr = fLCP.color;

  for (const c of ldp) { wsVT.getRange(`${CL(c)}5`).format.fill.color = ldpClr; wsVT.getRange(`${CL(c)}5`).format.font.bold = true; }
  if (nLDPTot !== -1) { wsVT.getRange(`${CL(nLDPTot)}5`).format.fill.color = ldpClr; wsVT.getRange(`${CL(nLDPTot)}5`).format.font.bold = true; }
  for (const c of lcp) { wsVT.getRange(`${CL(c)}5`).format.fill.color = lcpClr; wsVT.getRange(`${CL(c)}5`).format.font.bold = true; }
  if (nTBB !== -1) { wsVT.getRange(`${CL(nTBB)}5`).format.fill.color = lcpClr; wsVT.getRange(`${CL(nTBB)}5`).format.font.bold = true; }
  if (nLCPTot !== -1) { wsVT.getRange(`${CL(nLCPTot)}5`).format.fill.color = lcpClr; wsVT.getRange(`${CL(nLCPTot)}5`).format.font.bold = true; }

  wsVT.activate();
  await context.sync();
}

// ───────────────────────── Get PR (load existing lines into the panel) ─────────────────────────

async function runGetPR(context: Excel.RequestContext): Promise<{ desc: string; amt: number | "" }[]> {
  const wsInv = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Invoice Worksheet");
  wsInv.load("name");
  await context.sync();
  if (wsInv.isNullObject) throw new Error("Invoice Worksheet not found.");

  const hdr2 = (await readValues(context, wsInv.getRange("A2").getResizedRange(0, 51)))[0];
  const hdr3 = (await readValues(context, wsInv.getRange("A3").getResizedRange(0, 51)))[0];
  let tdIdx = -1, prTBBIdx = -1;
  for (let c = 0; c < hdr2.length; c++) {
    if (hdr2[c] && String(hdr2[c]).trim() === "Total To date" && tdIdx === -1) tdIdx = c;
  }
  for (let c = 0; c < hdr3.length; c++) {
    if (hdr3[c] && String(hdr3[c]).trim().toUpperCase() === "PR#TBB") { prTBBIdx = c; break; }
  }
  if (tdIdx === -1) throw new Error("'Total To date' not found.");
  if (prTBBIdx === -1) throw new Error("'PR#TBB' column not found.");

  const colA = await readValues(context, wsInv.getRange("A1:A100"));
  let grandRow = -1;
  for (let r = 4; r < 100; r++) {
    if (colA[r][0] && String(colA[r][0]).trim() === "Grand - TOTALS") { grandRow = r + 1; break; }
  }
  if (grandRow === -1) throw new Error("Grand Totals not found.");

  const descR = await readValues(context, wsInv.getRange(`A5:A${grandRow - 1}`));
  const prR = await readValues(context, wsInv.getRange(`${CL(prTBBIdx)}5:${CL(prTBBIdx)}${grandRow - 1}`));

  const out: { desc: string; amt: number | "" }[] = [];
  for (let i = 0; i < descR.length; i++) {
    const d = descR[i][0] ? String(descR[i][0]).trim() : "";
    if (d === "") continue;
    let amt: number | "" = "";
    const pv = prR[i] ? prR[i][0] : "";
    if (pv !== "" && pv !== null) {
      const n = Number(String(pv).replace(/\$/g, "").replace(/,/g, "").trim());
      if (!isNaN(n) && n !== 0) amt = n;
    }
    out.push({ desc: d, amt });
  }
  return out;
}

// ───────────────────────── Latest PR name (for the hint above Column Name) ─────────────────────────

async function runGetLatestPRName(context: Excel.RequestContext): Promise<string> {
  const wsInv = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Invoice Worksheet");
  wsInv.load("name");
  await context.sync();
  if (wsInv.isNullObject) return "";

  const hdr3 = (await readValues(context, wsInv.getRange("A3").getResizedRange(0, 51)))[0] as (string | number)[];
  let tbbCol = -1;
  for (let c = 0; c < hdr3.length; c++) {
    if (hdr3[c] && String(hdr3[c]).trim().toUpperCase() === "PR#TBB") { tbbCol = c; break; }
  }
  if (tbbCol < 1) return "";

  // Scan columns to the LEFT of PR#TBB for PR#N pattern, return highest
  let highest = 0;
  let highestName = "";
  for (let c = 0; c < tbbCol; c++) {
    const v = String(hdr3[c] || "").trim();
    const m = v.match(/^PR#(\d+)/i);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > highest) { highest = n; highestName = v; }
    }
  }
  return highestName;
}

// ───────────────────────── Run Invoice Generate ─────────────────────────

async function runInvoiceGenerate(context: Excel.RequestContext) {
  const sheets = context.workbook.worksheets;
  const wsInv = sheets.getItemOrNullObject("LDP & LCP - Invoice Worksheet");
  const wsTBB = sheets.getItemOrNullObject("PR#TBB");
  const wsPay = sheets.getItemOrNullObject("Payments");
  wsInv.load("name"); wsTBB.load("name"); wsPay.load("name");
  await context.sync();
  if (wsInv.isNullObject || wsTBB.isNullObject) throw new Error("Sheet not found (Invoice Worksheet / PR#TBB).");
  const hasPay = !wsPay.isNullObject;

  // Find columns
  const h2 = (await readValues(context, wsInv.getRange("A2").getResizedRange(0, 30)))[0];
  const h3 = (await readValues(context, wsInv.getRange("A3").getResizedRange(0, 30)))[0];
  let tdIdx = -1, paidIdx = -1, fPRIdx = -1, lPRIdx = -1, tbbIdx = -1;
  for (let c = 0; c < h2.length; c++) {
    const hh = h2[c] ? String(h2[c]).trim() : "";
    const nm = h3[c] ? String(h3[c]).trim() : "";
    if (hh === "Total To date" && tdIdx === -1) tdIdx = c;
    if (hh === "Paid to Date" && paidIdx === -1) paidIdx = c;
    if (hh === "Payment Request" || hh === "Deposit - LDP/LCP") { if (fPRIdx === -1) fPRIdx = c; lPRIdx = c; }
    if (nm.toUpperCase().includes("TBB")) tbbIdx = c;
  }
  if (tdIdx === -1 || fPRIdx === -1) throw new Error("Columns not found.");

  const prEnd = paidIdx !== -1 ? paidIdx : 30;
  const allPRs: { name: string; col: number }[] = [];
  for (let c = tdIdx + 1; c < prEnd; c++) {
    const nm = h3[c] ? String(h3[c]).trim() : "";
    if (nm !== "" && !nm.toUpperCase().includes("TBB")) allPRs.push({ name: nm, col: c });
  }

  // PRs without an existing tab
  sheets.load("items/name");
  await context.sync();
  const existingNames = new Set(sheets.items.map((s) => s.name));
  const toCreate = allPRs.filter((pr) => !existingNames.has(pr.name));

  // Grand Totals
  const invA = await readValues(context, wsInv.getRange("A1:A100"));
  let grandRow = -1;
  for (let r = 4; r < 100; r++) { if (invA[r][0] && String(invA[r][0]).trim() === "Grand - TOTALS") { grandRow = r + 1; break; } }
  if (grandRow === -1) throw new Error("Grand Totals not found.");

  // Bold detection (queue all font loads, one sync)
  const boldFonts: Excel.RangeFont[] = [];
  for (let r = 4; r < grandRow - 1; r++) {
    const f = wsInv.getRange(`A${r + 1}`).format.font;
    f.load("bold");
    boldFonts.push(f);
  }
  await context.sync();

  // Data rows + bold detection (preserve blank rows as gaps)
  const dataRows: { desc: string; invRow: number; isBold: boolean; isBlank: boolean }[] = [];
  let lastWasBlank = true;
  for (let r = 4; r < grandRow - 1; r++) {
    const a = invA[r][0] ? String(invA[r][0]).trim() : "";
    if (a.startsWith("SUB - TOTALS") || a === "Grand - TOTALS") continue;
    if (a === "") {
      if (lastWasBlank) continue;
      dataRows.push({ desc: "", invRow: r + 1, isBold: false, isBlank: true });
      lastWasBlank = true;
      continue;
    }
    const bold = !!boldFonts[r - 4].bold;
    dataRows.push({ desc: a, invRow: r + 1, isBold: bold, isBlank: false });
    lastWasBlank = false;
  }
  while (dataRows.length > 0 && dataRows[dataRows.length - 1].isBlank) dataRows.pop();
  dataRows.unshift({ desc: "", invRow: -1, isBold: false, isBlank: true });
  dataRows.push({ desc: "", invRow: -1, isBold: false, isBlank: true });

  const inv = "'LDP & LCP - Invoice Worksheet'";
  const tdCol = CL(tdIdx);
  const fPR = CL(fPRIdx);
  const tbbC = tbbIdx !== -1 ? CL(tbbIdx) : "";

  // Template structure
  const tbbA = await readValues(context, wsTBB.getRange("A1:A60"));
  const tbbStart = 7;
  let tbbSubRow = -1;
  for (let r = 0; r < tbbA.length; r++) {
    const a = tbbA[r][0] ? String(tbbA[r][0]).trim() : "";
    if (a === "SUB - TOTALS" || a === "SUB-TOTALS") tbbSubRow = r + 1;
  }
  const tplRows = tbbSubRow !== -1 ? tbbSubRow - tbbStart : 10;
  const needed = dataRows.length;

  // Payments
  const payments: { date: string | number | boolean; ptype: string; amount: number }[] = [];
  if (hasPay) {
    const pd = await readValues(context, wsPay.getRange("A2:C50"));
    for (const r of pd) {
      // Just skip empty rows instead of breaking on them — earlier the loop
      // bailed at the first blank row, so any payment the user added BELOW a
      // gap in the Payments sheet was silently ignored. Now we scan the whole
      // A2:C50 range and only keep rows that have both a date and an amount.
      if (r[0] !== null && r[0] !== "" && r[2] !== null && r[2] !== "" && !isNaN(Number(r[2]))) {
        payments.push({ date: r[0], ptype: r[1] ? String(r[1]).trim() : "", amount: Number(r[2]) });
      }
    }
  }

  const fmt = FMT_ACCT;

  // ───────── STEP 1: PR#TBB with formulas ─────────
  if (tbbSubRow !== -1) {
    for (let r = tbbStart; r < tbbSubRow; r++) wsTBB.getRange(`A${r}:M${r}`).values = [["", "", "", "", "", "", "", "", "", "", "", "", ""]];
    await context.sync();
  }

  if (needed > tplRows && tbbSubRow !== -1) {
    const insStart = tbbSubRow;
    const insEnd = tbbSubRow + needed - tplRows - 1;
    wsTBB.getRange(`${insStart}:${insEnd}`).insert(Excel.InsertShiftDirection.down);
    await context.sync();
    // The freshly inserted rows have no borders / fills of their own. Copy the
    // formatting (borders, alignment, fills) from a known-good template row
    // (the last existing data row, just above the original SUB-TOTALS row) so
    // the new rows look identical to the surrounding template rows.
    const templateRow = insStart - 1;
    if (templateRow >= tbbStart) {
      for (let r = insStart; r <= insEnd; r++) {
        wsTBB.getRange(`A${r}:M${r}`).copyFrom(
          wsTBB.getRange(`A${templateRow}:M${templateRow}`),
          Excel.RangeCopyType.formats
        );
      }
      await context.sync();
    }
  } else if (needed < tplRows && tbbSubRow !== -1) {
    wsTBB.getRange(`${tbbStart + needed}:${tbbStart + tplRows - 1}`).delete(Excel.DeleteShiftDirection.up);
    await context.sync();
  }

  const stRow = tbbStart + needed;
  const lastD = stRow - 1;

  for (let i = 0; i < dataRows.length; i++) {
    const t = tbbStart + i;
    const v = dataRows[i].invRow;

    if (dataRows[i].isBlank) {
      wsTBB.getRange(`A${t}:M${t}`).clear(Excel.ClearApplyTo.contents);
      continue;
    }

    // Set Aptos 9pt on the WHOLE A:D range BEFORE merging — Excel's merge logic
    // otherwise tends to inherit the right-most cell's font (Calibri default).
    wsTBB.getRange(`A${t}:D${t}`).format.font.name = "Aptos";
    wsTBB.getRange(`A${t}:D${t}`).format.font.size = 9;
    wsTBB.getRange(`A${t}:D${t}`).merge(false);
    wsTBB.getRange(`A${t}`).values = [[dataRows[i].desc]];
    wsTBB.getRange(`A${t}`).format.horizontalAlignment = Excel.HorizontalAlignment.center;
    // Re-assert font on the merged anchor in case merge clobbered it.
    wsTBB.getRange(`A${t}`).format.font.name = "Aptos";
    wsTBB.getRange(`A${t}`).format.font.size = 9;
    // Mirror the description's bold state from the Invoice Worksheet — if the
    // source row was bold (sub-header) keep it bold here, otherwise normal.
    wsTBB.getRange(`A${t}`).format.font.bold = dataRows[i].isBold;

    // (The previous explicit border-removal here wiped the borders inherited from
    // the template row above — keep the template's borders intact.)

    // Bold sub-header rows (e.g. "Original Contract - LDP", "LCP", "Lot 5
    // Boinapalli Landscape Construction Proposal") are section titles —
    // they shouldn't carry numeric values OR formulas. Earlier the same
    // formulas were written for every row, and bold headers ended up
    // showing $0 on PR#TBB because the Invoice Worksheet's Total To date
    // formula evaluated to 0 (SUM of an empty range). Skip them entirely
    // for bold rows; only the description in column A is kept.
    if (dataRows[i].isBold) {
      wsTBB.getRange(`E${t}:M${t}`).clear(Excel.ClearApplyTo.contents);
    } else {
      wsTBB.getRange(`E${t}`).formulas = [[`=IF(${inv}!${tdCol}${v}="","",${inv}!${tdCol}${v})`]];
      wsTBB.getRange(`F${t}`).values = [[""]];
      wsTBB.getRange(`G${t}`).formulas = [[`=IF(${inv}!${tdCol}${v}="","",${inv}!${tdCol}${v})`]];

      if (tbbIdx !== -1 && fPRIdx !== -1) {
        if (tbbIdx === fPRIdx) {
          wsTBB.getRange(`I${t}`).formulas = [[`=IF(${inv}!${tdCol}${v}="","",0)`]];
        } else {
          wsTBB.getRange(`I${t}`).formulas = [[`=IF(${inv}!${tdCol}${v}="","",SUMIFS(${inv}!$${fPR}${v}:$${CL(tbbIdx - 1)}${v},${inv}!$${fPR}${v}:$${CL(tbbIdx - 1)}${v},"<>"&""))`]];
        }
      }
      wsTBB.getRange(`J${t}`).formulas = [[`=IF(${inv}!${tbbC}${v}="","",${inv}!${tbbC}${v})`]];
      wsTBB.getRange(`K${t}`).formulas = [[`=IF(AND(I${t}="",J${t}=""),"",SUM(I${t},J${t}))`]];
      wsTBB.getRange(`L${t}`).formulas = [[`=IFERROR(K${t}/G${t},"")`]];
      wsTBB.getRange(`M${t}`).formulas = [[`=IF(OR(G${t}="",K${t}=""),"",G${t}-K${t})`]];
    }
  }
  await context.sync();

  wsTBB.getRange(`E${tbbStart}:G${lastD}`).numberFormat = [["\"$\"#,##0.00"]];
  wsTBB.getRange(`I${tbbStart}:K${lastD}`).numberFormat = [[fmt]];
  wsTBB.getRange(`L${tbbStart}:L${lastD}`).numberFormat = [["0%"]];
  wsTBB.getRange(`M${tbbStart}:M${lastD}`).numberFormat = [["\"$\"#,##0.00"]];
  wsTBB.getRange(`J${tbbStart}:J${lastD}`).format.fill.color = "#FFFF00";

  // SUB-TOTALS
  wsTBB.getRange(`A${stRow}:D${stRow}`).merge(false);
  wsTBB.getRange(`A${stRow}`).values = [["SUB - TOTALS"]];
  wsTBB.getRange(`A${stRow}`).format.font.bold = true;
  wsTBB.getRange(`A${stRow}`).format.horizontalAlignment = Excel.HorizontalAlignment.center;
  wsTBB.getRange(`E${stRow}`).formulas = [[`=SUM(E${tbbStart}:E${lastD})`]];
  wsTBB.getRange(`F${stRow}`).formulas = [[`=SUM(F${tbbStart}:F${lastD})`]];
  wsTBB.getRange(`G${stRow}`).formulas = [[`=SUM(G${tbbStart}:G${lastD})`]];
  wsTBB.getRange(`I${stRow}`).formulas = [[`=SUM(I${tbbStart}:I${lastD})`]];
  wsTBB.getRange(`J${stRow}`).formulas = [[`=SUM(J${tbbStart}:J${lastD})`]];
  wsTBB.getRange(`K${stRow}`).formulas = [[`=SUM(K${tbbStart}:K${lastD})`]];
  wsTBB.getRange(`L${stRow}`).formulas = [[`=IFERROR(K${stRow}/G${stRow},0)`]];
  wsTBB.getRange(`M${stRow}`).formulas = [[`=SUM(M${tbbStart}:M${lastD})`]];
  await context.sync();

  // Invoiced to Date + Payments
  let tbbFullA = await readValues(context, wsTBB.getRange("A1:A80"));
  let tbbFullG = await readValues(context, wsTBB.getRange("G1:G80"));
  let existInvToDt = -1;
  for (let r = 0; r < tbbFullA.length; r++) {
    const a = tbbFullA[r][0] ? String(tbbFullA[r][0]).trim() : "";
    if (a === "Invoiced to Date") existInvToDt = r + 1;
  }
  if (existInvToDt === -1) {
    existInvToDt = stRow + 1;
    wsTBB.getRange(`A${existInvToDt}`).values = [["Invoiced to Date"]];
    wsTBB.getRange(`A${existInvToDt}`).format.font.bold = true;
    await context.sync();
  }

  if (payments.length > 0) {
    let g = await readValues(context, wsTBB.getRange("G1:G80"));
    let curTotalPaid = -1;
    for (let r = 0; r < g.length; r++) { if (g[r][0] && String(g[r][0]).includes("Total Paid")) { curTotalPaid = r + 1; break; } }
    if (curTotalPaid !== -1) {
      const existingPayRows = curTotalPaid - existInvToDt - 1;
      if (existingPayRows > 0) {
        wsTBB.getRange(`${existInvToDt + 1}:${curTotalPaid - 1}`).delete(Excel.DeleteShiftDirection.up);
        await context.sync();
      }
    }
    g = await readValues(context, wsTBB.getRange("G1:G80"));
    let newTotalPaid = -1;
    for (let r = 0; r < g.length; r++) { if (g[r][0] && String(g[r][0]).includes("Total Paid")) { newTotalPaid = r + 1; break; } }
    if (newTotalPaid !== -1) {
      const totalInsert = payments.length + 1;
      const insertAt = newTotalPaid;
      wsTBB.getRange(`${insertAt}:${insertAt + totalInsert - 1}`).insert(Excel.InsertShiftDirection.down);
      wsTBB.getRange(`A${insertAt}:M${insertAt + totalInsert - 1}`).clear(Excel.ClearApplyTo.formats);
      for (let i = 0; i < payments.length; i++) {
        const r = insertAt + 1 + i;
        wsTBB.getRange(`G${r}`).values = [[payments[i].date]];
        wsTBB.getRange(`G${r}`).numberFormat = [["mm/dd/yyyy"]];
        wsTBB.getRange(`H${r}`).values = [[payments[i].ptype]];
        wsTBB.getRange(`I${r}`).values = [[payments[i].amount]];
        wsTBB.getRange(`I${r}`).numberFormat = [[fmt]];
      }
      const finalTpRow = newTotalPaid + totalInsert;
      // Total Paid To date = SUM down column I from SUB-TOTALS row through
      // every payment row just above this Total Paid row.
      wsTBB.getRange(`I${finalTpRow}`).formulas = [[`=SUM(I${stRow}:I${finalTpRow - 1})`]];
      wsTBB.getRange(`I${finalTpRow}`).numberFormat = [[fmt]];
      wsTBB.getRange(`G${finalTpRow}:M${finalTpRow}`).format.borders.getItem(Excel.BorderIndex.edgeTop).style = Excel.BorderLineStyle.continuous;
      await context.sync();
    }
  } else {
    const g = await readValues(context, wsTBB.getRange("G1:G80"));
    for (let r = 0; r < g.length; r++) {
      if (g[r][0] && String(g[r][0]).includes("Total Paid")) {
        const totalPaidRow = r + 1;
        // Same SUM range as the "with payments" branch — covers payments the
        // user typed directly into PR#TBB (not via the Payments sheet).
        wsTBB.getRange(`I${totalPaidRow}`).formulas = [[`=SUM(I${stRow}:I${totalPaidRow - 1})`]];
        wsTBB.getRange(`I${totalPaidRow}`).numberFormat = [[fmt]];
        wsTBB.getRange(`G${totalPaidRow}:M${totalPaidRow}`).format.borders.getItem(Excel.BorderIndex.edgeTop).style = Excel.BorderLineStyle.continuous;
        break;
      }
    }
    await context.sync();
  }

  // ───────── STEP 2: static invoice tabs ─────────
  // Batch read invoice columns used for static values
  const invTd = await readValues(context, wsInv.getRange(`${tdCol}5:${tdCol}${grandRow - 1}`));

  for (const pr of toCreate) {
    const existing = sheets.getItemOrNullObject(pr.name);
    existing.load("name");
    await context.sync();
    if (!existing.isNullObject) { existing.delete(); await context.sync(); }

    const ws = wsTBB.copy(Excel.WorksheetPositionType.end);
    ws.name = pr.name;
    await context.sync();

    ws.getRange("M3").values = [[pr.name]];
    const prDate = (await readValues(context, wsInv.getRange(`${CL(pr.col)}4`)))[0][0];
    ws.getRange("M2").values = [[prDate as string | number | boolean]];

    const prVals = await readValues(context, wsInv.getRange(`${CL(pr.col)}5:${CL(pr.col)}${grandRow - 1}`));

    // Track snapshot rows that should be removed entirely (this PR + every
    // earlier PR had nothing for them). Filled inside the per-row loop and
    // applied after SUB-TOTALS / footer freezing so cell references auto-adjust.
    const blankRowsToDelete: number[] = [];

    // PR columns before this one
    const prsBefore: number[] = [];
    for (let c = fPRIdx; c < pr.col; c++) {
      const nm = h3[c] ? String(h3[c]).trim() : "";
      if (nm !== "" && !nm.toUpperCase().includes("TBB")) prsBefore.push(c);
    }
    let beforeBlock: (string | number | boolean)[][] = [];
    if (pr.col > fPRIdx) beforeBlock = await readValues(context, wsInv.getRange(`${CL(fPRIdx)}5:${CL(pr.col - 1)}${grandRow - 1}`));

    // Phase 1: compute every row's state without writing.
    type RowState = {
      t: number; isBold: boolean; gNum: number; gIsEmpty: boolean;
      jVal: number | ""; paidSum: number; hasOwnValue: boolean; keep: boolean;
    };
    const rowStates: RowState[] = [];
    for (let i = 0; i < dataRows.length; i++) {
      if (dataRows[i].isBlank) continue;
      const t = tbbStart + i;
      const idx = dataRows[i].invRow - 5;

      const jRaw = prVals[idx] ? prVals[idx][0] : "";
      const jVal: number | "" = jRaw !== "" && jRaw !== null && jRaw !== 0 ? Number(jRaw) : "";

      let paidSum = 0;
      for (const pc of prsBefore) {
        const off = pc - fPRIdx;
        const val = beforeBlock[idx] ? beforeBlock[idx][off] : "";
        if (val !== null && val !== "" && !isNaN(Number(val))) paidSum += Number(val);
      }

      const gRaw = invTd[idx] ? invTd[idx][0] : "";
      // Track whether the invoice cell is TRULY empty vs literally 0 — this
      // is how we differentiate a bold header (Total To date is empty) from
      // a descriptive zero-value row like "Discussion Points" (Total To date
      // is the numeric 0). PR#TBB shows the former as blank and the latter
      // as $0, so we mirror that on the PR snapshots too.
      const gIsEmpty = gRaw === null || gRaw === "";
      const gNum = gIsEmpty ? 0 : Number(gRaw);

      const noCurrentPR = jVal === "" || jVal === 0;
      const noPriorPR = paidSum === 0;
      const hasOwnValue = !(gNum === 0 || (noCurrentPR && noPriorPR));

      rowStates.push({
        t, isBold: dataRows[i].isBold, gNum, gIsEmpty, jVal, paidSum,
        hasOwnValue, keep: hasOwnValue,
      });
    }

    // Phase 2: group-based keep decision. A "group" is a (possibly empty)
    // cluster of consecutive bold sub-header rows followed by the non-bold
    // rows below it, up to the next bold cluster. If any row in the group
    // has value, the entire group is kept — that pulls in the LDP / LCP
    // headers, sub-titles like "Lot 5 Boinapalli Landscape Construction
    // Proposal", AND descriptive zero-value rows (e.g. "Discussion Points")
    // that belong to a section that's actually being billed.
    {
      let p = 0;
      while (p < rowStates.length) {
        const clusterStart = p;
        while (p < rowStates.length && rowStates[p].isBold) p++;
        const contentEnd_ = (function () {
          let q = p;
          while (q < rowStates.length && !rowStates[q].isBold) q++;
          return q;
        })();
        const groupEnd = contentEnd_;
        let anyValue = false;
        for (let k = clusterStart; k < groupEnd; k++) {
          if (rowStates[k].hasOwnValue) { anyValue = true; break; }
        }
        for (let k = clusterStart; k < groupEnd; k++) {
          rowStates[k].keep = anyValue;
        }
        p = groupEnd;
      }
    }

    // Phase 3: apply the decisions.
    for (const st of rowStates) {
      ws.getRange(`F${st.t}`).values = [[""]];
      if (!st.keep) {
        // Fully blank, then queue for physical removal below.
        ws.getRange(`A${st.t}`).values = [[""]];
        ws.getRange(`E${st.t}`).values = [[""]];
        ws.getRange(`G${st.t}`).values = [[""]];
        ws.getRange(`I${st.t}`).values = [[""]];
        ws.getRange(`J${st.t}`).values = [[""]];
        ws.getRange(`K${st.t}`).values = [[""]];
        ws.getRange(`L${st.t}`).values = [[""]];
        ws.getRange(`M${st.t}`).values = [[""]];
        blankRowsToDelete.push(st.t);
      } else if (st.hasOwnValue) {
        const noCurrentPR = st.jVal === "" || st.jVal === 0;
        const jNum = noCurrentPR ? 0 : (st.jVal as number);
        const kNum = st.paidSum + jNum;
        ws.getRange(`E${st.t}`).values = [[st.gNum]];
        ws.getRange(`G${st.t}`).values = [[st.gNum]];
        // J shows only THIS PR's contribution — blank when this PR didn't
        // bill this row (but the row still renders because earlier PRs did).
        ws.getRange(`J${st.t}`).values = [[jNum !== 0 ? jNum : ""]];
        // Use !== 0 instead of > 0 so credit (negative) values flow through.
        ws.getRange(`I${st.t}`).values = [[st.paidSum !== 0 ? st.paidSum : ""]];
        ws.getRange(`K${st.t}`).values = [[kNum !== 0 ? kNum : ""]];
        ws.getRange(`L${st.t}`).values = [[kNum !== 0 ? kNum / st.gNum : ""]];
        ws.getRange(`M${st.t}`).values = [[st.gNum - kNum]];
      } else if (st.isBold || st.gIsEmpty) {
        // Bold sub-header rows ("Original Contract - LDP", "LCP", project
        // titles like "Lot 5 Boinapalli ...") always render blank in the
        // numeric columns regardless of whether the Invoice Worksheet
        // formula happens to evaluate to 0 — they are not billable line
        // items. Same blank treatment for any non-bold row whose source
        // cell is truly empty.
        ws.getRange(`E${st.t}`).values = [[""]];
        ws.getRange(`G${st.t}`).values = [[""]];
        ws.getRange(`I${st.t}`).values = [[""]];
        ws.getRange(`J${st.t}`).values = [[""]];
        ws.getRange(`K${st.t}`).values = [[""]];
        ws.getRange(`L${st.t}`).values = [[""]];
        ws.getRange(`M${st.t}`).values = [[""]];
      } else {
        // Kept descriptive zero-value row — non-bold, Invoice Worksheet
        // carries a literal 0 in Total To date (e.g. "Discussion Points").
        // Mirror PR#TBB: E / G / I / J / K / M render as $0; L stays
        // blank because it would be a divide-by-zero.
        ws.getRange(`E${st.t}`).values = [[0]];
        ws.getRange(`G${st.t}`).values = [[0]];
        ws.getRange(`I${st.t}`).values = [[0]];
        ws.getRange(`J${st.t}`).values = [[0]];
        ws.getRange(`K${st.t}`).values = [[0]];
        ws.getRange(`L${st.t}`).values = [[""]];
        ws.getRange(`M${st.t}`).values = [[0]];
      }
    }
    await context.sync();

    // Freeze SUB-TOTALS + footer (read recomputed values, write back static)
    const subCells = ["E", "F", "G", "I", "J", "K", "L", "M"];
    const subProxies: { c: string; r: Excel.Range }[] = [];
    for (const c of subCells) { const rng = ws.getRange(`${c}${stRow}`); rng.load("values"); subProxies.push({ c, r: rng }); }
    const gColCopy = ws.getRange("G1:G60"); gColCopy.load("values");
    await context.sync();

    for (const p of subProxies) {
      const val = p.r.values[0][0];
      ws.getRange(`${p.c}${stRow}`).values = [[val !== null ? val : 0]];
    }

    let copyTpRow = -1;
    const gvals = gColCopy.values;
    for (let r = 0; r < gvals.length; r++) { if (gvals[r][0] && String(gvals[r][0]).includes("Total Paid")) { copyTpRow = r + 1; break; } }
    if (copyTpRow !== -1) {
      const foot = ws.getRange(`I${copyTpRow}:M${copyTpRow + 2}`);
      foot.load("values");
      await context.sync();
      // Total Paid To date on the snapshot is a live SUM down the I column
      // from SUB-TOTALS row through every payment row, so edits to payment
      // rows on the snapshot recompute automatically.
      ws.getRange(`I${copyTpRow}`).formulas = [[`=SUM(I${stRow}:I${copyTpRow - 1})`]];
      const mv = foot.values[0][4];
      ws.getRange(`M${copyTpRow}`).values = [[mv !== null ? mv : 0]];
      const m1 = foot.values[1] ? foot.values[1][4] : 0;
      ws.getRange(`M${copyTpRow + 1}`).values = [[m1 !== null ? m1 : 0]];
      const m2 = foot.values[2] ? foot.values[2][4] : 0;
      ws.getRange(`M${copyTpRow + 2}`).values = [[m2 !== null ? m2 : 0]];
    }
    await context.sync();

    // Finally, physically remove the rows we blanked above. Done LAST so the
    // SUB-TOTALS and the Total Paid SUM formula are already in place — Excel
    // adjusts their row references automatically as rows shift up. Deleted in
    // descending order so earlier deletions don't invalidate later indices.
    if (blankRowsToDelete.length > 0) {
      blankRowsToDelete.sort((a, b) => b - a);
      for (const row of blankRowsToDelete) {
        ws.getRange(`${row}:${row}`).delete(Excel.DeleteShiftDirection.up);
      }
      await context.sync();
    }
  }

  wsTBB.activate();
  await context.sync();
  return `PR#TBB updated · ${toCreate.length} new invoice tab(s) created.`;
}

// ───────────────────────── Load Est (read contracts for the dropdowns) ─────────────────────────

async function runLoadEst(context: Excel.RequestContext): Promise<string[]> {
  const wsVT = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Vendor Tracking");
  wsVT.load("name");
  await context.sync();
  if (wsVT.isNullObject) throw new Error("Vendor Tracking sheet not found.");

  const vtA = await readValues(context, wsVT.getRange("A7:A50"));
  const vtB = await readValues(context, wsVT.getRange("B7:B50"));
  const vtE = await readValues(context, wsVT.getRange("E7:E50"));

  const contracts: string[] = [];
  for (let r = 0; r < vtA.length; r++) {
    const a = vtA[r][0] ? String(vtA[r][0]).trim() : "";
    const b = vtB[r][0];
    const e = vtE[r][0] ? String(vtE[r][0]).trim() : "";
    if (a === "" || a.includes("Client total") || a.includes("Client Total") || a.includes("Project Management")
      || a.includes("Sub-Contractor") || a.includes("Percentage") || a.includes("Gross Margin") || a === "Without Estimate") continue;
    if (e !== "" && b !== "" && b !== null && Number(b) > 0) contracts.push(e);
  }
  contracts.push("Without Estimate");
  return contracts;
}

// ───────────────────────── Update PO (Add / Update / Move) ─────────────────────────

// Returns the subset of PO numbers from `formData` that already exist anywhere
// in column E of the Vendor Tracking sheet. Case-insensitive, whitespace-trimmed.
// Empty rows (no vendor / no PO num) are ignored.
async function runFindDuplicatePOs(
  context: Excel.RequestContext,
  formData: (string | number)[][]
): Promise<string[]> {
  const inputPONums: string[] = [];
  for (const r of formData) {
    const vendor = r[1] ? String(r[1]).trim() : "";
    const poNum = r[2] ? String(r[2]).trim() : "";
    if (vendor && poNum) inputPONums.push(poNum);
  }
  if (inputPONums.length === 0) return [];

  const wsVT = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Vendor Tracking");
  wsVT.load("name");
  await context.sync();
  if (wsVT.isNullObject) return [];

  const used = wsVT.getUsedRange();
  used.load("rowCount");
  await context.sync();
  const lastRow = used && used.rowCount ? used.rowCount : 200;

  const colE = await readValues(context, wsVT.getRange(`E1:E${lastRow}`));
  const existing = new Set<string>();
  for (const row of colE) {
    const v = row[0] ? String(row[0]).trim() : "";
    if (v) existing.add(v.toUpperCase());
  }

  const dupes: string[] = [];
  for (const pn of inputPONums) {
    if (existing.has(pn.toUpperCase()) && dupes.indexOf(pn) === -1) dupes.push(pn);
  }
  return dupes;
}

async function runUpdatePo(context: Excel.RequestContext, form: { source: string; target: string; formData: (string | number)[][] }) {
  const wsVT = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Vendor Tracking");
  wsVT.load("name");
  await context.sync();
  if (wsVT.isNullObject) throw new Error("Vendor Tracking sheet not found.");

  const source = form.source.trim();
  const target = form.target.trim();
  const formData = form.formData;

  let hasYMarks = false;
  for (const r of formData) { if (r[0] && String(r[0]).trim().toUpperCase() === "Y") { hasYMarks = true; break; } }

  let action = "added";
  if (hasYMarks && target !== "" && target !== source) { await poDoMove(context, wsVT, source, target, formData); action = "moved"; }
  else if (hasYMarks) { await poDoUpdate(context, wsVT, source, formData); action = "updated"; }
  else { await poDoAdd(context, wsVT, source, formData); action = "added"; }

  await repairAllVendorTrackingFormulas(context, wsVT);
  await updateClientTotal(context, wsVT);
  wsVT.activate();
  await context.sync();
  return `POs ${action} · vendor tracking updated.`;
}

async function poDoAdd(context: Excel.RequestContext, wsVT: Excel.Worksheet, contract: string, formData: (string | number)[][]) {
  const entries: { vendor: string; poNum: string; amount: number; adj: number }[] = [];
  for (const r of formData) {
    const vendor = r[1] ? String(r[1]).trim() : "";
    if (vendor === "") continue;
    entries.push({
      vendor,
      poNum: r[2] ? String(r[2]).trim() : "",
      amount: r[3] !== null && r[3] !== "" && !isNaN(Number(r[3])) ? Number(r[3]) : 0,
      adj: r[4] !== null && r[4] !== "" && !isNaN(Number(r[4])) ? Number(r[4]) : 0,
    });
  }
  if (entries.length === 0) throw new Error("No PO rows entered.");

  const contractRow = await poFindContractRow(context, wsVT, contract);
  if (contractRow === -1) throw new Error(`Contract '${contract}' not found.`);

  const insertAt = (await poFindInsertAfter(context, wsVT, contractRow)) + 1;
  wsVT.getRange(`${insertAt}:${insertAt + entries.length - 1}`).insert(Excel.InsertShiftDirection.down);
  await context.sync();

  for (let i = 0; i < entries.length; i++) {
    const row = insertAt + i;
    const e = entries[i];
    wsVT.getRange(`A${row}`).values = [[e.vendor]];
    wsVT.getRange(`A${row}:H${row}`).format.font.bold = false; // vendor row must be normal, not bold (inherited from contract header)
    if (e.poNum) wsVT.getRange(`E${row}`).values = [[e.poNum]];
    wsVT.getRange(`F${row}`).values = [[e.amount]]; wsVT.getRange(`F${row}`).numberFormat = [[FMT_ACCT]];
    if (e.adj !== 0) { wsVT.getRange(`G${row}`).values = [[e.adj]]; wsVT.getRange(`G${row}`).numberFormat = [[FMT_ACCT]]; }
    wsVT.getRange(`H${row}`).formulas = [[`=SUM(F${row}:G${row})`]]; wsVT.getRange(`H${row}`).numberFormat = [[FMT_ACCT]];
  }
  await context.sync();
}

async function poDoUpdate(context: Excel.RequestContext, wsVT: Excel.Worksheet, contract: string, formData: (string | number)[][]) {
  const contractRow = await poFindContractRow(context, wsVT, contract);
  if (contractRow === -1) throw new Error(`Contract '${contract}' not found.`);

  for (const r of formData) {
    const mark = r[0] ? String(r[0]).trim().toUpperCase() : "";
    const vendor = r[1] ? String(r[1]).trim() : "";
    if (vendor === "") continue;
    if (mark !== "Y") continue;
    const poNum = r[2] ? String(r[2]).trim() : "";
    const amount = r[3] !== null && r[3] !== "" && !isNaN(Number(r[3])) ? Number(r[3]) : 0;
    const adj = r[4] !== null && r[4] !== "" && !isNaN(Number(r[4])) ? Number(r[4]) : 0;
    const vendorRow = await poFindVendorRow(context, wsVT, contractRow, vendor);
    if (vendorRow === -1) continue;
    if (poNum) wsVT.getRange(`E${vendorRow}`).values = [[poNum]];
    wsVT.getRange(`F${vendorRow}`).values = [[amount]]; wsVT.getRange(`F${vendorRow}`).numberFormat = [[FMT_ACCT]];
    wsVT.getRange(`G${vendorRow}`).values = [[adj]]; wsVT.getRange(`G${vendorRow}`).numberFormat = [[FMT_ACCT]];
    wsVT.getRange(`H${vendorRow}`).formulas = [[`=SUM(F${vendorRow}:G${vendorRow})`]]; wsVT.getRange(`H${vendorRow}`).numberFormat = [[FMT_ACCT]];
    await context.sync();
  }
}

async function poDoMove(context: Excel.RequestContext, wsVT: Excel.Worksheet, source: string, target: string, formData: (string | number)[][]) {
  const vendorsToMove: string[] = [];
  for (const r of formData) {
    const mark = r[0] ? String(r[0]).trim().toUpperCase() : "";
    const vendor = r[1] ? String(r[1]).trim() : "";
    if (vendor === "") continue;
    if (mark === "Y") vendorsToMove.push(vendor);
  }
  if (vendorsToMove.length === 0) throw new Error("No PO rows ticked to move.");

  for (let v = vendorsToMove.length - 1; v >= 0; v--) {
    const name = vendorsToMove[v];
    const srcRow = await poFindContractRow(context, wsVT, source);
    if (srcRow === -1) continue;
    const vRow = await poFindVendorRow(context, wsVT, srcRow, name);
    if (vRow === -1) continue;
    const rowData = (await readValues(context, wsVT.getRange(`A${vRow}:H${vRow}`)))[0];
    wsVT.getRange(`${vRow}:${vRow}`).delete(Excel.DeleteShiftDirection.up);
    await context.sync();
    const tgtRow = await poFindContractRow(context, wsVT, target);
    if (tgtRow === -1) continue;
    const insertAt = (await poFindInsertAfter(context, wsVT, tgtRow)) + 1;
    wsVT.getRange(`${insertAt}:${insertAt}`).insert(Excel.InsertShiftDirection.down);
    await context.sync();
    wsVT.getRange(`A${insertAt}:H${insertAt}`).format.font.bold = false; // moved vendor row must not be bold
    const cols = ["A", "B", "C", "D", "E", "F", "G", "H"];
    for (let c = 0; c < rowData.length; c++) {
      if (rowData[c] !== null && rowData[c] !== "") wsVT.getRange(`${cols[c]}${insertAt}`).values = [[rowData[c]]];
    }
    wsVT.getRange(`H${insertAt}`).formulas = [[`=SUM(F${insertAt}:G${insertAt})`]];
    wsVT.getRange(`H${insertAt}`).numberFormat = [[FMT_ACCT]];
    await context.sync();
  }
}

async function updateClientTotal(context: Excel.RequestContext, wsVT: Excel.Worksheet) {
  const vtA = await readValues(context, wsVT.getRange("A5:A80"));
  const s = scanVtSections(vtA);
  writeClientTotalFormulas(wsVT, s);
  await context.sync();
}

async function poFindContractRow(context: Excel.RequestContext, ws: Excel.Worksheet, contract: string): Promise<number> {
  const vtA = await readValues(context, ws.getRange("A5:A80"));
  const vtE = await readValues(context, ws.getRange("E5:E80"));
  if (contract === "Without Estimate") {
    for (let r = 0; r < vtA.length; r++) { if (vtA[r][0] && String(vtA[r][0]).trim() === "Without Estimate") return r + 5; }
    return -1;
  }
  for (let r = 0; r < vtE.length; r++) { if (vtE[r][0] && String(vtE[r][0]).trim() === contract) return r + 5; }
  return -1;
}

async function poFindInsertAfter(context: Excel.RequestContext, ws: Excel.Worksheet, contractRow: number): Promise<number> {
  const vtA = await readValues(context, ws.getRange("A5:A80"));
  const vtE = await readValues(context, ws.getRange("E5:E80"));
  let lastRow = contractRow;
  for (let r = contractRow - 5 + 1; r < vtA.length; r++) {
    const a = vtA[r][0] ? String(vtA[r][0]).trim() : "";
    const e = vtE[r][0] ? String(vtE[r][0]).trim() : "";
    if (a === "" && e === "") break;
    if (a.includes("Client total")) break;
    lastRow = r + 5;
  }
  return lastRow;
}

async function poFindVendorRow(context: Excel.RequestContext, ws: Excel.Worksheet, contractRow: number, vendorName: string): Promise<number> {
  const vtA = await readValues(context, ws.getRange("A5:A80"));
  for (let r = contractRow - 5 + 1; r < vtA.length; r++) {
    const a = vtA[r][0] ? String(vtA[r][0]).trim() : "";
    if (a === "" && r > contractRow - 5 + 2) break;
    if (a.includes("Client total")) break;
    if (a === vendorName) return r + 5;
  }
  return -1;
}

// ───────────────────────── Load Po to Move (read a contract's POs into the panel) ─────────────────────────

async function runLoadMove(context: Excel.RequestContext, source: string): Promise<{ vendor: string; poNum: string; amount: number; adj: number }[]> {
  const wsVT = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Vendor Tracking");
  wsVT.load("name");
  await context.sync();
  if (wsVT.isNullObject) throw new Error("Vendor Tracking sheet not found.");

  const vtA = await readValues(context, wsVT.getRange("A5:A80"));
  const vtE = await readValues(context, wsVT.getRange("E5:E80"));
  const fg = await readValues(context, wsVT.getRange("F5:G80")); // F=amount, G=adj

  let srcRow = -1;
  if (source === "Without Estimate" || source === "Unknown") {
    for (let r = 0; r < vtA.length; r++) {
      const a = vtA[r][0] ? String(vtA[r][0]).trim() : "";
      if (a === source || a === "Without Estimate") { srcRow = r + 5; break; }
    }
  } else {
    for (let r = 0; r < vtE.length; r++) {
      const e = vtE[r][0] ? String(vtE[r][0]).trim() : "";
      if (e === source) { srcRow = r + 5; break; }
    }
  }
  if (srcRow === -1) throw new Error(`'${source}' not found.`);

  const out: { vendor: string; poNum: string; amount: number; adj: number }[] = [];
  const startIdx = srcRow - 5 + 1;
  for (let r = startIdx; r < vtA.length; r++) {
    const a = vtA[r][0] ? String(vtA[r][0]).trim() : "";
    const e = vtE[r][0] ? String(vtE[r][0]).trim() : "";
    if (a === "" && e === "") break;
    if (a.includes("Client total") || a.includes("Client Total")) break;
    const f = fg[r] ? fg[r][0] : "";
    const g = fg[r] ? fg[r][1] : "";
    out.push({ vendor: a, poNum: e, amount: f ? Number(f) : 0, adj: g ? Number(g) : 0 });
  }
  return out;
}

// ───────────────────────── Contract Adjustment (negative adjustment row) ─────────────────────────

async function runContractAdjust(context: Excel.RequestContext, poNumber: string, adjustment: number, description: string) {
  const wsVT = context.workbook.worksheets.getItemOrNullObject("LDP & LCP - Vendor Tracking");
  wsVT.load("name");
  await context.sync();
  if (wsVT.isNullObject) throw new Error("Vendor Tracking sheet not found.");

  const finalAmount = -Math.abs(adjustment);

  const vtE = await readValues(context, wsVT.getRange("E5:E100"));
  let poRow = -1;
  for (let r = 0; r < vtE.length; r++) {
    const e = vtE[r][0] ? String(vtE[r][0]).trim() : "";
    if (e === poNumber) { poRow = r + 5; break; }
  }
  if (poRow === -1) throw new Error(`PO '${poNumber}' not found.`);

  const targetRow = poRow + 1;
  const nextCost = (await readValues(context, wsVT.getRange(`C${targetRow}`)))[0][0];
  if (nextCost === "" || nextCost === null) {
    wsVT.getRange(`${targetRow}:${targetRow}`).insert(Excel.InsertShiftDirection.down);
    await context.sync();
  }

  wsVT.getRange(`A${targetRow}:E${targetRow}`).format.font.bold = false;
  wsVT.getRange(`A${targetRow}`).values = [[description]];
  wsVT.getRange(`B${targetRow}`).clear(Excel.ClearApplyTo.contents);
  wsVT.getRange(`C${targetRow}`).values = [[finalAmount]];
  wsVT.getRange(`C${targetRow}`).numberFormat = [[FMT_ACCT]];
  wsVT.getRange(`D${targetRow}`).clear(Excel.ClearApplyTo.contents);
  wsVT.getRange(`E${targetRow}`).clear(Excel.ClearApplyTo.contents);
  await context.sync();

  await updateClientTotal(context, wsVT);
  wsVT.activate();
  await context.sync();
}
