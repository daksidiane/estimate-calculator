/* global XLSX */

"use strict";

const YANDEX_URL = "";

const money = (n) => Math.round(n).toLocaleString("ru-RU") + " ₽";
const moneyShort = (n) => Math.round(n).toLocaleString("ru-RU");
const moneyText = (n) =>
  typeof n === "number" && isFinite(n)
    ? Math.round(n).toLocaleString("ru-RU") + " ₽"
    : n == null ? "—" : String(n);

/* Маска телефона +7 (XXX) XXX-XX-XX: всегда начинается с +7, ввод мог идти и с 9 */
function phoneMask(v) {
  let d = String(v == null ? "" : v).replace(/\D+/g, "");
  if (d[0] === "7" || d[0] === "8") d = d.slice(1);
  d = d.slice(0, 10);
  let out = "+7";
  if (d.length) out += " (" + d.slice(0, 3);
  if (d.length > 3) out += ") " + d.slice(3, 6);
  if (d.length > 6) out += "-" + d.slice(6, 8);
  if (d.length > 8) out += "-" + d.slice(8, 10);
  return out;
}
const num = (v) => {
  const f = parseFloat(String(v == null ? "" : v).replace(",", "."));
  return isFinite(f) ? f : 0;
};
const byId = (arr, id) => arr.find((x) => String(x.ID) === String(id)) || null;
const esc = (s) =>
  String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[c]));
const today = () => new Date().toISOString().slice(0, 10);
const fmtDate = (iso) =>
  iso
    ? new Date(iso.slice(0, 19)).toLocaleString("ru-RU", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
      })
    : "—";
const euro = (s) => s.split(" ").filter(Boolean).join(" ").trim();

const CATEGORY_LABEL = {
  GLASS: "Стекло",
  PROFILE: "Профиль",
  HARDWARE: "Фурнитура",
};
const CATEGORY_ICON = { GLASS: "🪟", PROFILE: "▭", HARDWARE: "⚙" };

/* ================= state ================= */

let state = {
  construction: null,
  width: "1.2",
  height: "1.4",
  qty: "1",
  include: { glass: true, profile: true, hardware: true },
  glassId: null,
  profileId: null,
  hardware: [],
  addWorks: [],
  services: [],
  coeffs: [],
  manualDiscount: 0,
  autoDiscount: true,
  repeatClient: false,
  client: "",
  phone: "",
  wishes: "",
};

let DB = null;
let dbFrom = "bundled";
let appInfo = { version: "1.0.0" };
let history = [];
let prevTotal = null;

/* ================= helpers ================= */

const $ = (id) => document.getElementById(id);

const els = (id, html) => {
  $(id).innerHTML = html;
};

function active(arr) {
  return (arr || []).filter((x) => x.IsActive !== false);
}

function defaultStateFor(db) {
  const cons = active(db.constructions);
  const glass = active(db.components.GLASS);
  const prof = active(db.components.PROFILE);
  return {
    construction: cons.length ? cons[0].Code : null,
    width: "1.2",
    height: "1.4",
    qty: "1",
    include: { glass: true, profile: true, hardware: true },
    glassId: glass.length ? String(glass[0].ID) : null,
    profileId: prof.length ? String(prof[0].ID) : null,
    hardware: [],
    addWorks: [],
    services: [],
    coeffs: [],
    manualDiscount: 0,
    autoDiscount: true,
    repeatClient: false,
    client: "",
    phone: "",
    wishes: "",
  };
}

/* Сохраняет пользовательские параметры при обновлении базы, убирает только то,
   чего больше нет в новой версии прайса. */
function preserveStateAcrossDb(newDb, old) {
  const s = Object.assign({}, old);
  const cons = active(newDb.constructions);
  if (!cons.some((c) => c.Code === s.construction)) {
    s.construction = cons.length ? cons[0].Code : null;
  }
  const has = (cat, id) => (newDb.components[cat] || []).some((x) => String(x.ID) === String(id));
  const awHas = (id) => (newDb.additionalWorks || []).some((x) => String(x.ID) === String(id));
  const svHas = (id) => (newDb.services || []).some((x) => String(x.ID) === String(id));
  const cHas = (id) => (newDb.coefficients || []).some((x) => String(x.ID) === String(id));
  if (!has("GLASS", s.glassId)) {
    const glass = active(newDb.components.GLASS);
    s.glassId = glass.length ? String(glass[0].ID) : null;
  }
  if (!has("PROFILE", s.profileId)) {
    const prof = active(newDb.components.PROFILE);
    s.profileId = prof.length ? String(prof[0].ID) : null;
  }
  s.hardware = (s.hardware || []).filter((id) => has("HARDWARE", id));
  s.addWorks = (s.addWorks || []).filter((x) => awHas(x.id));
  s.services = (s.services || []).filter((x) => svHas(x.id));
  s.coeffs = (s.coeffs || []).filter((id) => cHas(id));
  return s;
}

/* Список изменённых цен между старой и новой версией базы */
function priceChanges(prev, next) {
  const changes = [];
  if (!prev) return changes;
  const prevCats = prev.components || {};
  const nextCats = next.components || {};
  for (const cat of Object.keys(nextCats)) {
    for (const c of nextCats[cat] || []) {
      const p = (prevCats[cat] || []).find((x) => String(x.ID) === String(c.ID));
      if (p && Number(p.Price) !== Number(c.Price)) {
        changes.push({ name: c.Name, unit: c.Unit || "", from: Number(p.Price), to: Number(c.Price) });
      }
    }
  }
  const prevCon = prev.constructions || [];
  const nextCon = next.constructions || [];
  for (const code of new Set([...prevCon.map((c) => c.Code), ...nextCon.map((c) => c.Code)])) {
    if (!prevCon.some((c) => c.Code === code)) changes.push({ name: "Тип «" + code + "»", from: "нет", to: "добавлен" });
    if (!nextCon.some((c) => c.Code === code)) changes.push({ name: "Тип «" + code + "»", from: "был", to: "удалён" });
  }
  return changes;
}

function shortDriveTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  return d.toLocaleDateString("ru-RU") + " " + d.toLocaleTimeString("ru-RU", { hour: "2-digit", minute: "2-digit" });
}

function constructionInfo(code) {
  return (DB.constructions || []).find((c) => c.Code === code) || null;
}

function wishesTokens() {
  const t = state.wishes.toLowerCase().replace(/[.,;:%"']/g, " ");
  return t.split(/\s+/).filter((w) => w.length >= 4);
}

function componentMatches(comp, tokens) {
  if (!tokens.length) return true;
  const hay = (comp.Name + " " + (comp.Description || "") + " " + comp.Code).toLowerCase();
  return tokens.some((t) => hay.includes(t));
}

/* ================= calculation engine ================= */

function compute(db, s) {
  const cons = constructionInfo(s.construction);
  const w = num(s.width);
  const h = num(s.height);
  const q = Math.max(1, Math.round(num(s.qty)) || 1);
  const area = w * h;
  const perim = 2 * (w + h);
  const rows = [];

  let glass = 0;
  if (s.include.glass && s.glassId != null) {
    const g = byId(db.components.GLASS || [], s.glassId);
    if (g) {
      const rate = num(g.Price);
      glass = area * rate;
      rows.push({ label: "Стекло", sub: g.Name + " · " + g.Description, unit: "м²", qty: area * q, rate, amount: glass * q });
    }
  }

  let profile = 0;
  if (s.include.profile && s.profileId != null) {
    const p = byId(db.components.PROFILE || [], s.profileId);
    if (p) {
      const rate = num(p.Price);
      profile = perim * rate;
      rows.push({ label: "Профиль", sub: p.Name + " · " + p.Description, unit: "пог.м", qty: perim * q, rate, amount: profile * q });
    }
  }

  let hardware = 0;
  const hwNames = [];
  if (s.include.hardware) {
    for (const id of s.hardware) {
      const hw = byId(db.components.HARDWARE || [], id);
      if (hw) {
        const rate = num(hw.Price);
        hardware += rate;
        hwNames.push(hw.Name);
      }
    }
  }
  if (hardware > 0) {
    rows.push({ label: "Фурнитура", sub: hwNames.join(" · "), unit: "компл.", qty: q, rate: null, amount: hardware * q });
  }

  const allCoefs = [];
  const mountCoefs = [];
  for (const id of s.coeffs) {
    const c = byId(db.coefficients, id);
    if (c) (c.ApplicableTo === "MOUNTING" ? mountCoefs : allCoefs).push(c);
  }
  const mulAll = allCoefs.reduce((acc, c) => acc * num(c.Coefficient), 1);
  const mulMount = mountCoefs.reduce((acc, c) => acc * num(c.Coefficient), 1);

  const constructionBase = (glass + profile + hardware) * q * mulAll;
  if (mulAll !== 1) {
    rows.push({
      label: "Коэффициент констр.",
      sub: allCoefs.map((c) => c.Name + " ×" + c.Coefficient).join(" · "),
      unit: "",
      qty: 1,
      rate: null,
      amount: constructionBase - (glass + profile + hardware) * q,
      kind: "coef",
    });
  }

  const awRows = [];
  let addWorks = 0;
  for (const sel of s.addWorks) {
    const a = byId(db.additionalWorks, sel.id);
    if (!a) continue;
    const rate = num(a.Price);
    const qty = sel.qty == null ? 1 : num(sel.qty);
    const amount = rate * qty;
    addWorks += amount;
    awRows.push({ label: "Доп. работы", sub: a.Name, unit: a.Unit || "", qty, rate, amount });
  }

  const svRows = [];
  let services = 0;
  let mountBase = 0;
  const mountingCodes = ["SRV-005", "SRV-006", "SRV-007", "SRV-008"];
  for (const sel of s.services) {
    const sv = byId(db.services, sel.id);
    if (!sv) continue;
    let rate = num(sv.Price);
    if (sv.CalculationType === "FIXED_MIN") rate = Math.max(rate, num(sv.MinPrice));
    const qty = sel.qty == null ? 1 : num(sel.qty);
    const amount = rate * qty;
    services += amount;
    const isMount = mountingCodes.includes(sv.Code);
    if (isMount) mountBase += amount;
    svRows.push({ label: "Услуги", sub: sv.Name, unit: sv.Unit || "", qty, rate, amount });
  }
  const mountBonus = mountBase * mulMount - mountBase;
  if (mountBonus !== 0) {
    svRows.push({
      label: "Коэфф. на монтаж",
      sub: mountCoefs.map((c) => c.Name + " ×" + c.Coefficient).join(" · ") || "коэффициент",
      unit: "",
      qty: 1,
      rate: null,
      amount: mountBonus,
      kind: "coef",
    });
    services += mountBonus;
  }

  const base = constructionBase + addWorks + services;

  /* discounts */
  let autoPct = 0;
  let appliedRule = null;
  if (s.autoDiscount) {
    let best = null;
    for (const r of db.discounts || []) {
      if (r.IsActive === false) continue;
      if (r.ValidFrom && today() < r.ValidFrom) continue;
      if (r.ValidTo && today() > r.ValidTo) continue;
      if (r.MinAmount != null && base < num(r.MinAmount)) continue;
      if (r.MaxAmount != null && base > num(r.MaxAmount)) continue;
      if (r.ApplicableTo !== "ALL" && r.ApplicableTo !== cons.Code) continue;
      const name = (r.Name + " " + (r.Description || "")).toLowerCase();
      if (name.includes("3+") && q < 3) continue;
      if (/постоянн/i.test(r.Name) && !s.repeatClient) continue;
      if (!best || num(r.DiscountPercent) > num(best.DiscountPercent)) best = r;
    }
    if (best) {
      autoPct = num(best.DiscountPercent);
      appliedRule = best;
    }
  }
  const useManual = num(s.manualDiscount) > 0;
  const pct = Math.min(100, Math.max(0, useManual ? num(s.manualDiscount) : autoPct));
  const disc = (base * pct) / 100;
  const total = base - disc;

  return {
    cons,
    rows: rows.concat(awRows, svRows),
    glass,
    profile,
    hardware,
    area,
    perim,
    q,
    addWorks,
    services,
    constructionBase,
    base,
    pct,
    disc,
    total,
    appliedRule,
    useManual,
    autoPct,
  };
}

/* ================= DB update (Yandex Disk) ================= */

function sheetToRows(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
}

function normalizeDb(payload, wb) {
  const db = {
    updatedAt: new Date().toISOString(),
    yandexUrl: YANDEX_URL,
    sourceFile: "calculator_database.xlsx",
  };
  db.constructions = sheetToRows(wb, "Constructions");
  const comps = sheetToRows(wb, "Components");
  const grouped = {};
  for (const c of comps) {
    const cat = c.Category || "OTHER";
    (grouped[cat] = grouped[cat] || []).push(c);
  }
  db.components = grouped;
  db.additionalWorks = sheetToRows(wb, "AdditionalWorks");
  db.services = sheetToRows(wb, "Services");
  db.discounts = sheetToRows(wb, "Discounts");
  db.coefficients = sheetToRows(wb, "CalculationCoefficients");
  db.formulas = sheetToRows(wb, "Formulas");
  db.units = sheetToRows(wb, "Units");
  db._payload = payload;
  Object.keys(db).forEach((k) => (db[k] == null ? delete db[k] : null));
  return db;
}

async function updateDatabase() {
  const btn = $("btnRefresh");
  btn.disabled = true;
  btn.classList.add("spinning");
  setDbStatus("busy", "Скачиваем актуальную базу…");
  try {
    const res = await window.api.downloadDb();
    if (!res || !res.buffer) throw new Error("Пустой ответ от сервера");
    const prevDb = DB;
    const wb = XLSX.read(new Uint8Array(res.buffer), { type: "array" });
    const db = normalizeDb(null, wb);
    if (!db.constructions || !db.components || !db.constructions.length)
      throw new Error("Не удалось распознать структуру базы");
    db.driveMeta = res.drive || { modified: null, md5: null };
    await window.api.saveDb(db);
    DB = db;
    dbFrom = "remote";
    state = preserveStateAcrossDb(DB, state);
    renderDbStatus();
    renderAll();
    const changes = priceChanges(prevDb, DB);
    const driveTxt = res.drive && res.drive.modified
      ? "файл на Диске изменён " + shortDriveTime(res.drive.modified)
      : "";
    let msg = `База обновлена · скачано ${res.bytes} байт` + (driveTxt ? " · " + driveTxt : "");
    let details = "";
    if (changes.length) {
      details = changes
        .map((c) => `<div>• <b>${esc(c.name)}</b>: ${esc(moneyText(c.from))} → <b>${esc(moneyText(c.to))}</b></div>`)
        .join("");
      msg += ` · <b>${changes.length} изм.</b>`;
    } else {
      msg += " · изменений цен нет";
    }
    toast(msg + (details ? ` <span class="toast-details">${details}</span>` : ""));
  } catch (e) {
    toast("Не удалось обновить базу: " + esc(e.message), true);
  } finally {
    btn.disabled = false;
    btn.classList.remove("spinning");
    renderDbStatus();
  }
}

/* ================= rendering ================= */

function renderDbStatus() {
  const ok = dbFrom === "remote";
  $("dbStatus").classList.toggle("ok", ok);
  $("dbStatus").classList.toggle("busy", false);
  const d = DB ? new Date(DB.updatedAt).toLocaleDateString("ru-RU") : "—";
  let txt;
  if (ok) {
    const m = DB.driveMeta && DB.driveMeta.modified;
    txt = m
      ? "База обновлена · файл изменён " + shortDriveTime(m)
      : "База обновлена · " + d;
  } else {
    txt = "База встроенная · " + d;
  }
  $("dbStatusText").textContent = txt;
}

function setDbStatus(mode, text) {
  $("dbStatus").classList.toggle("ok", mode === "ok");
  $("dbStatus").classList.toggle("busy", mode === "busy");
  $("dbStatusText").textContent = text;
}

function renderConstructions() {
  const list = active(DB.constructions);
  els(
    "constructionList",
    list
      .map(
        (c, i) => `
      <div class="chip-card ${state.construction === c.Code ? "active" : ""}"
           style="animation-delay:${i * 40}ms" data-cons="${esc(c.Code)}">
        <div class="name">${esc(c.Name)}</div>
        <div class="desc">${esc(c.Description || "")}</div>
      </div>`
      )
      .join("")
  );
  document.querySelectorAll("#constructionList .chip-card").forEach((el) =>
    el.addEventListener("click", () => {
      state.construction = el.dataset.cons;
      renderConstructions();
      renderAll();
    })
  );
}

function renderComponents() {
  const groups = ["GLASS", "PROFILE", "HARDWARE"];
  const tokens = wishesTokens();
  els(
    "componentGroups",
    groups
      .map((cat) => {
        const opts = active(DB.components[cat] || []).filter((c) => componentMatches(c, tokens));
        const label = CATEGORY_LABEL[cat] || cat;
        const icon = CATEGORY_ICON[cat] || "";
        const inc = state.include[cat.toLowerCase()];
        const isHardware = cat === "HARDWARE";
        const isOpen = inc;
        return `
      <div class="comp-group ${isOpen ? "open" : ""}" data-cat="${cat}">
        <div class="comp-head" data-toggle>
          <span class="title">${icon} ${label}</span>
          <span class="count">${opts.length} из ${active(DB.components[cat] || []).length}</span>
          <span class="chev">▾</span>
          <label class="switch" title="Включить/выключить">
            <input type="checkbox" ${inc ? "checked" : ""} data-include />
            <span class="knob"></span>
          </label>
        </div>
        <div class="comp-body ${isOpen ? "" : "hidden"}">
          ${
            tokens.length
              ? `<div class="tagbar" style="margin-bottom:10px">Подобрано по пожеланиям: ${tokens
                  .map((t) => `<span class="tag">${esc(t)}</span>`)
                  .join("")}</div>`
              : ""
          }
          <input type="text" class="comp-search" placeholder="Поиск по ${label.toLowerCase()}…" data-search />
          <div class="comp-options" data-options>
            ${opts
              .map((c, i) => {
                const cid = String(c.ID);
                const sel = isHardware
                  ? state.hardware.includes(cid)
                  : state[cat.toLowerCase() + "Id"] === cid;
                const radioCls = isHardware ? "radio square" : "radio";
                return `
              <div class="opt ${sel ? "sel" : ""}" data-id="${esc(cid)}" style="animation-delay:${i * 22}ms">
                <span class="${radioCls}"></span>
                <span>
                  <span class="m-name">${esc(c.Name)}</span>
                  <span class="m-desc">${esc(c.Description || "")}</span>
                </span>
                <span class="price">${money(num(c.Price))}<span class="unit"> / ${esc(c.Unit || "")}</span></span>
              </div>`;
              })
              .join("") || '<div class="hint">Ничего не найдено</div>'}
          </div>
        </div>
      </div>`;
      })
      .join("")
  );

  document.querySelectorAll("#componentGroups .comp-group").forEach((group) => {
    const cat = group.dataset.cat;
    const key = cat.toLowerCase();
    const toggleBtn = group.querySelector("[data-toggle]");
    const include = group.querySelector("[data-include]");
    const search = group.querySelector("[data-search]");
    const options = group.querySelector("[data-options]");

    const toggleOpen = () => {
      group.classList.toggle("open");
      group.querySelector(".comp-body").classList.toggle("hidden");
    };
    toggleBtn.addEventListener("click", (e) => {
      if (e.target.closest(".switch")) return;
      toggleOpen();
    });
    include.addEventListener("change", () => {
      state.include[key] = include.checked;
      renderComponents();
      renderSummary();
    });

    function applyFilter() {
      const q = search.value.trim().toLowerCase();
      options.querySelectorAll(".opt").forEach((el) => {
        const match = !q || el.textContent.toLowerCase().includes(q);
        el.style.display = match ? "" : "none";
      });
    }
    search.addEventListener("input", applyFilter);

    options.querySelectorAll(".opt").forEach((el) => {
      el.addEventListener("click", () => {
        const id = el.dataset.id;
        if (cat === "HARDWARE") {
          const set = new Set(state.hardware);
          set.has(id) ? set.delete(id) : set.add(id);
          state.hardware = [...set];
        } else {
          state[key + "Id"] = state[key + "Id"] === id ? null : id;
        }
        renderComponents();
        renderSummary();
      });
    });
  });
}

function defaultAddQty(a, perim, area) {
  if (a.CalculationType === "PER_METER") return perim;
  if (a.CalculationType === "PER_SQM") return area;
  return 1;
}

function renderAddWorks(c) {
  const perim = c.perim;
  const area = c.area;
  const list = active(DB.additionalWorks);
  const selMap = new Map(state.addWorks.map((x) => [String(x.id), x]));
  els(
    "addWorkList",
    list
      .map((a, i) => {
        const sel = selMap.get(String(a.ID));
        const checked = !!sel;
        const qty = sel ? sel.qty : defaultAddQty(a, perim, area);
        const rate = num(a.Price);
        return `
      <div class="opt ${checked ? "sel" : ""}" data-id="${esc(String(a.ID))}" style="animation-delay:${i * 18}ms" >
        <span class="radio square"></span>
        <span style="min-width:0">
          <span class="m-name">${esc(a.Name)}</span>
          <span class="m-desc">${esc(a.Description || "")}</span>
          <span class="m-desc">${money(rate)} / ${esc(a.Unit || "")}</span>
        </span>
        <span class="opt-row" style="gap:8px">
          <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:var(--text-dim)">
            <span>${esc(a.Unit || "")}:</span>
            <input type="number" class="qty" data-qty min="0" step="0.01" value="${qty}" ${checked ? "" : "disabled"} />
          </label>
        </span>
      </div>`;
      })
      .join("")
  );

  document.querySelectorAll("#addWorkList .opt").forEach((el) => {
    el.addEventListener("click", (e) => {
      const isQty = e.target.closest("[data-qty]");
      if (e.target.closest("input")) return;
      if (isQty) return;
      toggleAddWork(el);
    });
    const qtyEl = el.querySelector("[data-qty]");
    qtyEl.addEventListener("click", (e) => e.stopPropagation());
    qtyEl.addEventListener("input", () => {
      upsertAddWork(el.dataset.id, qtyEl.value);
    });
  });
}

function toggleAddWork(el) {
  const id = el.dataset.id;
  const idx = state.addWorks.findIndex((x) => String(x.id) === id);
  if (idx >= 0) {
    state.addWorks.splice(idx, 1);
  } else {
    const a = byId(DB.additionalWorks, id);
    state.addWorks.push({
      id,
      qty: defaultAddQty(a, 2 * (num(state.width) + num(state.height)), num(state.width) * num(state.height)),
    });
  }
  renderAll();
}

function upsertAddWork(id, qty) {
  const idx = state.addWorks.findIndex((x) => String(x.id) === String(id));
  if (idx >= 0) state.addWorks[idx].qty = num(qty);
  renderSummary();
}

function renderServices(c) {
  const list = active(DB.services);
  const selMap = new Map(state.services.map((x) => [String(x.id), x]));
  els(
    "serviceList",
    list
      .map((s, i) => {
        const sel = selMap.get(String(s.ID));
        const checked = !!sel;
        let qty = sel ? sel.qty : s.CalculationType === "PER_SQM" ? c.area : 1;
        if (s.CalculationType === "FIXED_MIN") qty = 1;
        let rate = num(s.Price);
        if (s.CalculationType === "FIXED_MIN") rate = Math.max(rate, num(s.MinPrice));
        return `
      <div class="opt ${checked ? "sel" : ""}" data-id="${esc(String(s.ID))}" style="animation-delay:${i * 18}ms">
        <span class="radio square"></span>
        <span style="min-width:0">
          <span class="m-name">${esc(s.Name)}</span>
          <span class="m-desc">${esc(s.Description || "")}</span>
          <span class="m-desc">${money(rate)} / ${esc(s.Unit || "")}</span>
        </span>
        <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:var(--text-dim)">
          <span>${esc(s.Unit || "")}:</span>
          <input type="number" class="qty" data-qty min="0" step="0.01" value="${qty}" ${checked ? "" : "disabled"} />
        </label>
      </div>`;
      })
      .join("")
  );

  document.querySelectorAll("#serviceList .opt").forEach((el) => {
    el.addEventListener("click", (e) => {
      if (e.target.closest("input")) return;
      toggleService(el);
    });
    const qtyEl = el.querySelector("[data-qty]");
    qtyEl.addEventListener("click", (e) => e.stopPropagation());
    qtyEl.addEventListener("input", () => upsertService(el.dataset.id, qtyEl.value));
  });
}

function toggleService(el) {
  const id = el.dataset.id;
  const idx = state.services.findIndex((x) => String(x.id) === id);
  if (idx >= 0) {
    state.services.splice(idx, 1);
  } else {
    const sv = byId(DB.services, id);
    const c = compute(DB, state);
    state.services.push({ id, qty: sv.CalculationType === "PER_SQM" ? c.area : 1 });
  }
  renderAll();
}

function upsertService(id, qty) {
  const idx = state.services.findIndex((x) => String(x.id) === String(id));
  if (idx >= 0) state.services[idx].qty = num(qty);
  renderSummary();
}

function renderCoeffs() {
  const list = active(DB.coefficients);
  els(
    "coeffList",
    list
      .map(
        (c, i) => `
      <button class="chip ${state.coeffs.includes(String(c.ID)) ? "active" : ""}" data-id="${esc(String(c.ID))}" style="animation-delay:${i * 35}ms">
        <span class="check">✓</span>
        ${esc(c.Name)} <b>×${c.Coefficient}</b>
      </button>`
      )
      .join("")
  );
  document.querySelectorAll("#coeffList .chip").forEach((el) =>
    el.addEventListener("click", () => {
      const id = el.dataset.id;
      const set = new Set(state.coeffs);
      set.has(id) ? set.delete(id) : set.add(id);
      state.coeffs = [...set];
      renderCoeffs();
      renderSummary();
    })
  );
}

function renderWishTags() {
  const tokens = wishesTokens();
  els(
    "wishTags",
    tokens.map((t, i) => `<span class="tag" style="animation-delay:${i * 40}ms">#${esc(t)}</span>`).join("")
  );
}

function renderDiscounts(c) {
  const rules = active(DB.discounts);
  els(
    "discountRules",
    rules
      .map((r) => {
        let applied = c.appliedRule && c.appliedRule.ID === r.ID && !c.useManual;
        return `<span class="rule-chip ${applied ? "applied" : ""}">${esc(r.Name)} · −${num(r.DiscountPercent)}%</span>`;
      })
      .join("")
  );
  $("inManualDiscount").value = state.manualDiscount;
  if (c.appliedRule) {
    $("discountHint").textContent = c.useManual
      ? `Ручная скидка −${c.pct}% (${money(c.disc)})`
      : `Автоскидка «${c.appliedRule.Name}» −${c.pct}% · сумма скидки ${money(c.disc)}`;
  } else {
    $("discountHint").textContent = "Правила скидок из базы применяются автоматически к сумме заказа.";
  }
}

function renderGroupsAndSummary() {
  const c = compute(DB, state);
  renderAddWorks(c);
  renderServices(c);
  renderSummary();
}

function renderSummary() {
  const c = compute(DB, state);
  const prev = prevTotal;
  prevTotal = c.total;
  $("totalValue").textContent = moneyShort(c.total).replace(/\s/g, "\u00A0") + "\u00A0₽";
  $("totalSub").textContent = `${esc(c.cons.Name)} · ${c.q} шт · ${c.area.toFixed(2)} м²/шт`;

  if (prev != null && Math.abs(prev - c.total) > 0.5) {
    const delta = (c.total - prev) / prev * 100;
    const el = $("totalDelta");
    el.classList.toggle("hidden", false);
    el.className = "total-delta " + (delta >= 0 ? "up" : "down");
    el.textContent = (delta >= 0 ? "▲" : "▼") + " " + Math.abs(delta).toFixed(1) + "%";
  } else {
    $("totalDelta").classList.add("hidden");
  }

  const rows = [
    { l: "Площадь", v: (c.area * c.q).toFixed(2) + " м²", badge: "×" + c.q },
    ...c.rows.map((r) => ({
      l: r.label,
      v: money(r.amount),
      badge: r.kind === "coef" ? r.sub : r.sub ? r.sub + (r.qty ? " · " + r.qty + " " + r.unit : "") : r.qty ? r.qty + " " + r.unit : "",
      coef: r.kind,
    })),
  ];

  const base = c.base;
  rows.push({ l: "Сумма до скидки", v: money(base), badge: "" });
  if (c.disc > 0) rows.push({ l: "Скидка −" + c.pct + "%", v: "− " + money(c.disc), badge: "", minus: true });
  rows.push({ l: "ИТОГО", v: money(c.total), badge: c.cons.Name, total: true });

  els(
    "breakdown",
    rows
      .map(
        (r, i) => `
      <div class="bd-row ${r.total ? "total" : ""}" style="animation-delay:${i * 25}ms">
        <span class="l ${r.minus ? "minus" : ""}">${esc(r.l)}</span>
        <span class="v ${r.total ? "" : ""}">${esc(r.v)}</span>
        ${r.badge ? `<span class="badge">${esc(r.badge)}</span>` : ""}
      </div>`
      )
      .join("")
  );

  renderDiscounts(c);
}

function renderAll() {
  if (!DB) return;
  renderConstructions();
  renderComponents();
  renderWishTags();
  renderCoeffs();
  renderGroupsAndSummary();
}

/* ================= history ================= */

function snapshot() {
  return JSON.parse(JSON.stringify(state));
}

function restoreState(s) {
  Object.keys(state).forEach((k) => delete state[k]);
  Object.assign(state, s);
  $("inPhone").value = phoneMask(state.phone);
  state.phone = $("inPhone").value;
  renderAll();
  const form = document.querySelector(".panel.form");
  if (form) form.scrollTo({ top: 0, behavior: "smooth" });
}

function renderHistory() {
  els(
    "historyList",
    `<div class="history-title">История расчётов (${history.length})</div>` +
      (history.length
        ? history
            .map(
              (h, i) => `
          <div class="hist-item" data-i="${i}">
            <div style="min-width:0">
              <div class="who">${esc(h.client || "Без клиента")} · ${esc((DB && h.consName) || "")}</div>
              <div class="what">${esc(h.wishes || "")}</div>
              <div class="when">${fmtDate(h.ts)}</div>
            </div>
            <div class="sum">${money(h.total)}</div>
            <button class="btn btn-icon del" data-del="${i}" title="Удалить" style="padding:4px">✕</button>
          </div>`
            )
            .join("")
        : '<div class="hist-empty">Пока нет сохранённых расчётов.</div>')
  );
  document.querySelectorAll("#historyList .hist-item:not(.hist-empty)").forEach((el) => {
    const i = el.dataset.i;
    el.addEventListener("click", (e) => {
      if (e.target.closest(".del")) return;
      restoreState(JSON.parse(JSON.stringify(history[i].state)));
      toast("Расчёт загружен из истории");
    });
    el.querySelector(".del").addEventListener("click", () => {
      history.splice(+i, 1);
      persistHistory();
      renderHistory();
    });
  });
}

async function persistHistory() {
  try {
    await window.api.saveHistory(history);
  } catch (_) {}
}

async function saveCurrent() {
  const c = compute(DB, state);
  const cons = c.cons ? c.cons.Name : "";
  history.unshift({
    id: Date.now(),
    ts: new Date().toISOString(),
    client: state.client.trim(),
    phone: state.phone.trim(),
    wishes: state.wishes.trim(),
    cons: state.construction,
    consName: cons,
    total: c.total,
    state: snapshot(),
  });
  if (history.length > 200) history.length = 200;
  await persistHistory();
  renderHistory();
  toast("Расчёт сохранён в истории");
}

/* ================= PDF / print ================= */

function buildPrint() {
  const c = compute(DB, state);
  const cons = c.cons ? c.cons.Name : "—";
  const areaStr = (c.area * c.q).toFixed(2);
  const rows = c.rows.map((r) => ({
    name: r.label + (r.sub ? " — " + r.sub : ""),
    qty: r.qty ? r.qty + (r.unit ? " " + r.unit : "") : "—",
    rate: r.rate != null ? money(r.rate) : "—",
    sum: money(r.amount),
  }));
  return `
  <div style="display:flex;justify-content:space-between;align-items:center">
    <div>
      <h1>Коммерческое предложение</h1>
      <div class="cp-meta">Стеклянные и алюминиевые конструкции · калькулятор v${appInfo.version}</div>
    </div>
    <div class="cp-meta">КП № ${new Date().toISOString().slice(0, 10)}-${String(Date.now()).slice(-4)}<br/>дата: ${fmtDate(new Date().toISOString())}</div>
  </div>
  <h3 style="margin-top:16px">Параметры</h3>
  <table>
    <tr><th>Тип конструкции</th><td>${esc(cons)}</td><th>Количество</th><td>${c.q} шт</td></tr>
    <tr><th>Размеры</th><td>${c.area ? c.area.toFixed(2) + " м² (ширина " + num(state.width) + " м × высота " + num(state.height) + " м)" : "—"}</td><th>Площадь</th><td>${areaStr} м²</td></tr>
    <tr><th>Клиент</th><td>${esc(state.client || "—")}</td><th>Телефон</th><td>${esc(state.phone || "—")}</td></tr>
  </table>
  <h3 style="margin-top:16px">Состав сметы</h3>
  <table>
    <tr><th style="width:52%">Позиция</th><th>Кол-во</th><th class="right">Цена</th><th class="right">Сумма</th></tr>
    ${rows.map((r) => `<tr><td>${esc(r.name)}</td><td>${esc(r.qty)}</td><td class="right">${esc(r.rate)}</td><td class="right">${esc(r.sum)}</td></tr>`).join("")}
    <tr><td colspan="3" class="right">Сумма до скидки</td><td class="right">${money(c.base)}</td></tr>
    ${c.disc > 0 ? `<tr><td colspan="3" class="right">Скидка (−${c.pct}%)</td><td class="right">− ${money(c.disc)}</td></tr>` : ""}
  </table>
  <div class="total-line">Итого к оплате: ${money(c.total)}</div>
  ${state.wishes ? `<div class="wishes"><b>Пожелания клиента:</b> ${esc(state.wishes)}</div>` : ""}
  <p style="color:#777;font-size:11px;margin-top:14px">Цены зафиксированы в момент расчёта и действуют до изменения прайс-листа. Расчёт выполнен в программе «Калькулятор конструкций».</p>
  <div class="sign">
    <div>Исполнитель: ____________________</div>
    <div>Клиент: ____________________</div>
  </div>
  `;
}

function printOffer() {
  const pr = $("print-area");
  pr.innerHTML = buildPrint();
  setTimeout(() => window.print(), 60);
}

/* ================= toast & misc ================= */

let toastTimer = null;
function toast(msg, isErr) {
  const t = $("toast");
  t.innerHTML = msg;
  t.classList.toggle("err", !!isErr);
  t.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove("show"), 5200);
}

/* ================= init ================= */

async function init() {
  setDbStatus("busy", "Загрузка базы…");
  try {
    appInfo = await window.api.appInfo();
    const res = await window.api.loadDb();
    DB = res.db;
    dbFrom = res.from;
    if (!DB.components || !DB.constructions) throw new Error("Повреждённая база");
    state = defaultStateFor(DB);
  } catch (e) {
    toast("Ошибка загрузки базы: " + esc(e.message), true);
  }
  renderDbStatus();
  renderAll();
  try {
    history = await window.api.loadHistory();
  } catch (_) {}
  renderHistory();

  const dims = () => {
    state.width = $("inWidth").value;
    state.height = $("inHeight").value;
    state.qty = $("inQty").value;
    const w = num(state.width);
    const h = num(state.height);
    const q = Math.max(1, Math.round(num(state.qty)));
    $("areaOne").textContent = (w * h).toFixed(2);
    $("areaTotal").textContent = (w * h * q).toFixed(2);
    renderGroupsAndSummary();
  };
  ["inWidth", "inHeight", "inQty"].forEach((id) => $(id).addEventListener("input", dims));

  $("inClient").addEventListener("input", (e) => (state.client = e.target.value));
  $("inPhone").addEventListener("input", (e) => {
    state.phone = phoneMask(e.target.value);
    e.target.value = state.phone;
  });
  $("inPhone").value = "+7";
  state.phone = $("inPhone").value;
  $("inWishes").addEventListener("input", (e) => {
    state.wishes = e.target.value;
    renderComponents();
    renderWishTags();
  });
  $("inManualDiscount").addEventListener(
    "input",
    (e) => ((state.manualDiscount = num(e.target.value)), renderSummary())
  );
  $("inAutoDiscount").addEventListener("change", (e) => {
    state.autoDiscount = e.target.checked;
    renderSummary();
  });
  $("inRepeatClient").addEventListener("change", (e) => {
    state.repeatClient = e.target.checked;
    renderSummary();
  });

  $("btnRefresh").addEventListener("click", updateDatabase);
  $("btnSave").addEventListener("click", saveCurrent);
  $("btnPdf").addEventListener("click", printOffer);
  $("btnHistory").addEventListener("click", () => {
    const panel = document.querySelector(".panel.summary");
    if (panel) panel.scrollTo({ top: panel.scrollHeight, behavior: "smooth" });
  });

  /* отдельные скроллы панелей: крутим ту, над которой курсор; на границе — соседнюю */
  const leftPanel = document.querySelector(".panel.form");
  const rightPanel = document.querySelector(".panel.summary");
  if (leftPanel && rightPanel) {
    for (const [cur, other] of [[leftPanel, rightPanel], [rightPanel, leftPanel]]) {
      cur.addEventListener("wheel", (e) => {
        const d = e.deltaY;
        const atTop = cur.scrollTop <= 0;
        const atBottom = cur.scrollTop + cur.clientHeight >= cur.scrollHeight - 1;
        if ((d < 0 && atTop) || (d > 0 && atBottom)) {
          other.scrollTop += d;
          e.preventDefault();
        }
      }, { passive: false });
    }
  }

  dims();
}

document.addEventListener("DOMContentLoaded", init);