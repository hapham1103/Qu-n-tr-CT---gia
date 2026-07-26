import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  AlertTriangle, CheckCircle2, Plus, Trash2, LayoutGrid, ListChecks, Package, Box, History, Upload, BarChart3, ChevronDown, ChevronRight, Settings2, Tag, FileDown, Bot, Send, X,
} from "lucide-react";

const RED = "#E31E24";
const BLACK = "#1A1A1A";
const GREEN = "#0F6E56";
const AMBER = "#B8860B";
const STORAGE_KEY = "pricing-app-multibrand-v1";

// internalOnly = true  -> chỉ trừ vào "Chi phí KM nội bộ"/"Giá nội bộ", KHÔNG trừ vào "Giá NET NPP"
// forcedBase = 'giaVonChia07' -> nền tính = Giá vốn / 0.7 thay vì Giá bán
const defaultGtCostCatalog = [
  { id: "voucher", name: "Voucher du lịch", internalOnly: false },
  { id: "voucherQuyDoi", name: "CK Voucher du lịch quy đổi thành hàng", internalOnly: false, forcedBase: "giaVonChia07" },
  { id: "tichLuy", name: "Tích luỹ", internalOnly: false },
  { id: "ckThang", name: "Chiết khấu tháng", internalOnly: false },
  { id: "ckQuy", name: "Chiết khấu quý", internalOnly: false },
  { id: "ckNam", name: "Chiết khấu năm", internalOnly: false },
  { id: "ckGiaiDoan", name: "Chiết khấu giai đoạn", internalOnly: false },
  { id: "chiPhiDaiLy", name: "Chi phí cho đại lý", internalOnly: false },
  { id: "hnkh", name: "Hội nghị khách hàng (HNKH)", internalOnly: true },
  { id: "tienMat", name: "Tiền mặt", internalOnly: false },
  { id: "roadshow", name: "Chi phí roadshow", internalOnly: true },
  { id: "thanhToanNgay", name: "Chi phí thanh toán ngay", internalOnly: false },
  { id: "duLichNam", name: "Chi phí du lịch năm", internalOnly: true },
  { id: "tichLuyChienDich", name: "Chi phí tích luỹ chiến dịch", internalOnly: true },
  { id: "khacGT", name: "Chi phí khác", internalOnly: false },
];

// MT (kênh siêu thị) - rolls up into one aggregate "Chiết khấu Sellout".
// group:'so_support' items only count the MAX among them, not the sum.
const defaultMtCostCatalog = [
  { id: "sfm", name: "SFM (Sale for man)", group: "" },
  { id: "soXuyenThang", name: "Hỗ trợ SO xuyên tháng", group: "so_support" },
  { id: "soCuoiTuan", name: "Hỗ trợ SO cuối tuần", group: "so_support" },
  { id: "soEvent", name: "Hỗ trợ SO event", group: "so_support" },
  { id: "traGop", name: "Trả góp", group: "" },
  { id: "mkt", name: "MKT", group: "" },
  { id: "hopDong", name: "Chiết khấu hợp đồng", group: "" },
  { id: "luong", name: "Chi phí lương", group: "" },
  { id: "khacMT", name: "Chi phí khác", group: "" },
];

const defaultCustomFields = [
  { id: "cf_netdukien", name: "Net dự kiến" },
  { id: "cf_niemyet", name: "Giá niêm yết (GNY)" },
  { id: "cf_webmin", name: "Giá bán lẻ tối thiểu Web" },
];

const sampleModels = [
  { id: "m1", name: "BOR-IH2000", group: "Bếp từ đôi cao cấp", giaVon: 4200000, giaBan: 6500000, target: 15, minMarginTarget: 37, customValues: {} },
  { id: "m2", name: "BOR-IH2100", group: "Bếp từ đôi cao cấp", giaVon: 4500000, giaBan: 6900000, target: 15, minMarginTarget: 37, customValues: {} },
  { id: "m3", name: "BOR-IH900", group: "Bếp từ đơn phổ thông", giaVon: 1800000, giaBan: 2900000, target: 12, minMarginTarget: 30, customValues: {} },
  { id: "m4", name: "BOR-HD600", group: "Máy hút mùi 60cm", giaVon: 2200000, giaBan: 3600000, target: 12, minMarginTarget: 30, customValues: {} },
];

let idCounter = 1;
const newId = () => `id${idCounter++}_${Math.random().toString(36).slice(2, 7)}`;

function formatVND(n) { return Math.round(n || 0).toLocaleString("vi-VN") + "đ"; }
function formatDate(iso) { try { return new Date(iso).toLocaleString("vi-VN"); } catch { return iso; } }
function normalizeHeader(h) { return String(h || "").trim().toLowerCase(); }
function todayISO() { return new Date().toISOString().slice(0, 10); }
function addDaysISO(days) { const d = new Date(); d.setDate(d.getDate() + days); return d.toISOString().slice(0, 10); }
function normalizeDistName(raw) {
  const s = String(raw || "").trim();
  if (!s) return { key: "khac", display: "Chưa xác định" };
  if (/^kh[aá]ch h[aà]ng l[eẻ]/i.test(s)) return { key: "__retail__", display: "Bán lẻ / Online (gộp chung)" };
  return { key: s.toUpperCase().replace(/\s+/g, " "), display: s };
}
function mergeCostLine(existing, patch) {
  return { on: false, mode: "percent", value: 0, ...(existing || {}), ...patch };
}
function ensureCostsFor(entry, key, catalog) {
  const map = { ...(entry[key] || {}) };
  let changed = false;
  catalog.forEach((c) => { if (!map[c.id]) { map[c.id] = { on: false, mode: "percent", value: 0 }; changed = true; } });
  return changed ? { ...entry, [key]: map } : entry;
}
function makeEntry(modelId, channel = "GT") {
  return { modelId, channel, muaQty: 1, tangQty: 0, giftModelId: "", targetOverride: null, netMinOverride: null, costsGT: {}, costsMT: {} };
}
function makePackage(type) {
  return { id: newId(), name: type === "sales" ? "Gói doanh số mới" : "Gói máy mới", type, salesTarget: type === "sales" ? 100000000 : 0, entries: [] };
}
function makeProgram(seq) {
  return { id: newId(), seq: seq || 1, name: "Chương trình mới", startDate: todayISO(), endDate: addDaysISO(30), standaloneEntries: [], packages: [] };
}
function getProgramStatus(p) {
  if (!p.startDate || !p.endDate) return { label: "Chưa đặt thời gian", color: "#999" };
  const now = new Date(); const start = new Date(p.startDate); const end = new Date(p.endDate);
  end.setHours(23, 59, 59, 999);
  if (now < start) return { label: "Sắp chạy", color: AMBER };
  if (now > end) return { label: "Đã kết thúc", color: "#888" };
  return { label: "Đang chạy", color: GREEN };
}
function exportProgramExcel(program, calcEntry, brandName) {
  const rows = [];
  const addRows = (entries, pkgName) => {
    entries.forEach((entry) => {
      const r = calcEntry(entry);
      if (!r) return;
      rows.push({
        "Chương trình": `#${program.seq || ""} ${program.name}`,
        "Gói": pkgName || "(riêng lẻ)",
        "Model": r.model.name,
        "Kênh": r.channel,
        "Giá vốn": r.model.giaVon,
        "Giá bán (NPP)": r.model.giaBan,
        "Mua": r.mua,
        "Tặng": r.tang,
        "NET Đại lý": Math.round(r.netDaiLy),
        "Giá NET NPP": Math.round(r.netPhanTich),
        "Giá nội bộ": Math.round(r.giaNoiBo),
        "Chi phí KM nội bộ": Math.round(r.chiPhiKM),
        "% CP": Number(r.pctCP.toFixed(1)),
        "LN/máy": Math.round(r.perUnitProfit),
        "%LN": Number(r.marginPct.toFixed(1)),
        "%LN khoán": r.khoanPct,
        "Chênh lệch": Math.round(r.chenhLech),
        "Trạng thái": r.ok ? "Đạt" : "Dưới mục tiêu",
        "Tổng SL": r.tongSL,
        "Thành tiền": Math.round(r.thanhTien),
        "Tổng CP KM": Math.round(r.tongChiPhiKM),
        "Tổng LN thu về": Math.round(r.tongLoiNhuan),
      });
    });
  };
  addRows(program.standaloneEntries, null);
  program.packages.forEach((pkg) => addRows(pkg.entries, `${pkg.type === "sales" ? "Gói doanh số" : "Gói máy"}: ${pkg.name}`));

  const byModel = {};
  const byPkg = {};
  rows.forEach((r) => {
    byModel[r.Model] = byModel[r.Model] || { "Model": r.Model, "Tổng SL": 0, "Thành tiền": 0, "Tổng CP KM": 0, "Tổng LN thu về": 0 };
    byModel[r.Model]["Tổng SL"] += r["Tổng SL"]; byModel[r.Model]["Thành tiền"] += r["Thành tiền"];
    byModel[r.Model]["Tổng CP KM"] += r["Tổng CP KM"]; byModel[r.Model]["Tổng LN thu về"] += r["Tổng LN thu về"];
    byPkg[r.Gói] = byPkg[r.Gói] || { "Gói": r.Gói, "Tổng SL": 0, "Thành tiền": 0, "Tổng CP KM": 0, "Tổng LN thu về": 0 };
    byPkg[r.Gói]["Tổng SL"] += r["Tổng SL"]; byPkg[r.Gói]["Thành tiền"] += r["Thành tiền"];
    byPkg[r.Gói]["Tổng CP KM"] += r["Tổng CP KM"]; byPkg[r.Gói]["Tổng LN thu về"] += r["Tổng LN thu về"];
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Chi tiết");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(Object.values(byModel)), "Theo Model");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(Object.values(byPkg)), "Theo Gói");
  const safeName = `${brandName}_CT${program.seq || ""}_${program.name}`.replace(/[^\w\d]+/g, "_").slice(0, 80);
  XLSX.writeFile(wb, `${safeName}.xlsx`);
}
function makeEmptyBrandData(withSample) {
  const p = makeProgram();
  return {
    models: withSample ? sampleModels.map((m) => ({ ...m })) : [],
    gtCostCatalog: defaultGtCostCatalog.map((c) => ({ ...c })),
    mtCostCatalog: defaultMtCostCatalog.map((c) => ({ ...c })),
    customFields: defaultCustomFields.map((c) => ({ ...c })),
    programs: [p],
    priceHistory: [],
    actualRecords: [],
  };
}

// Controlled number input that keeps its own text while typing, so clearing
// the field or typing fresh digits never gets stuck showing a leading "0".
function NumInput({ value, onChange, style, allowEmpty, ...rest }) {
  const [raw, setRaw] = useState(value === null || value === undefined ? "" : String(value));
  useEffect(() => {
    const next = value === null || value === undefined ? "" : String(value);
    setRaw((prev) => (Number(prev) === Number(next) && prev !== "" ? prev : next));
  }, [value]);
  return (
    <input
      type="text"
      inputMode="decimal"
      style={style}
      value={raw}
      onChange={(e) => {
        const v = e.target.value;
        if (v === "" || v === "-" || /^-?\d*\.?\d*$/.test(v)) {
          setRaw(v);
          if (v === "" || v === "-" || v === ".") {
            if (allowEmpty) onChange(null);
            else onChange(0);
          } else {
            onChange(Number(v));
          }
        }
      }}
      onFocus={(e) => e.target.select()}
      onBlur={() => { if (raw === "" || raw === "-" || raw === ".") setRaw(allowEmpty ? "" : "0"); }}
      {...rest}
    />
  );
}

export default function App() {
  const [brands, setBrands] = useState([
    { id: "brand1", name: "Thương hiệu 1" },
    { id: "brand2", name: "Thương hiệu 2" },
  ]);
  const [activeBrandId, setActiveBrandId] = useState("brand1");
  const [dataByBrand, setDataByBrand] = useState({
    brand1: makeEmptyBrandData(true),
    brand2: makeEmptyBrandData(false),
  });
  const [activeProgramId, setActiveProgramId] = useState(dataByBrand.brand1.programs[0].id);
  const [tab, setTab] = useState("program");
  const [loaded, setLoaded] = useState(false);
  const [saveState, setSaveState] = useState("idle");
  const [importMsg, setImportMsg] = useState("");
  const [actualsImportMsg, setActualsImportMsg] = useState("");
  const [monthFilter, setMonthFilter] = useState("all");
  const [expandedProgram, setExpandedProgram] = useState(null);
  const [newGtCostName, setNewGtCostName] = useState("");
  const [newMtCostName, setNewMtCostName] = useState("");
  const [newFieldName, setNewFieldName] = useState("");
  const [newBrandName, setNewBrandName] = useState("");
  const [editingBrandId, setEditingBrandId] = useState(null);
  const [aiMessages, setAiMessages] = useState([]);
  const [aiInput, setAiInput] = useState("");
  const [aiLoading, setAiLoading] = useState(false);
  const [aiDraft, setAiDraft] = useState(null); // { entry, packageId }

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, false);
        if (res && !cancelled) {
          const parsed = JSON.parse(res.value);
          if (parsed.brands?.length) setBrands(parsed.brands);
          if (parsed.dataByBrand) setDataByBrand(parsed.dataByBrand);
          if (parsed.activeBrandId) setActiveBrandId(parsed.activeBrandId);
          const bId = parsed.activeBrandId || "brand1";
          const progs = parsed.dataByBrand?.[bId]?.programs;
          if (progs?.length) setActiveProgramId(progs[0].id);
        }
      } catch (e) { /* chưa có dữ liệu lưu trước đó */ } finally { if (!cancelled) setLoaded(true); }
    })();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!loaded) return;
    setSaveState("saving");
    const t = setTimeout(async () => {
      try {
        await window.storage.set(STORAGE_KEY, JSON.stringify({ brands, activeBrandId, dataByBrand }), false);
        setSaveState("saved");
      } catch (e) { console.error("Lỗi lưu dữ liệu:", e); setSaveState("idle"); }
    }, 700);
    return () => clearTimeout(t);
  }, [brands, activeBrandId, dataByBrand, loaded]);

  const bd = dataByBrand[activeBrandId] || makeEmptyBrandData(false);
  const { models, gtCostCatalog, mtCostCatalog, customFields, programs, priceHistory, actualRecords } = bd;
  const activeProgram = programs.find((p) => p.id === activeProgramId) || programs[0];

  function updateBrand(patch) {
    setDataByBrand((prev) => ({ ...prev, [activeBrandId]: { ...prev[activeBrandId], ...patch } }));
  }
  function switchBrand(id) {
    setActiveBrandId(id);
    const progs = dataByBrand[id]?.programs;
    if (progs?.length) setActiveProgramId(progs[0].id);
  }
  function addBrand() {
    const name = newBrandName.trim();
    if (!name) return;
    const id = newId();
    setBrands((prev) => [...prev, { id, name }]);
    setDataByBrand((prev) => ({ ...prev, [id]: makeEmptyBrandData(false) }));
    setNewBrandName("");
    switchBrand(id);
  }
  function renameBrand(id, name) { setBrands((prev) => prev.map((b) => (b.id === id ? { ...b, name } : b))); }
  function removeBrand(id) {
    if (brands.length <= 1) return;
    const rest = brands.filter((b) => b.id !== id);
    setBrands(rest);
    setDataByBrand((prev) => { const cp = { ...prev }; delete cp[id]; return cp; });
    if (activeBrandId === id) switchBrand(rest[0].id);
  }

  // ---------- Custom reference fields ----------
  function addCustomField() {
    const name = newFieldName.trim();
    if (!name) return;
    updateBrand({ customFields: [...customFields, { id: newId(), name }] });
    setNewFieldName("");
  }
  function updateCustomField(id, patch) { updateBrand({ customFields: customFields.map((f) => (f.id === id ? { ...f, ...patch } : f)) }); }
  function removeCustomField(id) { updateBrand({ customFields: customFields.filter((f) => f.id !== id) }); }

  // ---------- Cost catalog management ----------
  function addGtCost() {
    const name = newGtCostName.trim();
    if (!name) return;
    updateBrand({ gtCostCatalog: [...gtCostCatalog, { id: newId(), name, internalOnly: false }] });
    setNewGtCostName("");
  }
  function updateGtCost(id, patch) { updateBrand({ gtCostCatalog: gtCostCatalog.map((c) => (c.id === id ? { ...c, ...patch } : c)) }); }
  function removeGtCost(id) { updateBrand({ gtCostCatalog: gtCostCatalog.filter((c) => c.id !== id) }); }
  function addMtCost() {
    const name = newMtCostName.trim();
    if (!name) return;
    updateBrand({ mtCostCatalog: [...mtCostCatalog, { id: newId(), name, group: "" }] });
    setNewMtCostName("");
  }
  function updateMtCost(id, patch) { updateBrand({ mtCostCatalog: mtCostCatalog.map((c) => (c.id === id ? { ...c, ...patch } : c)) }); }
  function removeMtCost(id) { updateBrand({ mtCostCatalog: mtCostCatalog.filter((c) => c.id !== id) }); }

  // ---------- Program / package / entry management ----------
  function updateProgram(id, patch) { updateBrand({ programs: programs.map((p) => (p.id === id ? { ...p, ...patch } : p)) }); }
  function addProgram() { const p = makeProgram(programs.length + 1); updateBrand({ programs: [...programs, p] }); setActiveProgramId(p.id); }
  function removeProgram(id) {
    const rest = programs.filter((p) => p.id !== id);
    updateBrand({ programs: rest });
    if (activeProgramId === id && rest.length) setActiveProgramId(rest[0].id);
  }
  function addStandaloneModel(modelId) {
    if (activeProgram.standaloneEntries.some((e) => e.modelId === modelId)) return;
    updateProgram(activeProgram.id, { standaloneEntries: [...activeProgram.standaloneEntries, makeEntry(modelId)] });
  }
  function removeStandaloneEntry(modelId) {
    updateProgram(activeProgram.id, { standaloneEntries: activeProgram.standaloneEntries.filter((e) => e.modelId !== modelId) });
  }
  function updateStandaloneEntry(modelId, patch) {
    updateProgram(activeProgram.id, { standaloneEntries: activeProgram.standaloneEntries.map((e) => (e.modelId !== modelId ? e : { ...e, ...patch })) });
  }
  function updateStandaloneCost(modelId, catalogKey, costId, patch) {
    const entries = activeProgram.standaloneEntries.map((e) => {
      if (e.modelId !== modelId) return e;
      const map = { ...e[catalogKey], [costId]: mergeCostLine(e[catalogKey][costId], patch) };
      return { ...e, [catalogKey]: map };
    });
    updateProgram(activeProgram.id, { standaloneEntries: entries });
  }
  function updatePackage(pkgId, patch) {
    updateProgram(activeProgram.id, { packages: activeProgram.packages.map((pk) => (pk.id === pkgId ? { ...pk, ...patch } : pk)) });
  }
  function addPackage(type) { updateProgram(activeProgram.id, { packages: [...activeProgram.packages, makePackage(type)] }); }
  function removePackage(pkgId) { updateProgram(activeProgram.id, { packages: activeProgram.packages.filter((pk) => pk.id !== pkgId) }); }
  function addModelToPackage(pkgId, modelId) {
    const pkg = activeProgram.packages.find((pk) => pk.id === pkgId);
    if (pkg.entries.some((e) => e.modelId === modelId)) return;
    updatePackage(pkgId, { entries: [...pkg.entries, makeEntry(modelId)] });
  }
  function removePackageEntry(pkgId, modelId) {
    const pkg = activeProgram.packages.find((pk) => pk.id === pkgId);
    updatePackage(pkgId, { entries: pkg.entries.filter((e) => e.modelId !== modelId) });
  }
  function updatePackageEntry(pkgId, modelId, patch) {
    const pkg = activeProgram.packages.find((pk) => pk.id === pkgId);
    updatePackage(pkgId, { entries: pkg.entries.map((e) => (e.modelId !== modelId ? e : { ...e, ...patch })) });
  }
  function updatePackageCost(pkgId, modelId, catalogKey, costId, patch) {
    const pkg = activeProgram.packages.find((pk) => pk.id === pkgId);
    const entries = pkg.entries.map((e) => {
      if (e.modelId !== modelId) return e;
      const map = { ...e[catalogKey], [costId]: mergeCostLine(e[catalogKey][costId], patch) };
      return { ...e, [catalogKey]: map };
    });
    updatePackage(pkgId, { entries });
  }

  function updateModel(id, patch) {
    const current = models.find((m) => m.id === id);
    if (!current) return;
    const historyAdds = [];
    ["giaVon", "giaBan"].forEach((field) => {
      if (patch[field] !== undefined && Number(patch[field]) !== Number(current[field])) {
        historyAdds.push({ id: newId(), modelId: id, modelName: current.name, field, oldValue: current[field], newValue: patch[field], date: new Date().toISOString() });
      }
    });
    updateBrand({
      models: models.map((m) => (m.id === id ? { ...m, ...patch } : m)),
      priceHistory: historyAdds.length ? [...historyAdds, ...priceHistory] : priceHistory,
    });
  }
  function updateModelCustomValue(id, fieldId, value) {
    const current = models.find((m) => m.id === id);
    if (!current) return;
    updateBrand({ models: models.map((m) => (m.id === id ? { ...m, customValues: { ...m.customValues, [fieldId]: value } } : m)) });
  }
  function addModel() { updateBrand({ models: [...models, { id: newId(), name: "Model mới", group: "Chưa phân nhóm", giaVon: 0, giaBan: 0, target: 12, minMarginTarget: 30, customValues: {} }] }); }
  function removeModel(id) { updateBrand({ models: models.filter((m) => m.id !== id) }); }

  async function handleImportExcel(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportMsg("Đang đọc file...");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      let working = [...models];
      const historyAdds = [];
      let added = 0, updated = 0;
      rows.forEach((row) => {
        const keys = Object.keys(row);
        const findVal = (candidates) => { const key = keys.find((k) => candidates.includes(normalizeHeader(k))); return key !== undefined ? row[key] : undefined; };
        const name = findVal(["tên model", "ten model", "model", "name"]);
        if (!name || String(name).trim() === "") return;
        const group = findVal(["nhóm", "nhom", "group", "cấu hình", "cau hinh"]) || "Chưa phân nhóm";
        const giaVon = Number(findVal(["giá vốn", "gia von", "cost"])) || 0;
        const giaBan = Number(findVal(["giá bán", "gia ban", "price"])) || 0;
        const target = Number(findVal(["ln mục tiêu", "ln muc tieu", "target", "ln mục tiêu %", "ln muc tieu %", "%ln khoan"])) || 12;
        const idx = working.findIndex((m) => m.name.trim().toLowerCase() === String(name).trim().toLowerCase());
        if (idx >= 0) {
          const old = working[idx];
          if (Number(old.giaVon) !== giaVon) historyAdds.push({ id: newId(), modelId: old.id, modelName: old.name, field: "giaVon", oldValue: old.giaVon, newValue: giaVon, date: new Date().toISOString() });
          if (Number(old.giaBan) !== giaBan) historyAdds.push({ id: newId(), modelId: old.id, modelName: old.name, field: "giaBan", oldValue: old.giaBan, newValue: giaBan, date: new Date().toISOString() });
          working[idx] = { ...old, group, giaVon, giaBan, target };
          updated++;
        } else { working.push({ id: newId(), name: String(name), group, giaVon, giaBan, target, minMarginTarget: 30, customValues: {} }); added++; }
      });
      updateBrand({ models: working, priceHistory: historyAdds.length ? [...historyAdds, ...priceHistory] : priceHistory });
      setImportMsg(`Nhập xong: ${added} model mới, ${updated} model cập nhật giá.`);
    } catch (err) {
      console.error(err);
      setImportMsg("Không đọc được file. Kiểm tra cột: Tên model, Nhóm, Giá vốn, Giá bán, LN mục tiêu %.");
    } finally { e.target.value = ""; }
  }

  async function handleImportActuals(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setActualsImportMsg("Đang đọc file...");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
      const G = (row, name) => (row[name] !== undefined ? row[name] : "");
      const asDate = (v) => {
        if (v instanceof Date) return v.toISOString().slice(0, 10);
        if (!v) return null;
        const d = new Date(v);
        return isNaN(d) ? null : d.toISOString().slice(0, 10);
      };
      const parsed = rows.map((row) => {
        const dist = normalizeDistName(G(row, "Tên khách"));
        return {
          id: newId(), date: asDate(G(row, "Ngày ct")) || todayISO(), month: String(G(row, "Tháng") || ""),
          channel: String(G(row, "Tên Kênh") || ""), distKey: dist.key, distName: dist.display,
          modelCode: String(G(row, "Model") || "").trim(), itemName: String(G(row, "Tên mặt hàng") || "").trim(),
          group: String(G(row, "Tên nhóm") || G(row, "Tênh ngành") || "Chưa phân nhóm"),
          qty: Number(G(row, "Sl xuất bán")) || 0, unitPrice: Number(G(row, "Giá bán tb")) || 0,
          revenue: Number(G(row, "Doanh thu xuất bán")) || 0, netRevenue: Number(G(row, "Thực bán")) || 0,
          cost: Number(G(row, "Tiền vốn thực tế")) || 0, profit: Number(G(row, "Lãi")) || 0,
          ctkmCode: String(G(row, "Mã CTKM") || "").trim(), ctkmName: String(G(row, "Tên ct km") || "").trim(),
          promoStart: asDate(G(row, "Ngày bắt đầu")), promoEnd: asDate(G(row, "Ngày kết thúc")), nvbh: String(G(row, "Mã NVBH") || ""),
        };
      }).filter((r) => r.modelCode || r.itemName);
      updateBrand({ actualRecords: [...parsed, ...actualRecords] });
      const months = [...new Set(parsed.map((r) => r.month).filter(Boolean))];
      setActualsImportMsg(`Nhập xong ${parsed.length} dòng (${months.join(", ") || "không rõ tháng"}).`);
    } catch (err) {
      console.error(err);
      setActualsImportMsg("Không đọc được file. Cần đúng định dạng báo cáo tháng ERP.");
    } finally { e.target.value = ""; }
  }

  // ---------- Core calculation engine ----------
  function calcEntry(entry) {
    const model = models.find((m) => m.id === entry.modelId);
    if (!model) return null;

    if (entry.channel === "MT") {
      const qty = Number(entry.muaQty) || 0;
      const sub = {};
      mtCostCatalog.forEach((c) => {
        const line = entry.costsMT[c.id];
        if (line && line.on) {
          const amt = line.mode === "percent" ? (line.value / 100) * model.giaBan : line.value;
          sub[c.id] = amt;
        }
      });
      const soIds = mtCostCatalog.filter((c) => c.group === "so_support").map((c) => c.id);
      let so = 0;
      soIds.forEach((id) => { if (sub[id] !== undefined) so = Math.max(so, sub[id]); });
      const otherIds = mtCostCatalog.filter((c) => c.group !== "so_support").map((c) => c.id);
      let sellout = so;
      otherIds.forEach((id) => { if (sub[id] !== undefined) sellout += sub[id]; });
      const perCostDisplay = Object.keys(sub).length > 0 ? { sellout_mt: { name: "Chiết khấu Sellout (tổng)", amount: sellout } } : {};

      const perUnitProfit = model.giaBan - model.giaVon - sellout;
      const marginPct = model.giaBan > 0 ? (perUnitProfit / model.giaBan) * 100 : 0;
      const giaNetUnit = model.giaBan - sellout;
      const pctCP = model.giaBan > 0 ? (sellout / model.giaBan) * 100 : 0;
      const khoanPct = (entry.targetOverride !== null && entry.targetOverride !== undefined && entry.targetOverride !== "") ? Number(entry.targetOverride) : model.target;
      const ok = marginPct >= khoanPct;
      const lnGopPct = model.giaBan > 0 ? ((model.giaBan - model.giaVon) / model.giaBan) * 100 : 0;
      const netMinPct = (entry.netMinOverride !== null && entry.netMinOverride !== undefined && entry.netMinOverride !== "") ? Number(entry.netMinOverride) : (model.minMarginTarget ?? 30);
      const netMinFrac = Math.min(Math.max(netMinPct, 0), 99) / 100;
      const netToiThieu = netMinFrac < 1 ? model.giaVon / (1 - netMinFrac) : model.giaVon;
      const lnThucToiThieu = netToiThieu - model.giaVon;
      const chenhLech = perUnitProfit - lnThucToiThieu;
      return {
        model, channel: "MT", perCostDisplay, chiPhiKM: sellout, netDaiLy: model.giaBan, netPhanTich: giaNetUnit, giaNoiBo: giaNetUnit, pctCP,
        perUnitProfit, marginPct, ok, khoanPct, netMinPct, mua: qty, tang: 0, tongSL: qty, lnGopPct, netToiThieu, lnThucToiThieu, chenhLech,
        thanhTien: model.giaBan * qty, tongChiPhiKM: sellout * qty, tongLoiNhuan: perUnitProfit * qty,
      };
    }

    // ---- GT / Delta ----
    const giftModel = entry.giftModelId ? models.find((m) => m.id === entry.giftModelId) : model;
    const mua = Number(entry.muaQty) || 0;
    const tang = Number(entry.tangQty) || 0;
    const giftGiaBan = giftModel ? giftModel.giaBan : model.giaBan;
    const giftGiaVon = giftModel ? giftModel.giaVon : model.giaVon;
    // NET Đại lý = Giá bán − (Tặng × Giá bán model tặng) / (Mua + Tặng)
    const netDaiLy = (mua + tang) > 0 ? model.giaBan - (tang * giftGiaBan) / (mua + tang) : model.giaBan;
    const giftCostPerSold = mua > 0 ? (tang * giftGiaVon) / mua : 0;

    let sumBoth = 0; // affects NPP-facing Giá NET
    let sumAll = 0; // all % costs (both + internalOnly), goes into Chi phí KM nội bộ
    const perCostDisplay = {};
    const entryCosts = entry.costsGT || {};
    gtCostCatalog.forEach((c) => {
      const line = entryCosts[c.id];
      if (line && line.on) {
        let amt;
        if (c.forcedBase === "giaVonChia07") {
          const base = model.giaVon / 0.7;
          amt = line.mode === "percent" ? (line.value / 100) * base : line.value;
        } else {
          amt = line.mode === "percent" ? (line.value / 100) * model.giaBan : line.value;
        }
        sumAll += amt;
        if (!c.internalOnly) sumBoth += amt;
        perCostDisplay[c.id] = { name: c.name, amount: amt, internalOnly: !!c.internalOnly };
      }
    });

    const netPhanTich = netDaiLy - sumBoth; // Giá NET NPP
    const chiPhiKMNoiBo = giftCostPerSold + sumAll; // Chi phí KM nội bộ theo giá vốn
    const giaNoiBo = model.giaBan - chiPhiKMNoiBo; // Giá nội bộ
    const pctCP = model.giaBan > 0 ? (chiPhiKMNoiBo / model.giaBan) * 100 : 0; // % CP
    const perUnitProfit = model.giaBan - model.giaVon - chiPhiKMNoiBo; // giá trị LN thu về / máy bán
    const marginPct = model.giaBan > 0 ? (perUnitProfit / model.giaBan) * 100 : 0; // LN thu về %
    const khoanPct = (entry.targetOverride !== null && entry.targetOverride !== undefined && entry.targetOverride !== "") ? Number(entry.targetOverride) : model.target;
    const ok = marginPct >= khoanPct;
    const sameModelGift = !entry.giftModelId || entry.giftModelId === entry.modelId;
    const tongSL = sameModelGift ? mua + tang : mua;
    const lnGopPct = model.giaBan > 0 ? ((model.giaBan - model.giaVon) / model.giaBan) * 100 : 0;
    const netMinPct = (entry.netMinOverride !== null && entry.netMinOverride !== undefined && entry.netMinOverride !== "") ? Number(entry.netMinOverride) : (model.minMarginTarget ?? 30);
    const netMinFrac = Math.min(Math.max(netMinPct, 0), 99) / 100;
    const netToiThieu = netMinFrac < 1 ? model.giaVon / (1 - netMinFrac) : model.giaVon;
    const lnThucToiThieu = netToiThieu - model.giaVon;
    const chenhLech = perUnitProfit - lnThucToiThieu;

    return {
      model, channel: "GT", mua, tang, giftModel, perCostDisplay,
      netDaiLy, netPhanTich, chiPhiKM: chiPhiKMNoiBo, giaNoiBo, pctCP, lnGopPct, netToiThieu, lnThucToiThieu, khoanPct, netMinPct, chenhLech,
      perUnitProfit, marginPct, ok, tongSL,
      thanhTien: model.giaBan * mua, tongChiPhiKM: chiPhiKMNoiBo * mua, tongLoiNhuan: perUnitProfit * mua,
    };
  }

  // ---------- AI assistant ----------
  function buildAiSystemPrompt() {
    const modelList = models.map((m) => `- id="${m.id}" tên="${m.name}" nhóm="${m.group}" giá vốn=${m.giaVon} giá bán=${m.giaBan} %LN khoán mặc định=${m.target} % tỷ lệ tối thiểu mặc định=${m.minMarginTarget ?? 30}`).join("\n");
    const gtList = gtCostCatalog.map((c) => `- id="${c.id}" tên="${c.name}"${c.internalOnly ? " (chỉ tính nội bộ)" : ""}`).join("\n");
    const mtList = mtCostCatalog.map((c) => `- id="${c.id}" tên="${c.name}"${c.group === "so_support" ? " (nhóm Hỗ trợ SO)" : ""}`).join("\n");
    const pkgList = activeProgram.packages.map((p) => `- id="${p.id}" tên="${p.name}" loại="${p.type === "sales" ? "Gói doanh số" : "Gói máy"}"`).join("\n") || "(chương trình hiện chưa có gói nào)";
    return `Bạn là trợ lý phân tích giá & khuyến mại cho chương trình "#${activeProgram.seq} ${activeProgram.name}".

DANH MỤC MODEL:
${modelList}

DANH MỤC CHI PHÍ KÊNH GT/DELTA:
${gtList}

DANH MỤC CHI PHÍ KÊNH MT (SIÊU THỊ) — gộp chung thành 1 khoản Chiết khấu Sellout:
${mtList}

CÁC GÓI TRONG CHƯƠNG TRÌNH ĐANG CHỌN:
${pkgList}

Phân loại yêu cầu người dùng thành đúng 1 trong 3 loại, CHỈ trả về JSON hợp lệ, không kèm chữ nào khác, không dùng markdown code fence:

1) Nếu người dùng muốn TÍNH/THÊM PHÂN TÍCH cho 1 model (có thể kèm gói, số lượng mua/tặng, các chi phí):
{"type":"action","modelId":"<id đúng từ danh mục trên>","channel":"GT hoặc MT","muaQty":<số>,"tangQty":<số, mặc định 0>,"giftModelId":"<id, để null nếu tặng cùng loại>","packageId":"<id gói nếu người dùng nhắc tên gói khớp, để null nếu để riêng lẻ>","targetOverride":<số hoặc null>,"netMinOverride":<số hoặc null>,"costs":[{"costId":"<id đúng từ danh mục chi phí>","mode":"percent hoặc fixed","value":<số>}]}

2) Nếu người dùng HỎI thông tin (giá, chi phí, model nào đang có...), KHÔNG yêu cầu tính toán mới:
{"type":"answer","text":"<câu trả lời ngắn gọn dựa trên dữ liệu trên>"}

3) Nếu người dùng yêu cầu XUẤT BÁO CÁO/FILE EXCEL:
{"type":"report","text":"<xác nhận ngắn gọn>"}

Nếu không chắc model/chi phí nào người dùng nhắc tới, chọn model/chi phí có tên gần đúng nhất trong danh mục. Không tự bịa id không có trong danh mục.`;
  }

  async function sendAiMessage() {
    const msg = aiInput.trim();
    if (!msg || aiLoading) return;
    setAiMessages((prev) => [...prev, { role: "user", content: msg }]);
    setAiInput("");
    setAiLoading(true);
    try {
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1024,
          system: buildAiSystemPrompt(),
          messages: [{ role: "user", content: msg }],
        }),
      });
      const data = await response.json();
      const text = (data.content || []).map((b) => b.text || "").join("");
      const clean = text.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      handleAiResponse(parsed);
    } catch (err) {
      console.error(err);
      setAiMessages((prev) => [...prev, { role: "assistant", content: "Mình chưa xử lý được yêu cầu này. Bạn thử diễn đạt rõ hơn (tên model, gói, chi phí) nhé." }]);
    } finally {
      setAiLoading(false);
    }
  }

  function handleAiResponse(parsed) {
    if (parsed.type === "answer") {
      setAiMessages((prev) => [...prev, { role: "assistant", content: parsed.text }]);
    } else if (parsed.type === "report") {
      exportProgramExcel(activeProgram, calcEntry, brands.find((b) => b.id === activeBrandId)?.name || "");
      setAiMessages((prev) => [...prev, { role: "assistant", content: parsed.text || "Đã xuất file Excel báo cáo cho chương trình đang chọn." }]);
    } else if (parsed.type === "action") {
      const model = models.find((m) => m.id === parsed.modelId);
      if (!model) {
        setAiMessages((prev) => [...prev, { role: "assistant", content: "Mình không tìm thấy đúng model bạn nhắc tới trong danh mục. Bạn kiểm tra lại tên model nhé." }]);
        return;
      }
      const channel = parsed.channel === "MT" ? "MT" : "GT";
      const entry = makeEntry(model.id, channel);
      entry.muaQty = Number(parsed.muaQty) || 1;
      entry.tangQty = Number(parsed.tangQty) || 0;
      if (parsed.giftModelId && models.find((m) => m.id === parsed.giftModelId)) entry.giftModelId = parsed.giftModelId;
      entry.targetOverride = parsed.targetOverride ?? null;
      entry.netMinOverride = parsed.netMinOverride ?? null;
      const catalog = channel === "MT" ? mtCostCatalog : gtCostCatalog;
      const key = channel === "MT" ? "costsMT" : "costsGT";
      catalog.forEach((c) => { entry[key][c.id] = { on: false, mode: "percent", value: 0 }; });
      (parsed.costs || []).forEach((c) => {
        const found = catalog.find((cc) => cc.id === c.costId);
        if (found) entry[key][found.id] = { on: true, mode: c.mode === "fixed" ? "fixed" : "percent", value: Number(c.value) || 0 };
      });
      const pkg = parsed.packageId ? activeProgram.packages.find((p) => p.id === parsed.packageId) : null;
      setAiDraft({ entry, packageId: pkg ? pkg.id : null });
      setAiMessages((prev) => [...prev, { role: "assistant", content: "Đã tính xong, xem kết quả bên dưới:", isDraft: true }]);
    } else {
      setAiMessages((prev) => [...prev, { role: "assistant", content: "Mình chưa hiểu rõ yêu cầu, bạn nói lại giúp mình được không?" }]);
    }
  }

  function saveAiDraft() {
    if (!aiDraft) return;
    if (aiDraft.packageId) {
      const pkg = activeProgram.packages.find((p) => p.id === aiDraft.packageId);
      if (pkg) updatePackage(pkg.id, { entries: [...pkg.entries.filter((e) => e.modelId !== aiDraft.entry.modelId), aiDraft.entry] });
    } else {
      updateProgram(activeProgram.id, { standaloneEntries: [...activeProgram.standaloneEntries.filter((e) => e.modelId !== aiDraft.entry.modelId), aiDraft.entry] });
    }
    setAiMessages((prev) => [...prev, { role: "assistant", content: `Đã lưu vào chương trình #${activeProgram.seq} ${activeProgram.name}.` }]);
    setAiDraft(null);
  }
  function discardAiDraft() {
    setAiDraft(null);
    setAiMessages((prev) => [...prev, { role: "assistant", content: "Đã xoá nháp, không lưu vào chương trình." }]);
  }

  // ---------- Actuals aggregation ----------
  const availableMonths = useMemo(() => [...new Set(actualRecords.map((r) => r.month).filter(Boolean))].sort(), [actualRecords]);
  const filteredRecords = useMemo(() => (monthFilter === "all" ? actualRecords : actualRecords.filter((r) => r.month === monthFilter)), [actualRecords, monthFilter]);
  const aggByDist = useMemo(() => {
    const g = {};
    filteredRecords.forEach((r) => { const key = r.distName; if (!g[key]) g[key] = { qty: 0, revenue: 0, cost: 0, profit: 0 }; g[key].qty += r.qty; g[key].revenue += r.revenue; g[key].cost += r.cost; g[key].profit += r.profit; });
    return g;
  }, [filteredRecords]);
  const aggByModel = useMemo(() => {
    const g = {};
    filteredRecords.forEach((r) => { const key = r.modelCode || r.itemName; if (!g[key]) g[key] = { qty: 0, revenue: 0, cost: 0, profit: 0 }; g[key].qty += r.qty; g[key].revenue += r.revenue; g[key].cost += r.cost; g[key].profit += r.profit; });
    return g;
  }, [filteredRecords]);
  const aggByProgram = useMemo(() => {
    const g = {};
    filteredRecords.forEach((r) => {
      const key = r.ctkmCode || (r.ctkmName ? r.ctkmName : "__none__");
      if (!g[key]) g[key] = { name: r.ctkmName || "Không có CTKM", qty: 0, revenue: 0, cost: 0, profit: 0, byDist: {}, models: new Set() };
      g[key].qty += r.qty; g[key].revenue += r.revenue; g[key].cost += r.cost; g[key].profit += r.profit;
      g[key].models.add(r.modelCode || r.itemName);
      const dk = r.distName;
      if (!g[key].byDist[dk]) g[key].byDist[dk] = { qty: 0, revenue: 0, cost: 0, profit: 0 };
      g[key].byDist[dk].qty += r.qty; g[key].byDist[dk].revenue += r.revenue; g[key].byDist[dk].cost += r.cost; g[key].byDist[dk].profit += r.profit;
    });
    return g;
  }, [filteredRecords]);
  const monthlyTrend = useMemo(() => {
    const g = {};
    actualRecords.forEach((r) => { const key = r.month || "?"; if (!g[key]) g[key] = { qty: 0, revenue: 0, cost: 0, profit: 0 }; g[key].qty += r.qty; g[key].revenue += r.revenue; g[key].cost += r.cost; g[key].profit += r.profit; });
    return Object.entries(g).sort((a, b) => a[0].localeCompare(b[0], "vi", { numeric: true }));
  }, [actualRecords]);

  return (
    <div style={{ fontFamily: "system-ui, sans-serif", color: BLACK, maxWidth: 1360, margin: "0 auto", padding: 20 }}>
      <header style={{ borderBottom: `3px solid ${RED}`, paddingBottom: 12, marginBottom: 14, display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 8 }}>
        <div>
          <div style={{ fontSize: 12, letterSpacing: 1, color: RED, fontWeight: 700 }}>QUẢN LÝ GIÁ & KHUYẾN MẠI</div>
          <h1 style={{ fontSize: 20, fontWeight: 700, margin: "2px 0 0" }}>Giá, doanh số thực tế & chương trình khuyến mại</h1>
        </div>
        <div style={{ fontSize: 11, color: "#999" }}>{saveState === "saving" ? "Đang lưu..." : loaded ? "Đã lưu tự động" : "Đang tải dữ liệu..."}</div>
      </header>

      {/* Brand switcher — dữ liệu hoàn toàn tách biệt theo từng thương hiệu */}
      <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 18, flexWrap: "wrap", background: "#fafafa", border: "1px solid #eee", borderRadius: 8, padding: 10 }}>
        <Tag size={15} color={RED} />
        {brands.map((b) => (
          <div key={b.id} style={{ display: "flex", alignItems: "center", gap: 4 }}>
            {editingBrandId === b.id ? (
              <input
                style={{ ...inputXs, width: 130 }} autoFocus value={b.name}
                onChange={(e) => renameBrand(b.id, e.target.value)}
                onBlur={() => setEditingBrandId(null)}
                onKeyDown={(e) => e.key === "Enter" && setEditingBrandId(null)}
              />
            ) : (
              <button
                onClick={() => switchBrand(b.id)} onDoubleClick={() => setEditingBrandId(b.id)}
                title="Bấm để chọn, bấm đúp để đổi tên"
                style={{
                  padding: "6px 14px", borderRadius: 6, border: `1px solid ${b.id === activeBrandId ? RED : "#ddd"}`,
                  background: b.id === activeBrandId ? RED : "#fff", color: b.id === activeBrandId ? "#fff" : BLACK,
                  fontSize: 13, fontWeight: 600, cursor: "pointer",
                }}
              >
                {b.name}
              </button>
            )}
            {brands.length > 1 && b.id === activeBrandId && (
              <button onClick={() => removeBrand(b.id)} style={iconBtn} title="Xoá thương hiệu này"><Trash2 size={12} /></button>
            )}
          </div>
        ))}
        <input style={{ ...inputSm, maxWidth: 160 }} placeholder="Tên thương hiệu mới..." value={newBrandName} onChange={(e) => setNewBrandName(e.target.value)} />
        <button onClick={addBrand} style={addBtn}><Plus size={14} /> Thêm thương hiệu</button>
      </div>

      <nav style={{ display: "flex", gap: 8, marginBottom: 20, flexWrap: "wrap" }}>
        <TabButton active={tab === "ai"} onClick={() => setTab("ai")} icon={<Bot size={16} />} label="Trợ lý AI" />
        <TabButton active={tab === "program"} onClick={() => setTab("program")} icon={<ListChecks size={16} />} label="Kế hoạch chương trình" />
        <TabButton active={tab === "actuals"} onClick={() => setTab("actuals")} icon={<BarChart3 size={16} />} label={`Dữ liệu thực tế (${actualRecords.length})`} />
        <TabButton active={tab === "danhmuc"} onClick={() => setTab("danhmuc")} icon={<LayoutGrid size={16} />} label="Danh mục model" />
        <TabButton active={tab === "chiphi"} onClick={() => setTab("chiphi")} icon={<Settings2 size={16} />} label="Danh mục chi phí" />
        <TabButton active={tab === "history"} onClick={() => setTab("history")} icon={<History size={16} />} label={`Lịch sử giá (${priceHistory.length})`} />
      </nav>

      {tab === "ai" && (
        <section>
          <div style={{ fontSize: 12.5, color: "#666", marginBottom: 14, background: "#fafafa", border: "1px solid #eee", borderRadius: 8, padding: 12 }}>
            Gõ tự nhiên, ví dụ: <em>"Model KG35A2, kênh GT, mua 2 tặng 1, tích luỹ 5%, HNKH 10%"</em>. Bạn có thể hỏi thông tin (giá, chi phí...) hoặc yêu cầu xuất báo cáo Excel. Trợ lý làm việc trên chương trình đang chọn: <strong>#{activeProgram.seq} {activeProgram.name}</strong>.
          </div>

          <div style={{ border: "1px solid #eee", borderRadius: 10, padding: 14, minHeight: 200, marginBottom: 14, maxHeight: 420, overflowY: "auto" }}>
            {aiMessages.length === 0 && <div style={{ fontSize: 13, color: "#aaa" }}>Chưa có hội thoại nào. Gõ yêu cầu bên dưới để bắt đầu.</div>}
            {aiMessages.map((m, i) => (
              <div key={i} style={{ display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start", marginBottom: 10 }}>
                <div style={{
                  maxWidth: "80%", padding: "8px 12px", borderRadius: 8, fontSize: 13,
                  background: m.role === "user" ? RED : "#f2f2f2", color: m.role === "user" ? "#fff" : BLACK,
                }}>
                  {m.content}
                </div>
              </div>
            ))}
            {aiLoading && <div style={{ fontSize: 12.5, color: "#999" }}>Đang xử lý...</div>}

            {aiDraft && (() => {
              const r = calcEntry(aiDraft.entry);
              if (!r) return null;
              const pkgName = aiDraft.packageId ? activeProgram.packages.find((p) => p.id === aiDraft.packageId)?.name : null;
              return (
                <div style={{ border: `1px solid ${r.ok ? GREEN : RED}`, borderRadius: 8, padding: 12, marginTop: 6, background: r.ok ? "#F0FAF7" : "#FCEBEC" }}>
                  <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 6 }}>
                    {r.model.name} — {r.channel === "MT" ? "kênh MT" : "kênh GT/Delta"} {pkgName ? `— ${pkgName}` : "— model riêng lẻ"}
                  </div>
                  <div style={{ display: "flex", gap: 16, flexWrap: "wrap", fontSize: 12.5, marginBottom: 8 }}>
                    <span>Giá NET NPP: <strong>{formatVND(r.netPhanTich)}</strong></span>
                    <span>LN/máy: <strong>{formatVND(r.perUnitProfit)}</strong></span>
                    <span>%LN: <strong style={{ color: r.ok ? GREEN : RED }}>{r.marginPct.toFixed(1)}%</strong></span>
                    <span>Tổng LN thu về: <strong>{formatVND(r.tongLoiNhuan)}</strong></span>
                    <span style={{ color: r.ok ? GREEN : RED, fontWeight: 700 }}>{r.ok ? "Đạt mục tiêu" : "Dưới mục tiêu"}</span>
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button onClick={saveAiDraft} style={{ ...addBtn, borderColor: GREEN, color: GREEN }}><CheckCircle2 size={14} /> Lưu vào chương trình đang chạy</button>
                    <button onClick={discardAiDraft} style={{ ...addBtn }}><X size={14} /> Xoá nháp</button>
                  </div>
                </div>
              );
            })()}
          </div>

          <div style={{ display: "flex", gap: 8 }}>
            <input
              style={{ ...inputSm }} placeholder="Gõ yêu cầu của bạn..." value={aiInput}
              onChange={(e) => setAiInput(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && sendAiMessage()}
            />
            <button onClick={sendAiMessage} disabled={aiLoading} style={addBtn}><Send size={14} /> Gửi</button>
          </div>
        </section>
      )}

      {tab === "chiphi" && (
        <section>
          <SectionTitle>Kênh GT / Delta</SectionTitle>
          <div style={{ fontSize: 12.5, color: "#666", marginBottom: 12, background: "#fafafa", border: "1px solid #eee", borderRadius: 8, padding: 12 }}>
            Tính theo <strong>%</strong> (trên giá bán/giá xuất hoá đơn) hoặc <strong>giá trị cố định</strong>. Tick <strong>"Chỉ tính nội bộ"</strong> cho chi phí công ty tự chịu, KHÔNG trừ vào Giá NET của nhà phân phối (giống HNKH).
          </div>
          {gtCostCatalog.map((c) => (
            <div key={c.id} style={rowCard}>
              <input style={inputSm} value={c.name} onChange={(e) => updateGtCost(c.id, { name: e.target.value })} />
              <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, whiteSpace: "nowrap" }}>
                <input type="checkbox" checked={!!c.internalOnly} onChange={(e) => updateGtCost(c.id, { internalOnly: e.target.checked })} />
                Chỉ tính nội bộ (không trừ Giá NET NPP)
              </label>
              {c.forcedBase === "giaVonChia07" && <span style={{ fontSize: 11, color: AMBER }}>Nền: Giá vốn ÷ 0.7</span>}
              <button onClick={() => removeGtCost(c.id)} style={iconBtn}><Trash2 size={14} /></button>
            </div>
          ))}
          <div style={{ display: "flex", gap: 8, marginTop: 10, marginBottom: 28 }}>
            <input style={{ ...inputSm, maxWidth: 280 }} placeholder="Tên chi phí mới..." value={newGtCostName} onChange={(e) => setNewGtCostName(e.target.value)} />
            <button onClick={addGtCost} style={addBtn}><Plus size={14} /> Thêm chi phí</button>
          </div>

          <SectionTitle>Kênh MT (Siêu thị)</SectionTitle>
          <div style={{ fontSize: 12.5, color: "#666", marginBottom: 12, background: "#fafafa", border: "1px solid #eee", borderRadius: 8, padding: 12 }}>
            Tất cả các khoản dưới đây gộp thành 1 con số <strong>"Chiết khấu Sellout"</strong>. Tick <strong>"Thuộc nhóm Hỗ trợ SO"</strong> cho các khoản chỉ nên tính GIÁ TRỊ LỚN NHẤT (không cộng dồn).
          </div>
          {mtCostCatalog.map((c) => (
            <div key={c.id} style={rowCard}>
              <input style={inputSm} value={c.name} onChange={(e) => updateMtCost(c.id, { name: e.target.value })} />
              <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, whiteSpace: "nowrap" }}>
                <input type="checkbox" checked={c.group === "so_support"} onChange={(e) => updateMtCost(c.id, { group: e.target.checked ? "so_support" : "" })} />
                Thuộc nhóm Hỗ trợ SO (lấy giá trị lớn nhất)
              </label>
              <button onClick={() => removeMtCost(c.id)} style={iconBtn}><Trash2 size={14} /></button>
            </div>
          ))}
          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <input style={{ ...inputSm, maxWidth: 280 }} placeholder="Tên chi phí mới..." value={newMtCostName} onChange={(e) => setNewMtCostName(e.target.value)} />
            <button onClick={addMtCost} style={addBtn}><Plus size={14} /> Thêm chi phí</button>
          </div>
        </section>
      )}

      {tab === "danhmuc" && (
        <section>
          <div style={{ border: "1px solid #eee", borderRadius: 10, padding: 14, marginBottom: 18, background: "#fafafa" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}><Upload size={15} color={RED} /><span style={{ fontSize: 13, fontWeight: 700 }}>Nhập danh mục model từ Excel</span></div>
            <div style={{ fontSize: 12, color: "#666", marginBottom: 8 }}>Cột cần có: <strong>Tên model, Nhóm, Giá vốn, Giá bán, LN mục tiêu %</strong>.</div>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleImportExcel} style={{ fontSize: 12 }} />
            {importMsg && <div style={{ fontSize: 12, marginTop: 8, color: RED }}>{importMsg}</div>}
          </div>

          <SectionTitle>Cột tham chiếu tuỳ chỉnh</SectionTitle>
          <div style={{ fontSize: 12.5, color: "#666", marginBottom: 10 }}>Ví dụ: Net dự kiến, Giá niêm yết, Giá bán lẻ tối thiểu Web... Thêm/xoá tuỳ ý, chỉ để theo dõi tham khảo, không ảnh hưởng công thức tính.</div>
          {customFields.map((f) => (
            <div key={f.id} style={rowCard}>
              <input style={inputSm} value={f.name} onChange={(e) => updateCustomField(f.id, { name: e.target.value })} />
              <button onClick={() => removeCustomField(f.id)} style={iconBtn}><Trash2 size={14} /></button>
            </div>
          ))}
          <div style={{ display: "flex", gap: 8, marginTop: 10, marginBottom: 24 }}>
            <input style={{ ...inputSm, maxWidth: 280 }} placeholder="Tên cột tham chiếu mới..." value={newFieldName} onChange={(e) => setNewFieldName(e.target.value)} />
            <button onClick={addCustomField} style={addBtn}><Plus size={14} /> Thêm cột</button>
          </div>

          <SectionTitle>Danh mục model</SectionTitle>
          {models.map((m) => (
            <div key={m.id} style={rowCard}>
              <input style={inputSm} value={m.name} onChange={(e) => updateModel(m.id, { name: e.target.value })} />
              <input style={inputSm} value={m.group} onChange={(e) => updateModel(m.id, { group: e.target.value })} placeholder="Nhóm cấu hình" />
              <LabeledNum label="Giá vốn" value={m.giaVon} onChange={(v) => updateModel(m.id, { giaVon: v })} />
              <LabeledNum label="Giá bán (giá NPP)" value={m.giaBan} onChange={(v) => updateModel(m.id, { giaBan: v })} />
              <LabeledNum label="%LN khoán mặc định" value={m.target} onChange={(v) => updateModel(m.id, { target: v })} small />
              <LabeledNum label="% tỷ lệ tối thiểu mặc định" value={m.minMarginTarget ?? 30} onChange={(v) => updateModel(m.id, { minMarginTarget: v })} small />
              {customFields.map((f) => (
                <LabeledNum key={f.id} label={f.name} value={m.customValues?.[f.id] ?? 0} onChange={(v) => updateModelCustomValue(m.id, f.id, v)} />
              ))}
              <button onClick={() => removeModel(m.id)} style={iconBtn}><Trash2 size={14} /></button>
            </div>
          ))}
          <button onClick={addModel} style={addBtn}><Plus size={14} /> Thêm model</button>
        </section>
      )}

      {tab === "history" && (
        <section>
          <SectionTitle>Lịch sử thay đổi giá</SectionTitle>
          {priceHistory.length === 0 ? <div style={{ fontSize: 13, color: "#888" }}>Chưa có thay đổi giá nào được ghi nhận.</div> : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                <thead><tr style={{ background: "#f6f5f2", textAlign: "left" }}><th style={th}>Thời gian</th><th style={th}>Model</th><th style={th}>Trường</th><th style={th}>Giá trị cũ</th><th style={th}>Giá trị mới</th><th style={th}>Chênh lệch</th></tr></thead>
                <tbody>
                  {priceHistory.map((h) => {
                    const diff = Number(h.newValue) - Number(h.oldValue);
                    return (
                      <tr key={h.id} style={{ borderBottom: "1px solid #eee" }}>
                        <td style={td}>{formatDate(h.date)}</td><td style={{ ...td, fontWeight: 600 }}>{h.modelName}</td>
                        <td style={td}>{h.field === "giaVon" ? "Giá vốn" : "Giá bán"}</td><td style={td}>{formatVND(h.oldValue)}</td><td style={td}>{formatVND(h.newValue)}</td>
                        <td style={{ ...td, color: diff > 0 ? GREEN : diff < 0 ? RED : BLACK, fontWeight: 600 }}>{diff > 0 ? "+" : ""}{formatVND(diff)}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </section>
      )}

      {tab === "actuals" && (
        <section>
          <div style={{ border: "1px solid #eee", borderRadius: 10, padding: 14, marginBottom: 18, background: "#fafafa" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}><Upload size={15} color={RED} /><span style={{ fontSize: 13, fontWeight: 700 }}>Nhập báo cáo bán hàng tháng (file ERP thật)</span></div>
            <div style={{ fontSize: 12, color: "#666", marginBottom: 8 }}>Đọc đúng cột file gốc. Có thể nhập nhiều file/nhiều tháng, dữ liệu cộng dồn.</div>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleImportActuals} style={{ fontSize: 12 }} />
            {actualsImportMsg && <div style={{ fontSize: 12, marginTop: 8, color: RED }}>{actualsImportMsg}</div>}
          </div>
          {actualRecords.length > 0 && (
            <>
              <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 18 }}>
                <span style={{ fontSize: 12, color: "#888" }}>Lọc theo tháng:</span>
                <select style={{ ...inputSm, flex: "none", width: 160 }} value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}>
                  <option value="all">Tất cả các tháng</option>
                  {availableMonths.map((m) => <option key={m} value={m}>{m}</option>)}
                </select>
              </div>
              <SectionTitle>Xu hướng theo tháng</SectionTitle>
              <div style={{ overflowX: "auto", marginBottom: 24 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                  <thead><tr style={{ background: "#f6f5f2", textAlign: "left" }}><th style={th}>Tháng</th><th style={th}>Số lượng</th><th style={th}>Doanh thu</th><th style={th}>Giá vốn</th><th style={th}>Lợi nhuận</th><th style={th}>% LN/Doanh thu</th></tr></thead>
                  <tbody>
                    {monthlyTrend.map(([month, v]) => {
                      const pct = v.revenue > 0 ? (v.profit / v.revenue) * 100 : 0;
                      return (<tr key={month} style={{ borderBottom: "1px solid #eee" }}><td style={{ ...td, fontWeight: 600 }}>{month}</td><td style={td}>{v.qty}</td><td style={td}>{formatVND(v.revenue)}</td><td style={td}>{formatVND(v.cost)}</td><td style={{ ...td, fontWeight: 600 }}>{formatVND(v.profit)}</td><td style={{ ...td, color: pct >= 0 ? GREEN : RED, fontWeight: 600 }}>{pct.toFixed(1)}%</td></tr>);
                    })}
                  </tbody>
                </table>
              </div>
              <AggTable title="Lợi nhuận theo Nhà phân phối / Khách hàng" data={aggByDist} />
              <AggTable title="Lợi nhuận theo Model" data={aggByModel} />
              <SectionTitle>Lợi nhuận theo Chương trình khuyến mại — tự động tách theo NPP áp dụng</SectionTitle>
              <div style={{ marginBottom: 20 }}>
                {Object.entries(aggByProgram).sort((a, b) => b[1].revenue - a[1].revenue).map(([code, v]) => {
                  const pct = v.revenue > 0 ? (v.profit / v.revenue) * 100 : 0;
                  const distEntries = Object.entries(v.byDist).sort((a, b) => b[1].revenue - a[1].revenue);
                  const isOpen = expandedProgram === code;
                  return (
                    <div key={code} style={{ border: "1px solid #eee", borderRadius: 8, marginBottom: 8 }}>
                      <button onClick={() => setExpandedProgram(isOpen ? null : code)} style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, padding: "10px 12px", background: "#fafafa", border: "none", cursor: "pointer", textAlign: "left" }}>
                        <span style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5 }}>{isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}<strong>{v.name}</strong><span style={{ color: "#999", fontSize: 11 }}>({distEntries.length} NPP · {v.models.size} model)</span></span>
                        <span style={{ display: "flex", gap: 16, fontSize: 12.5 }}><span>SL {v.qty}</span><span>DT {formatVND(v.revenue)}</span><span style={{ color: v.profit >= 0 ? GREEN : RED, fontWeight: 600 }}>LN {formatVND(v.profit)} ({pct.toFixed(1)}%)</span></span>
                      </button>
                      {isOpen && (
                        <div style={{ padding: "8px 12px" }}>
                          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                            <thead><tr style={{ textAlign: "left", color: "#888" }}><th style={th}>Nhà phân phối</th><th style={th}>Số lượng</th><th style={th}>Doanh thu</th><th style={th}>Lợi nhuận</th></tr></thead>
                            <tbody>{distEntries.map(([dname, dv]) => (<tr key={dname} style={{ borderTop: "1px dashed #eee" }}><td style={td}>{dname}</td><td style={td}>{dv.qty}</td><td style={td}>{formatVND(dv.revenue)}</td><td style={{ ...td, color: dv.profit >= 0 ? GREEN : RED }}>{formatVND(dv.profit)}</td></tr>))}</tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </>
          )}
          {actualRecords.length === 0 && <div style={{ fontSize: 13, color: "#888" }}>Chưa có dữ liệu thực tế. Upload file báo cáo tháng ở trên để bắt đầu.</div>}
        </section>
      )}

      {tab === "program" && activeProgram && (
        <section>
          <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
            {[...programs].sort((a, b) => (a.startDate || "").localeCompare(b.startDate || "")).map((p) => {
              const st = getProgramStatus(p);
              return (
                <button key={p.id} onClick={() => setActiveProgramId(p.id)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 14px", borderRadius: 6, border: `1px solid ${p.id === activeProgramId ? RED : "#ddd"}`, background: p.id === activeProgramId ? RED : "#fff", color: p.id === activeProgramId ? "#fff" : BLACK, fontSize: 13, cursor: "pointer" }}>
                  <span style={{ width: 7, height: 7, borderRadius: "50%", background: st.color, display: "inline-block" }} />
                  <span style={{ opacity: 0.7, fontSize: 11 }}>#{p.seq || "?"}</span> {p.name}
                  <span style={{ opacity: 0.6, fontSize: 10.5 }}>({p.startDate ? new Date(p.startDate).toLocaleDateString("vi-VN") : "?"})</span>
                </button>
              );
            })}
            <button onClick={addProgram} style={{ ...addBtn, marginLeft: 4 }}><Plus size={14} /> Chương trình mới</button>
          </div>

          <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
            <input style={{ ...inputSm, fontWeight: 600, fontSize: 15, flex: 1, minWidth: 180 }} value={activeProgram.name} onChange={(e) => updateProgram(activeProgram.id, { name: e.target.value })} />
            <button onClick={() => exportProgramExcel(activeProgram, calcEntry, brands.find((b) => b.id === activeBrandId)?.name || "")} style={addBtn}><FileDown size={14} /> Xuất Excel</button>
            <button onClick={() => removeProgram(activeProgram.id)} style={iconBtn}><Trash2 size={14} /></button>
          </div>

          <div style={{ display: "flex", gap: 16, alignItems: "center", marginBottom: 20, flexWrap: "wrap" }}>
            <LabeledDate label="Ngày bắt đầu" value={activeProgram.startDate} onChange={(v) => updateProgram(activeProgram.id, { startDate: v })} />
            <LabeledDate label="Ngày kết thúc" value={activeProgram.endDate} onChange={(v) => updateProgram(activeProgram.id, { endDate: v })} />
            <StatusBadge status={getProgramStatus(activeProgram)} />
          </div>

          <div style={{ border: "1px solid #eee", borderRadius: 10, padding: 16, marginBottom: 20 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}><Box size={15} color={GREEN} /><span style={{ fontSize: 13, fontWeight: 700, color: GREEN, textTransform: "uppercase", letterSpacing: 0.5 }}>Model tính riêng lẻ (không thuộc gói)</span></div>
            <AddModelSelect models={models} excludeIds={new Set(activeProgram.standaloneEntries.map((e) => e.modelId))} onAdd={addStandaloneModel} />
            <EntryTable entries={activeProgram.standaloneEntries} models={models} gtCostCatalog={gtCostCatalog} mtCostCatalog={mtCostCatalog} calcEntry={calcEntry} onRemove={removeStandaloneEntry} onUpdateEntry={updateStandaloneEntry} onUpdateCost={updateStandaloneCost} />
          </div>

          <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
            <button onClick={() => addPackage("machine")} style={addBtn}><Package size={14} /> + Gói máy</button>
            <button onClick={() => addPackage("sales")} style={addBtn}><Package size={14} /> + Gói doanh số</button>
          </div>

          {activeProgram.packages.map((pkg) => {
            const entryResults = pkg.entries.map(calcEntry).filter(Boolean);
            const totalGiaBan = entryResults.reduce((s, r) => s + r.thanhTien, 0);
            const totalLoiNhuan = entryResults.reduce((s, r) => s + r.tongLoiNhuan, 0);
            return (
              <div key={pkg.id} style={{ border: "1px solid #eee", borderRadius: 10, padding: 16, marginBottom: 20 }}>
                <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", padding: "3px 8px", borderRadius: 4, background: pkg.type === "sales" ? "#FCEBEC" : "#EFF3F2", color: pkg.type === "sales" ? RED : GREEN }}>{pkg.type === "sales" ? "Gói doanh số" : "Gói máy"}</span>
                  <input style={{ ...inputSm, fontWeight: 600, flex: 1, minWidth: 160 }} value={pkg.name} onChange={(e) => updatePackage(pkg.id, { name: e.target.value })} />
                  {pkg.type === "sales" && <LabeledNum label="Mục tiêu doanh số" value={pkg.salesTarget} onChange={(v) => updatePackage(pkg.id, { salesTarget: v })} />}
                  <button onClick={() => removePackage(pkg.id)} style={iconBtn}><Trash2 size={14} /></button>
                </div>
                <AddModelSelect models={models} excludeIds={new Set(pkg.entries.map((e) => e.modelId))} onAdd={(modelId) => addModelToPackage(pkg.id, modelId)} />
                <EntryTable
                  entries={pkg.entries} models={models} gtCostCatalog={gtCostCatalog} mtCostCatalog={mtCostCatalog} calcEntry={calcEntry}
                  onRemove={(modelId) => removePackageEntry(pkg.id, modelId)}
                  onUpdateEntry={(modelId, patch) => updatePackageEntry(pkg.id, modelId, patch)}
                  onUpdateCost={(modelId, catalogKey, costId, patch) => updatePackageCost(pkg.id, modelId, catalogKey, costId, patch)}
                />
                {pkg.type === "sales" && pkg.entries.length > 0 && (
                  <div style={{ marginTop: 12, fontSize: 12.5, display: "flex", gap: 20, flexWrap: "wrap" }}>
                    <div>Tổng giá bán gói (thành tiền): <strong>{formatVND(totalGiaBan)}</strong></div>
                    <div>Mục tiêu doanh số: <strong>{formatVND(pkg.salesTarget)}</strong></div>
                    <div style={{ color: totalGiaBan >= pkg.salesTarget ? GREEN : RED, fontWeight: 600 }}>{totalGiaBan >= pkg.salesTarget ? "Đạt mục tiêu doanh số" : "Chưa đạt mục tiêu doanh số"}</div>
                    <div>Tổng lợi nhuận gói: <strong>{formatVND(totalLoiNhuan)}</strong></div>
                  </div>
                )}
                {pkg.type === "machine" && pkg.entries.length > 0 && (
                  <div style={{ marginTop: 12, fontSize: 12.5 }}>Tổng lợi nhuận gói: <strong>{formatVND(totalLoiNhuan)}</strong></div>
                )}
              </div>
            );
          })}
        </section>
      )}
    </div>
  );
}

function AggTable({ title, data }) {
  const rows = Object.entries(data).sort((a, b) => b[1].revenue - a[1].revenue);
  return (
    <div style={{ marginBottom: 22 }}>
      <SectionTitle>{title}</SectionTitle>
      {rows.length === 0 ? <div style={{ fontSize: 13, color: "#888", marginBottom: 10 }}>Chưa có dữ liệu.</div> : (
        <div style={{ overflowX: "auto", marginBottom: 10 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
            <thead><tr style={{ background: "#f6f5f2", textAlign: "left" }}><th style={th}>Tên</th><th style={th}>Số lượng</th><th style={th}>Doanh thu</th><th style={th}>Tổng giá vốn</th><th style={th}>Lợi nhuận</th><th style={th}>% LN</th></tr></thead>
            <tbody>
              {rows.map(([name, v]) => {
                const pct = v.revenue > 0 ? (v.profit / v.revenue) * 100 : 0;
                return (<tr key={name} style={{ borderBottom: "1px solid #eee" }}><td style={{ ...td, fontWeight: 600 }}>{name}</td><td style={td}>{v.qty}</td><td style={td}>{formatVND(v.revenue)}</td><td style={td}>{formatVND(v.cost)}</td><td style={{ ...td, fontWeight: 600 }}>{formatVND(v.profit)}</td><td style={{ ...td, color: pct >= 0 ? GREEN : RED, fontWeight: 600 }}>{pct.toFixed(1)}%</td></tr>);
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
function StatusBadge({ status }) {
  return (<span style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 700, padding: "4px 10px", borderRadius: 6, background: `${status.color}1A`, color: status.color }}><span style={{ width: 7, height: 7, borderRadius: "50%", background: status.color, display: "inline-block" }} />{status.label}</span>);
}
function LabeledDate({ label, value, onChange }) {
  return (<div style={{ display: "flex", flexDirection: "column", gap: 2 }}><span style={{ fontSize: 10, color: "#888" }}>{label}</span><input type="date" value={value || ""} onChange={(e) => onChange(e.target.value)} style={inputSm} /></div>);
}
function AddModelSelect({ models, excludeIds, onAdd }) {
  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 12, flexWrap: "wrap" }}>
      <span style={{ fontSize: 12, color: "#888" }}>Thêm model:</span>
      <select style={{ ...inputSm, flex: "none", width: 220 }} value="" onChange={(e) => e.target.value && onAdd(e.target.value)}>
        <option value="">-- Chọn model --</option>
        {models.filter((m) => !excludeIds.has(m.id)).map((m) => <option key={m.id} value={m.id}>{m.name} ({m.group})</option>)}
      </select>
    </div>
  );
}

function EntryTable({ entries, models, gtCostCatalog, mtCostCatalog, calcEntry, onRemove, onUpdateEntry, onUpdateCost }) {
  if (entries.length === 0) return <div style={{ fontSize: 13, color: "#888" }}>Chưa có model nào.</div>;

  const safeEntries = entries.map((e) => {
    const withGt = ensureCostsFor(e, "costsGT", gtCostCatalog);
    return ensureCostsFor(withGt, "costsMT", mtCostCatalog);
  });
  const results = safeEntries.map(calcEntry).filter(Boolean);
  const usedCostKeys = [];
  results.forEach((r) => Object.entries(r.perCostDisplay || {}).forEach(([key, v]) => { if (!usedCostKeys.find((u) => u.key === key)) usedCostKeys.push({ key, name: v.name }); }));

  return (
    <div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
          <thead>
            <tr style={{ background: "#f6f5f2", textAlign: "left" }}>
              <th style={th}>Model</th><th style={th}>Kênh</th><th style={th}>Giá vốn</th>
              <th style={th}>% tỷ lệ tối thiểu</th><th style={th}>LN thực tối thiểu thu về</th><th style={th}>Net tối thiểu (+VAT)</th><th style={th}>LN gộp %</th>
              <th style={th}>Giá bán (NPP)</th>
              <th style={th}>Mua</th><th style={th}>Tặng</th><th style={th}>NET Đại lý</th>
              {usedCostKeys.map((c) => <th key={c.key} style={th}>{c.name}</th>)}
              <th style={th}>Giá NET NPP</th><th style={th}>Giá nội bộ</th><th style={th}>Chi phí KM nội bộ</th><th style={th}>% CP</th>
              <th style={th}>LN/máy</th><th style={th}>%LN</th><th style={th}>%LN khoán</th><th style={th}>Chênh lệch</th><th style={th}>Trạng thái</th>
              <th style={th}>Tổng SL</th><th style={th}>Thành tiền</th><th style={th}>Tổng CP KM</th><th style={th}>Tổng LN thu về</th><th style={th}></th>
            </tr>
          </thead>
          <tbody>
            {safeEntries.map((entry) => {
              const r = calcEntry(entry);
              if (!r) return null;
              const isGT = entry.channel !== "MT";
              const netMinVal = entry.netMinOverride !== null && entry.netMinOverride !== undefined ? entry.netMinOverride : (r.model.minMarginTarget ?? 30);
              const khoanVal = entry.targetOverride !== null && entry.targetOverride !== undefined ? entry.targetOverride : r.model.target;
              return (
                <tr key={entry.modelId} style={{ borderBottom: "1px solid #eee", verticalAlign: "top" }}>
                  <td style={{ ...td, fontWeight: 600 }}>{r.model.name}</td>
                  <td style={td}>
                    <select style={inputXs} value={entry.channel} onChange={(e) => onUpdateEntry(entry.modelId, { channel: e.target.value })}>
                      <option value="GT">GT/Delta</option>
                      <option value="MT">MT (siêu thị)</option>
                    </select>
                  </td>
                  <td style={td}>{formatVND(r.model.giaVon)}</td>
                  <td style={td}>
                    <NumInput style={{ ...inputXs, width: 52 }} value={netMinVal} allowEmpty
                      title="% tỷ lệ tối thiểu (trên giá vốn) — ra Net tối thiểu/LN thực tối thiểu. Độc lập với %LN khoán. Để trống = dùng mặc định của model."
                      onChange={(v) => onUpdateEntry(entry.modelId, { netMinOverride: v })} />%
                  </td>
                  <td style={td}>{formatVND(r.lnThucToiThieu)}</td>
                  <td style={td}>{formatVND(r.netToiThieu)}</td>
                  <td style={td}>{r.lnGopPct.toFixed(1)}%</td>
                  <td style={td}>{formatVND(r.model.giaBan)}</td>
                  <td style={td}><NumInput style={{ ...inputXs, width: 50 }} value={entry.muaQty} onChange={(v) => onUpdateEntry(entry.modelId, { muaQty: v ?? 0 })} /></td>
                  <td style={td}>{isGT ? <NumInput style={{ ...inputXs, width: 50 }} value={entry.tangQty} onChange={(v) => onUpdateEntry(entry.modelId, { tangQty: v ?? 0 })} /> : <span style={{ color: "#ccc" }}>—</span>}</td>
                  <td style={td}>{formatVND(r.netDaiLy)}</td>
                  {usedCostKeys.map((c) => {
                    const rc = r.perCostDisplay?.[c.key];
                    return <td key={c.key} style={td}>{rc ? `${formatVND(rc.amount)}${rc.internalOnly ? " *" : ""}` : <span style={{ color: "#ccc" }}>—</span>}</td>;
                  })}
                  <td style={td}>{formatVND(r.netPhanTich)}</td>
                  <td style={td}>{formatVND(r.giaNoiBo)}</td>
                  <td style={td}>{formatVND(r.chiPhiKM)}</td>
                  <td style={td}>{r.pctCP.toFixed(1)}%</td>
                  <td style={{ ...td, fontWeight: 600 }}>{formatVND(r.perUnitProfit)}</td>
                  <td style={{ ...td, color: r.ok ? GREEN : RED, fontWeight: 600 }}>{r.marginPct.toFixed(1)}%</td>
                  <td style={td}>
                    <NumInput style={{ ...inputXs, width: 52 }} value={khoanVal} allowEmpty
                      title="%LN khoán (trên giá bán) — ngưỡng so Đạt/Dưới mục tiêu. Độc lập với % tỷ lệ tối thiểu. Để trống = dùng mặc định của model."
                      onChange={(v) => onUpdateEntry(entry.modelId, { targetOverride: v })} />%
                  </td>
                  <td style={{ ...td, color: r.chenhLech >= 0 ? GREEN : RED, fontWeight: 600 }}>{r.chenhLech >= 0 ? "+" : ""}{formatVND(r.chenhLech)}</td>
                  <td style={td}>{r.ok ? <span style={{ display: "flex", alignItems: "center", gap: 4, color: GREEN }}><CheckCircle2 size={14} /> Đạt</span> : <span style={{ display: "flex", alignItems: "center", gap: 4, color: RED }}><AlertTriangle size={14} /> Dưới mục tiêu</span>}</td>
                  <td style={td}>{r.tongSL}</td>
                  <td style={td}>{formatVND(r.thanhTien)}</td>
                  <td style={td}>{formatVND(r.tongChiPhiKM)}</td>
                  <td style={{ ...td, fontWeight: 700 }}>{formatVND(r.tongLoiNhuan)}</td>
                  <td style={td}><button onClick={() => onRemove(entry.modelId)} style={iconBtn}><Trash2 size={12} /></button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <div style={{ fontSize: 10.5, color: "#aaa", marginTop: 4 }}>* chi phí chỉ tính nội bộ, không trừ vào Giá NET NPP. "Tổng SL" chỉ cộng Mua+Tặng khi hàng tặng CÙNG loại với hàng mua.</div>

      <div style={{ marginTop: 10 }}>
        <div style={{ fontSize: 11, color: "#888", marginBottom: 6 }}>Chọn chi phí áp dụng riêng cho từng model (theo đúng kênh của model đó):</div>
        {safeEntries.map((entry) => {
          const model = models.find((m) => m.id === entry.modelId);
          const isGT = entry.channel !== "MT";
          const catalog = isGT ? gtCostCatalog : mtCostCatalog;
          const catalogKey = isGT ? "costsGT" : "costsMT";
          return (
            <div key={entry.modelId} style={{ padding: "8px 0", borderTop: "1px dashed #eee" }}>
              <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 4 }}>
                {model?.name} <span style={{ fontWeight: 400, color: "#999" }}>({isGT ? "kênh GT/Delta" : "kênh MT — gộp thành Chiết khấu Sellout"})</span>
                {isGT && (
                  <span style={{ fontWeight: 400, color: "#999" }}>
                    {" "}· Model hàng tặng:{" "}
                    <select style={inputXs} value={entry.giftModelId || ""} onChange={(e) => onUpdateEntry(entry.modelId, { giftModelId: e.target.value })}>
                      <option value="">(cùng loại — giống model mua)</option>
                      {models.filter((m) => m.id !== entry.modelId).map((m) => <option key={m.id} value={m.id}>{m.name} (khác loại)</option>)}
                    </select>
                  </span>
                )}
              </div>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                {catalog.map((c) => {
                  const line = entry[catalogKey][c.id] || { on: false, mode: "percent", value: 0 };
                  return (
                    <label key={c.id} style={{ display: "flex", alignItems: "center", gap: 3, fontSize: 11.5 }}>
                      <input type="checkbox" checked={line.on} onChange={(e) => onUpdateCost(entry.modelId, catalogKey, c.id, { on: e.target.checked })} />
                      {c.name}{c.group === "so_support" ? " *" : ""}{c.internalOnly ? " (nội bộ)" : ""}
                      {line.on && (
                        <span style={{ display: "flex", gap: 2 }}>
                          <select style={inputXs} value={line.mode} onChange={(e) => onUpdateCost(entry.modelId, catalogKey, c.id, { mode: e.target.value })}>
                            <option value="percent">%</option><option value="fixed">đ</option>
                          </select>
                          <NumInput style={{ ...inputXs, width: 60 }} value={line.value} onChange={(v) => onUpdateCost(entry.modelId, catalogKey, c.id, { value: v ?? 0 })} />
                        </span>
                      )}
                    </label>
                  );
                })}
              </div>
              {!isGT && <div style={{ fontSize: 10.5, color: "#aaa", marginTop: 3 }}>* nhóm Hỗ trợ SO chỉ tính giá trị LỚN NHẤT, không cộng dồn.</div>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function TabButton({ active, onClick, icon, label }) {
  return (<button onClick={onClick} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", borderRadius: 6, border: `1px solid ${active ? RED : "#ddd"}`, background: active ? RED : "#fff", color: active ? "#fff" : BLACK, fontSize: 13, fontWeight: 500, cursor: "pointer" }}>{icon} {label}</button>);
}
function SectionTitle({ children }) { return <h2 style={{ fontSize: 14, fontWeight: 700, marginBottom: 10, textTransform: "uppercase", letterSpacing: 0.5, color: "#5F5E5A" }}>{children}</h2>; }
function LabeledNum({ label, value, onChange, small }) {
  return (<div style={{ display: "flex", flexDirection: "column", gap: 2 }}><span style={{ fontSize: 10, color: "#888" }}>{label}</span><NumInput value={value} onChange={onChange} style={{ ...inputSm, width: small ? 56 : 140 }} /></div>);
}

const rowCard = { display: "flex", gap: 8, alignItems: "center", padding: "8px 10px", border: "1px solid #eee", borderRadius: 8, marginBottom: 8, flexWrap: "wrap" };
const inputSm = { border: "1px solid #ddd", borderRadius: 6, padding: "6px 8px", fontSize: 13, flex: 1 };
const inputXs = { border: "1px solid #ddd", borderRadius: 5, padding: "3px 5px", fontSize: 12 };
const iconBtn = { border: "1px solid #eee", borderRadius: 6, background: "#fff", padding: 6, cursor: "pointer", color: "#888" };
const addBtn = { display: "flex", alignItems: "center", gap: 4, border: `1px dashed ${RED}`, color: RED, background: "#fff", borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer" };
const th = { padding: "8px 8px", fontWeight: 600, whiteSpace: "nowrap" };
const td = { padding: "8px 8px", whiteSpace: "nowrap" };
