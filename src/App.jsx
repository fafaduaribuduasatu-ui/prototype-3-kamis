import { useState, useEffect, useMemo, useRef } from "react";
import { Plus, Radio, Settings2, X, Trash2, Lock, Unlock, ChevronDown, ChevronRight, Crown } from "lucide-react";
import { BarChart, Bar, LineChart, Line, Legend, XAxis, YAxis, ResponsiveContainer, Tooltip, CartesianGrid } from "recharts";
import { FileDown } from "lucide-react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";

const FONT_STYLE = `
@import url('https://fonts.googleapis.com/css2?family=Fredoka:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');
`;

const JAM_OPTIONS = ["05.00","07.00","09.00","11.00","13.00","15.00","17.00","19.00"];
const SESI_OPTIONS = ["SESI 1","SESI 2","SESI 3 (overtime)"];
const TARGET_OPTIONS = ["0 - 2JT","2 - 3JT","3 - 4JT","4 - 5JT","5 - 6JT","7 - 8JT","8 - 9JT","9 - 10JT"];
const DEFAULT_HOSTS = ["ALIN","APRIL","AUFA","HILDA"];

const ACCENT = "#FE2C55";
const CYAN = "#25F4EE";
const GOOD = "#25F4EE";
const BAD = "#FFB800";
const HOST_COLORS = ["#25F4EE", "#FE2C55", "#FFB800", "#8B5CF6", "#4ADE80", "#FF7A5C", "#60A5FA"];
const BASE_UPAH = 25000;
const BONUS_PER_JT = 5000;

function hitungBonus(gmv) {
  const tier = Math.floor((gmv || 0) / 1_000_000);
  return tier >= 2 ? (tier - 1) * BONUS_PER_JT : 0;
}
function hitungTotalGaji(gmv) {
  return BASE_UPAH + hitungBonus(gmv);
}

function parseTargetRange(t) {
  const nums = (t.match(/\d+/g) || []).map(Number);
  if (nums.length < 2) return [0, 0];
  return [nums[0] * 1_000_000, nums[1] * 1_000_000];
}

function formatRupiah(n) {
  return "Rp " + Number(n || 0).toLocaleString("id-ID");
}

function nowTimestamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}
function formatTanggalDMY(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

function excelSerialToISO(serial) {
  const utcDays = Math.floor(serial - 25569);
  const date = new Date(utcDays * 86400 * 1000);
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, "0");
  const d = String(date.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}
function parseTanggalFleksibel(val) {
  if (typeof val === "number") return excelSerialToISO(val);
  if (typeof val === "string") {
    let m = val.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
    m = val.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  }
  return "";
}
function parseGmvFleksibel(val) {
  if (typeof val === "number") return Math.round(val);
  if (typeof val === "string") {
    const cleaned = val.replace(/[^\d]/g, "");
    return cleaned ? parseInt(cleaned, 10) : 0;
  }
  return 0;
}
function cariKolom(row, keywords) {
  const keys = Object.keys(row);
  for (const k of keys) {
    const norm = k.toLowerCase().replace(/[^a-z0-9]/g, "");
    if (keywords.some((kw) => norm.includes(kw))) return k;
  }
  return null;
}
function parseBarisExcel(rows) {
  const hasil = [];
  for (const row of rows) {
    const hostKey = cariKolom(row, ["host", "nama"]);
    const tglKey = cariKolom(row, ["tanggal", "tgl", "date"]);
    const jamKey = cariKolom(row, ["jam"]);
    const sesiKey = cariKolom(row, ["sesi"]);
    const gmvKey = cariKolom(row, ["gmv", "perolehan"]);
    const targetKey = cariKolom(row, ["target"]);
    const tsKey = cariKolom(row, ["timestamp"]);
    if (!hostKey || !tglKey || !gmvKey) continue;

    const host = String(row[hostKey] || "").trim().toUpperCase();
    const tanggal = parseTanggalFleksibel(row[tglKey]);
    if (!host || !tanggal) continue;

    let sesi = sesiKey ? String(row[sesiKey]).trim().toUpperCase() : "SESI 1";
    if (sesi && !sesi.startsWith("SESI")) sesi = `SESI ${sesi}`;

    hasil.push({
      host,
      tanggal,
      jamMulai: jamKey ? String(row[jamKey]).trim() : "",
      sesi,
      gmv: parseGmvFleksibel(row[gmvKey]),
      target: targetKey && row[targetKey] ? String(row[targetKey]) : "0 - 2JT",
      timestamp: tsKey ? String(row[tsKey]) : nowTimestamp(),
    });
  }
  return hasil;
}

async function loadShared(key, fallback) {
  try {
    const { data, error } = await supabase.from("kv_store").select("value").eq("key", key).maybeSingle();
    if (error || !data) return fallback;
    return data.value;
  } catch {
    return fallback;
  }
}
async function saveShared(key, value) {
  try {
    await supabase.from("kv_store").upsert({ key, value, updated_at: new Date().toISOString() });
  } catch (e) {
    console.error("storage error", e);
  }
}

// --- Laporan sekarang disimpan sebagai baris sungguhan di tabel "laporan" ---
// (bukan lagi gumpalan JSON di kv_store), supaya submit/hapus dari HP lain
// tidak saling menimpa dan anti-dobel dijamin langsung oleh database.
async function loadEntries() {
  try {
    const { data, error } = await supabase
      .from("laporan")
      .select("*")
      .order("tanggal", { ascending: false })
      .order("jam_mulai", { ascending: false });
    if (error || !data) return [];
    return data.map((row) => ({
      id: row.id,
      host: row.host,
      tanggal: row.tanggal,
      jamMulai: row.jam_mulai,
      sesi: row.sesi,
      gmv: row.gmv,
      target: row.target,
      timestamp: row.timestamp,
    }));
  } catch {
    return [];
  }
}

export default function App() {
  const [tab, setTab] = useState("input");
  const [hosts, setHosts] = useState(DEFAULT_HOSTS);
  const [entries, setEntries] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showHostSettings, setShowHostSettings] = useState(false);
  const [newHostName, setNewHostName] = useState("");
  const [adminPin, setAdminPin] = useState("1234");
  const [isAdmin, setIsAdmin] = useState(false);
  const [showPinPrompt, setShowPinPrompt] = useState(false);
  const [pinInput, setPinInput] = useState("");
  const [pinError, setPinError] = useState(false);
  const [newPin, setNewPin] = useState("");
  const [excelMsg, setExcelMsg] = useState("");
  const fileInputRef = useRef(null);
  const [expandedHosts, setExpandedHosts] = useState({});
  function toggleHostExpand(host) {
    setExpandedHosts((prev) => ({ ...prev, [host]: !prev[host] }));
  }

  const [form, setForm] = useState({
    host: "", tanggal: "", jamMulai: "", sesi: "", gmv: "", target: "",
  });
  const [justSaved, setJustSaved] = useState(false);

  useEffect(() => {
    (async () => {
      const [h, e, pin] = await Promise.all([
        loadShared("live-hosts", DEFAULT_HOSTS),
        loadEntries(),
        loadShared("live-admin-pin", "1234"),
      ]);
      setHosts(h);
      setEntries(e);
      setAdminPin(pin);
      setLoading(false);
    })();
  }, []);

  // Auto-refresh tiap 30 detik, biar laporan/rekap/grafik ikut update
  // walau ada host lain yang input dari HP lain tanpa perlu refresh manual
  useEffect(() => {
    const interval = setInterval(async () => {
      const fresh = await loadEntries();
      setEntries(fresh);
    }, 30000);
    return () => clearInterval(interval);
  }, []);

  const [submitError, setSubmitError] = useState("");

  const isDuplicate = useMemo(() => {
    if (!form.host || !form.tanggal || !form.sesi) return false;
    return entries.some((e) => e.host === form.host && e.tanggal === form.tanggal && e.sesi === form.sesi);
  }, [form.host, form.tanggal, form.sesi, entries]);

  const canSubmit = form.host && form.tanggal && form.jamMulai && form.sesi && form.gmv && form.target && !isDuplicate;

  async function submitEntry() {
    if (!canSubmit) return;
    setSubmitError("");
    const { error } = await supabase.from("laporan").insert({
      host: form.host,
      tanggal: form.tanggal,
      jam_mulai: form.jamMulai,
      sesi: form.sesi,
      gmv: Number(form.gmv),
      target: form.target,
      timestamp: nowTimestamp(),
    });
    if (error) {
      if (error.code === "23505") {
        setSubmitError("Laporan ini sudah pernah diinput (mungkin baru saja oleh orang lain). Coba refresh dulu.");
      } else {
        setSubmitError("Gagal menyimpan, coba lagi.");
      }
      return;
    }
    const fresh = await loadEntries();
    setEntries(fresh);
    setForm({ host: "", tanggal: "", jamMulai: "", sesi: "", gmv: "", target: "" });
    setJustSaved(true);
    setTimeout(() => setJustSaved(false), 2000);
  }

  async function deleteEntry(id) {
    if (!isAdmin) return;
    const { error } = await supabase.from("laporan").delete().eq("id", id);
    if (!error) {
      const fresh = await loadEntries();
      setEntries(fresh);
    }
  }

  async function addHost() {
    const name = newHostName.trim().toUpperCase();
    if (!name || hosts.includes(name)) return;
    const next = [...hosts, name];
    setHosts(next);
    await saveShared("live-hosts", next);
    setNewHostName("");
  }
  async function removeHost(name) {
    const next = hosts.filter((h) => h !== name);
    setHosts(next);
    await saveShared("live-hosts", next);
  }

  function tryUnlock() {
    if (pinInput === adminPin) {
      setIsAdmin(true);
      setShowPinPrompt(false);
      setPinInput("");
      setPinError(false);
    } else {
      setPinError(true);
    }
  }

  async function changePin() {
    const p = newPin.trim();
    if (!p) return;
    setAdminPin(p);
    await saveShared("live-admin-pin", p);
    setNewPin("");
  }

  async function handleExcelFile(fileEvent) {
    const file = fileEvent.target.files?.[0];
    fileEvent.target.value = "";
    if (!file) return;
    setExcelMsg("Membaca file…");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      let allRows = [];
      wb.SheetNames.forEach((name) => {
        const sheet = wb.Sheets[name];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        allRows = allRows.concat(rows);
      });
      const parsed = parseBarisExcel(allRows);
      if (parsed.length === 0) {
        setExcelMsg("Tidak ada baris yang bisa dibaca. Pastikan ada kolom Nama Host, Tanggal, dan GMV.");
        return;
      }
      const newHosts = [...hosts];
      parsed.forEach((p) => { if (!newHosts.includes(p.host)) newHosts.push(p.host); });
      if (newHosts.length !== hosts.length) {
        setHosts(newHosts);
        await saveShared("live-hosts", newHosts);
      }
      const rows = parsed.map((p) => ({
        host: p.host, tanggal: p.tanggal, jam_mulai: p.jamMulai, sesi: p.sesi, gmv: p.gmv, target: p.target, timestamp: p.timestamp,
      }));
      const { data, error } = await supabase
        .from("laporan")
        .upsert(rows, { onConflict: "host,tanggal,sesi", ignoreDuplicates: true })
        .select();
      if (error) {
        setExcelMsg("Gagal mengimpor: " + error.message);
        return;
      }
      const berhasil = data ? data.length : 0;
      const fresh = await loadEntries();
      setEntries(fresh);
      setExcelMsg(`${berhasil} laporan berhasil diimpor dari Excel (${parsed.length - berhasil} dilewati karena duplikat).`);
    } catch (err) {
      console.error(err);
      setExcelMsg("Gagal membaca file. Pastikan formatnya .xlsx atau .xls.");
    }
  }

  const months = useMemo(() => {
    const set = new Set(entries.map((e) => e.tanggal?.slice(0, 7)).filter(Boolean));
    return Array.from(set).sort().reverse();
  }, [entries]);
  const [monthFilter, setMonthFilter] = useState("");
  useEffect(() => { if (months.length && !monthFilter) setMonthFilter(months[0]); }, [months]);

  const [remindStart, setRemindStart] = useState("");
  const [remindEnd, setRemindEnd] = useState("");
  useEffect(() => {
    const today = new Date();
    const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
    const toISO = (d) => d.toISOString().slice(0, 10);
    setRemindStart(toISO(firstDay));
    setRemindEnd(toISO(today));
  }, []);

  const filtered = useMemo(
    () => entries.filter((e) => !monthFilter || e.tanggal?.startsWith(monthFilter)),
    [entries, monthFilter]
  );

  const perHost = useMemo(() => {
    const map = {};
    hosts.forEach((h) => (map[h] = { host: h, total: 0, sesi: 0, tercapai: 0, gajiPokok: 0, bonus: 0, totalGaji: 0, hariSet: new Set() }));
    filtered.forEach((e) => {
      if (!map[e.host]) map[e.host] = { host: e.host, total: 0, sesi: 0, tercapai: 0, gajiPokok: 0, bonus: 0, totalGaji: 0, hariSet: new Set() };
      map[e.host].total += e.gmv;
      map[e.host].sesi += 1;
      map[e.host].gajiPokok += BASE_UPAH;
      map[e.host].bonus += hitungBonus(e.gmv);
      map[e.host].totalGaji += hitungTotalGaji(e.gmv);
      if (e.gmv >= 2_000_000) map[e.host].tercapai += 1;
      if (e.tanggal) map[e.host].hariSet.add(e.tanggal);
    });
    return Object.values(map)
      .map((p) => ({ ...p, hariKerja: p.hariSet.size }))
      .sort((a, b) => b.total - a.total);
  }, [filtered, hosts]);

  const topPerformer = perHost.length > 0 && perHost[0].total > 0 ? perHost[0].host : null;

  const dailyChartData = useMemo(() => {
    if (!monthFilter) return [];
    const [y, m] = monthFilter.split("-").map(Number);
    if (!y || !m) return [];
    const daysInMonth = new Date(y, m, 0).getDate();
    const arr = [];
    for (let d = 1; d <= daysInMonth; d++) {
      const row = { day: d };
      hosts.forEach((h) => { row[h] = 0; });
      arr.push(row);
    }
    filtered.forEach((e) => {
      const day = parseInt((e.tanggal || "").slice(8, 10), 10);
      if (day >= 1 && day <= daysInMonth) {
        arr[day - 1][e.host] = (arr[day - 1][e.host] || 0) + e.gmv;
      }
    });
    return arr;
  }, [filtered, monthFilter, hosts]);

  const detailByHost = useMemo(() => {
    const map = {};
    hosts.forEach((h) => (map[h] = []));
    filtered.forEach((e) => {
      if (!map[e.host]) map[e.host] = [];
      map[e.host].push(e);
    });
    Object.keys(map).forEach((h) => {
      map[h].sort((a, b) => (b.tanggal + b.jamMulai).localeCompare(a.tanggal + a.jamMulai));
    });
    return map;
  }, [filtered, hosts]);

  const missingReport = useMemo(() => {
    if (!remindStart || !remindEnd) return {};
    const result = {};
    hosts.forEach((h) => { result[h] = {}; });
    const start = new Date(remindStart);
    const end = new Date(remindEnd);
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      const iso = d.toISOString().slice(0, 10);
      hosts.forEach((h) => {
        const daySesi = entries.filter((e) => e.host === h && e.tanggal === iso).map((e) => e.sesi);
        const missing = ["SESI 1", "SESI 2"].filter((s) => !daySesi.includes(s));
        if (missing.length > 0) result[h][iso] = missing;
      });
    }
    return result;
  }, [remindStart, remindEnd, entries, hosts]);

  function generateReminderText() {
    const lines = [`Reminder Laporan Live (${formatTanggalDMY(remindStart)} - ${formatTanggalDMY(remindEnd)})`, ""];
    hosts.forEach((h) => {
      const missingDates = Object.keys(missingReport[h] || {}).sort();
      if (missingDates.length === 0) {
        lines.push(`${h} — lengkap ✅`);
      } else {
        lines.push(h);
        missingDates.forEach((iso) => {
          const [, m, dd] = iso.split("-");
          lines.push(`${dd}/${m} — ${missingReport[h][iso].join(", ")} belum input`);
        });
      }
      lines.push("");
    });
    lines.push("Mohon segera dilengkapi ya 🙏");
    return lines.join("\n");
  }

  function kirimReminderWA() {
    const text = generateReminderText();
    window.open(`https://wa.me/?text=${encodeURIComponent(text)}`, "_blank");
  }

  const grandTotal = filtered.reduce((s, e) => s + e.gmv, 0);
  const grandGajiPokok = filtered.length * BASE_UPAH;
  const grandBonus = filtered.reduce((s, e) => s + hitungBonus(e.gmv), 0);
  const grandTotalGaji = grandGajiPokok + grandBonus;

  function exportExcel() {
    const summaryRows = perHost.map((p) => ({
      Host: p.host,
      "Total GMV": p.total,
      "Jumlah Sesi": p.sesi,
      "Sesi Tercapai": p.tercapai,
      Gaji: p.gajiPokok,
      Bonus: p.bonus,
      "Total Gaji": p.totalGaji,
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "Rekap Host");

    hosts.forEach((host) => {
      const rows = filtered
        .filter((e) => e.host === host)
        .sort((a, b) => (a.tanggal + a.jamMulai).localeCompare(b.tanggal + b.jamMulai))
        .map((e) => ({
          TANGGAL: formatTanggalDMY(e.tanggal),
          "JAM MULAI": e.jamMulai,
          SESI: e.sesi,
          "LAPORAN PEROLEHAN GMV TIKTOK": e.gmv,
          TARGET: e.target,
          GAJI: BASE_UPAH,
          BONUS: hitungBonus(e.gmv),
          "TOTAL GAJI": hitungTotalGaji(e.gmv),
        }));
      if (rows.length === 0) return;
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), host.slice(0, 31));
    });

    XLSX.writeFile(wb, `laporan-live-${monthFilter || "semua"}.xlsx`);
  }

  return (
    <div style={{ minHeight: "100vh", background: "#0a0a0f", backgroundImage: "radial-gradient(circle at 15% 0%, rgba(37,244,238,0.15), transparent 45%), radial-gradient(circle at 90% 15%, rgba(254,44,85,0.18), transparent 40%)", fontFamily: "Inter, sans-serif", color: "#F5F1E8", paddingBottom: 40 }}>
      <style>{FONT_STYLE}</style>

      <header style={{ padding: "28px 20px 20px", borderBottom: "1px solid #24242f" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <Radio size={20} color={ACCENT} />
            <span style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12, letterSpacing: 2, color: CYAN, textTransform: "uppercase" }}>
              Live GMV Tracker
            </span>
          </div>
          <button
            onClick={() => (isAdmin ? setIsAdmin(false) : setShowPinPrompt(true))}
            title={isAdmin ? "Kunci mode admin" : "Buka mode admin"}
            style={{ background: "none", border: "none", color: isAdmin ? GOOD : "#8a8a9a", cursor: "pointer" }}>
            {isAdmin ? <Unlock size={18} /> : <Lock size={18} />}
          </button>
        </div>
        <h1 style={{
          fontFamily: "Fredoka, sans-serif", fontWeight: 600, fontSize: 30, margin: "10px 0 0",
          background: `linear-gradient(90deg, ${CYAN}, ${ACCENT})`,
          WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text",
        }}>
          Perolehan Live Streaming Faradisaaaa
        </h1>
        <div style={{ display: "flex", gap: 8, marginTop: 18 }}>
          {["input", "rekap"].map((t) => (
            <button key={t} onClick={() => setTab(t)}
              style={{
                flex: 1, padding: "10px 0", borderRadius: 10, border: "none", cursor: "pointer",
                fontFamily: "Inter", fontWeight: 600, fontSize: 14,
                background: tab === t ? `linear-gradient(135deg, ${ACCENT}, ${CYAN})` : "#1c1c26",
                color: tab === t ? "#14141c" : "#F5F1E8",
              }}>
              {t === "input" ? "Input Laporan" : "Rekap"}
            </button>
          ))}
        </div>
      </header>

      {loading ? (
        <div style={{ padding: 40, textAlign: "center", opacity: 0.6 }}>Memuat data…</div>
      ) : tab === "input" ? (
        <div style={{ padding: 20, maxWidth: 480, margin: "0 auto" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
            <label style={labelStyle}>NAMA HOST</label>
            {isAdmin && (
              <button onClick={() => setShowHostSettings(true)} style={{ background: "none", border: "none", color: "#8a8a9a", cursor: "pointer" }}>
                <Settings2 size={16} />
              </button>
            )}
          </div>
          <select value={form.host} onChange={(e) => setForm({ ...form, host: e.target.value })} style={selectStyle}>
            <option value="">Pilih host</option>
            {hosts.map((h) => <option key={h} value={h}>{h}</option>)}
          </select>

          <label style={labelStyle}>TANGGAL</label>
          <input type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} style={selectStyle} />

          <label style={labelStyle}>JAM MULAI</label>
          <select value={form.jamMulai} onChange={(e) => setForm({ ...form, jamMulai: e.target.value })} style={selectStyle}>
            <option value="">Pilih jam</option>
            {JAM_OPTIONS.map((j) => <option key={j} value={j}>{j}</option>)}
          </select>

          <label style={labelStyle}>SESI</label>
          <select value={form.sesi} onChange={(e) => setForm({ ...form, sesi: e.target.value })} style={selectStyle}>
            <option value="">Pilih sesi</option>
            {SESI_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>

          <label style={labelStyle}>LAPORAN PEROLEHAN GMV TIKTOK</label>
          <input type="number" placeholder="misal 8500000" value={form.gmv}
            onChange={(e) => setForm({ ...form, gmv: e.target.value })} style={{ ...selectStyle, fontFamily: "'JetBrains Mono', monospace" }} />

          <label style={labelStyle}>TARGET</label>
          <select value={form.target} onChange={(e) => setForm({ ...form, target: e.target.value })} style={selectStyle}>
            <option value="">Pilih target</option>
            {TARGET_OPTIONS.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>

          <button onClick={submitEntry} disabled={!canSubmit}
            style={{
              width: "100%", marginTop: 20, padding: "14px 0", borderRadius: 10, border: "none",
              background: canSubmit ? `linear-gradient(135deg, ${ACCENT}, ${CYAN})` : "#2a2a35", color: canSubmit ? "#0a0a0f" : "#77778a",
              fontWeight: 700, fontSize: 15, cursor: canSubmit ? "pointer" : "not-allowed",
              display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
            }}>
            <Plus size={18} /> Simpan Laporan
          </button>
          {isDuplicate && (
            <div style={{ marginTop: 10, textAlign: "center", color: BAD, fontSize: 13, fontWeight: 600 }}>
              ⚠️ Laporan {form.host} tanggal {formatTanggalDMY(form.tanggal)} {form.sesi} sudah pernah diinput.
            </div>
          )}
          {submitError && (
            <div style={{ marginTop: 10, textAlign: "center", color: BAD, fontSize: 13, fontWeight: 600 }}>
              ⚠️ {submitError}
            </div>
          )}
          {justSaved && <div style={{ marginTop: 10, textAlign: "center", color: GOOD, fontSize: 13, fontWeight: 600 }}>✨ Laporan tersimpan!</div>}
        </div>
      ) : (
        <div style={{ padding: 20, maxWidth: 640, margin: "0 auto" }}>
          <div style={{ display: "flex", gap: 8 }}>
            <select value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)} style={{ ...selectStyle, maxWidth: 220 }}>
              {months.length === 0 && <option value="">Belum ada data</option>}
              {months.map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
            {isAdmin && (
              <button onClick={exportExcel} disabled={filtered.length === 0}
                style={{
                  display: "flex", alignItems: "center", gap: 6, whiteSpace: "nowrap",
                  padding: "0 16px", borderRadius: 10, border: "1px solid #2a2a36",
                  background: filtered.length === 0 ? "#1c1c26" : GOOD,
                  color: filtered.length === 0 ? "#77778a" : "#14141c",
                  fontWeight: 600, fontSize: 13, cursor: filtered.length === 0 ? "not-allowed" : "pointer",
                }}>
                <FileDown size={16} /> Excel
              </button>
            )}
          </div>

          {isAdmin && (
            <div style={{ ...cardStyle, marginTop: 16 }}>
              <h3 style={{ fontFamily: "Fredoka, sans-serif", fontSize: 16, margin: "0 0 10px" }}>🔔 Cek Kelengkapan Laporan</h3>
              <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
                <div style={{ flex: 1 }}>
                  <label style={labelStyle}>DARI</label>
                  <input type="date" value={remindStart} onChange={(e) => setRemindStart(e.target.value)} style={selectStyle} />
                </div>
                <div style={{ flex: 1 }}>
                  <label style={labelStyle}>SAMPAI</label>
                  <input type="date" value={remindEnd} onChange={(e) => setRemindEnd(e.target.value)} style={selectStyle} />
                </div>
              </div>
              <div style={{ fontSize: 13, lineHeight: 1.7, marginBottom: 12 }}>
                {hosts.map((h) => {
                  const missingDates = Object.keys(missingReport[h] || {}).sort();
                  return (
                    <div key={h} style={{ marginBottom: 8 }}>
                      <div style={{ fontWeight: 600 }}>{h}</div>
                      {missingDates.length === 0 ? (
                        <div style={{ color: GOOD, fontSize: 12 }}>Lengkap ✅</div>
                      ) : (
                        missingDates.map((iso) => {
                          const [, m, dd] = iso.split("-");
                          return (
                            <div key={iso} style={{ color: BAD, fontSize: 12 }}>
                              {dd}/{m} — {missingReport[h][iso].join(", ")} belum input
                            </div>
                          );
                        })
                      )}
                    </div>
                  );
                })}
              </div>
              <button onClick={kirimReminderWA}
                style={{ width: "100%", padding: "12px 0", borderRadius: 10, border: "none", background: "#25D366", color: "#0a0a0f", fontWeight: 700, cursor: "pointer" }}>
                Kirim ke WhatsApp
              </button>
            </div>
          )}

          <div style={{ display: "flex", gap: 10, margin: "16px 0" }}>
            <div style={cardStyle}>
              <div style={{ fontSize: 12, color: "#8a8a9a" }}>TOTAL GMV</div>
              <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 18, fontWeight: 700, color: ACCENT }}>{formatRupiah(grandTotal)}</div>
            </div>
            <div style={cardStyle}>
              <div style={{ fontSize: 12, color: "#8a8a9a" }}>TOTAL SESI</div>
              <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 18, fontWeight: 700 }}>{filtered.length}</div>
            </div>
          </div>
          {isAdmin && (
            <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
              <div style={cardStyle}>
                <div style={{ fontSize: 12, color: "#8a8a9a" }}>GAJI</div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 16, fontWeight: 700 }}>{formatRupiah(grandGajiPokok)}</div>
              </div>
              <div style={cardStyle}>
                <div style={{ fontSize: 12, color: "#8a8a9a" }}>BONUS</div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 16, fontWeight: 700, color: ACCENT }}>{formatRupiah(grandBonus)}</div>
              </div>
              <div style={cardStyle}>
                <div style={{ fontSize: 12, color: "#8a8a9a" }}>TOTAL GAJI</div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 16, fontWeight: 700, color: GOOD }}>{formatRupiah(grandTotalGaji)}</div>
              </div>
            </div>
          )}

          <div style={{ height: 200, marginBottom: 20 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={perHost}>
                <CartesianGrid strokeDasharray="3 3" stroke="#2a2a36" />
                <XAxis dataKey="host" tick={{ fill: "#8a8a9a", fontSize: 11 }} />
                <YAxis tick={{ fill: "#8a8a9a", fontSize: 10 }} tickFormatter={(v) => `${v / 1e6}jt`} />
                <Tooltip formatter={(v) => formatRupiah(v)} contentStyle={{ background: "#20202b", border: "none", borderRadius: 8 }} />
                <Bar dataKey="total" fill={CYAN} radius={[6, 6, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>

          <h3 style={{ fontFamily: "Fredoka, sans-serif", fontSize: 16, marginBottom: 10, color: "#8a8a9a" }}>Tren GMV Harian</h3>
          <div style={{ height: 220, marginBottom: 20 }}>
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={dailyChartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#2a2a36" />
                <XAxis dataKey="day" tick={{ fill: "#8a8a9a", fontSize: 10 }} />
                <YAxis tick={{ fill: "#8a8a9a", fontSize: 10 }} tickFormatter={(v) => `${v / 1e6}jt`} />
                <Tooltip formatter={(v) => formatRupiah(v)} labelFormatter={(d) => `Tanggal ${d}`} contentStyle={{ background: "#20202b", border: "none", borderRadius: 8 }} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                {hosts.map((h, i) => (
                  <Line key={h} type="monotone" dataKey={h} stroke={HOST_COLORS[i % HOST_COLORS.length]} strokeWidth={2} dot={false} />
                ))}
              </LineChart>
            </ResponsiveContainer>
          </div>

          <h3 style={{ fontFamily: "Fredoka, sans-serif", fontSize: 18, marginBottom: 10 }}>Rekap per Host</h3>
          {perHost.map((p) => (
            <div key={p.host} style={{ ...cardStyle, display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div>
                <div style={{ fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
                  {p.host}
                  {topPerformer === p.host && <Crown size={14} color={BAD} />}
                </div>
                <div style={{ fontSize: 12, color: "#8a8a9a" }}>{p.hariKerja} hari kerja · {p.sesi} sesi · {p.tercapai} tercapai target</div>
                {isAdmin && (
                  <div style={{ fontSize: 11, color: "#8a8a9a", marginTop: 2 }}>
                    Gaji {formatRupiah(p.gajiPokok)} + Bonus {formatRupiah(p.bonus)}
                  </div>
                )}
              </div>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontWeight: 700, color: ACCENT }}>{formatRupiah(p.total)}</div>
                {isAdmin && (
                  <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12, color: GOOD }}>Total: {formatRupiah(p.totalGaji)}</div>
                )}
              </div>
            </div>
          ))}

          <h3 style={{ fontFamily: "Fredoka, sans-serif", fontSize: 18, margin: "20px 0 10px" }}>Detail Laporan</h3>
          {filtered.length === 0 && <div style={{ opacity: 0.5, fontSize: 14 }}>Belum ada laporan bulan ini.</div>}
          {hosts.filter((h) => (detailByHost[h] || []).length > 0).map((host) => {
            const hostEntries = detailByHost[host] || [];
            const isOpen = !!expandedHosts[host];
            return (
              <div key={host} style={{ ...cardStyle, marginBottom: 8, padding: 0, overflow: "hidden" }}>
                <button onClick={() => toggleHostExpand(host)}
                  style={{ width: "100%", background: "none", border: "none", color: "#F5F1E8", cursor: "pointer", padding: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
                    {isOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                    {host}
                  </span>
                  <span style={{ fontSize: 12, color: "#8a8a9a" }}>{hostEntries.length} laporan</span>
                </button>
                {isOpen && (
                  <div style={{ padding: "0 14px 14px" }}>
                    {hostEntries.map((e) => {
                      const tercapai = e.gmv >= 2_000_000;
                      return (
                        <div key={e.id} style={{ borderTop: "1px solid #24242f", paddingTop: 10, marginTop: 10 }}>
                          <div style={{ display: "flex", justifyContent: "space-between" }}>
                            <div style={{ fontSize: 13, color: "#8a8a9a" }}>{formatTanggalDMY(e.tanggal)} · {e.jamMulai} · {e.sesi}</div>
                            {isAdmin && (
                              <button onClick={() => deleteEntry(e.id)} style={{ background: "none", border: "none", color: "#8a8a9a", cursor: "pointer" }}>
                                <Trash2 size={14} />
                              </button>
                            )}
                          </div>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 6, flexWrap: "wrap" }}>
                            <span style={{ fontFamily: "'JetBrains Mono', monospace", fontWeight: 700 }}>{formatRupiah(e.gmv)}</span>
                            <span style={{ fontSize: 11, padding: "2px 8px", borderRadius: 20, background: tercapai ? "rgba(37,244,238,0.15)" : "rgba(255,184,0,0.15)", color: tercapai ? GOOD : BAD, fontWeight: 600 }}>
                              {tercapai ? "🎉 Tercapai" : "Belum"}
                            </span>
                          </div>
                          {isAdmin && (
                            <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12, color: "#8a8a9a", marginTop: 4 }}>
                              Gaji {formatRupiah(BASE_UPAH)} + Bonus {formatRupiah(hitungBonus(e.gmv))} = <span style={{ color: GOOD }}>{formatRupiah(hitungTotalGaji(e.gmv))}</span>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {showHostSettings && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", zIndex: 10 }}>
          <div style={{ background: "#1c1c26", width: "100%", borderRadius: "16px 16px 0 0", padding: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <h3 style={{ fontFamily: "Fredoka, sans-serif", fontSize: 18, margin: 0 }}>Kelola Host</h3>
              <button onClick={() => setShowHostSettings(false)} style={{ background: "none", border: "none", color: "#F5F1E8", cursor: "pointer" }}>
                <X size={20} />
              </button>
            </div>
            {hosts.map((h) => (
              <div key={h} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid #2a2a36" }}>
                <span>{h}</span>
                <button onClick={() => removeHost(h)} style={{ background: "none", border: "none", color: BAD, cursor: "pointer" }}>
                  <Trash2 size={16} />
                </button>
              </div>
            ))}
            <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
              <input value={newHostName} onChange={(e) => setNewHostName(e.target.value)} placeholder="Nama host baru"
                style={{ ...selectStyle, margin: 0, flex: 1 }} />
              <button onClick={addHost} style={{ background: ACCENT, border: "none", borderRadius: 10, padding: "0 16px", color: "#14141c", fontWeight: 700, cursor: "pointer" }}>
                Tambah
              </button>
            </div>

            <div style={{ marginTop: 20, paddingTop: 16, borderTop: "1px solid #2a2a36" }}>
              <label style={labelStyle}>GANTI PIN ADMIN</label>
              <div style={{ display: "flex", gap: 8 }}>
                <input value={newPin} onChange={(e) => setNewPin(e.target.value)} placeholder="PIN baru" type="text"
                  style={{ ...selectStyle, margin: 0, flex: 1 }} />
                <button onClick={changePin} style={{ background: GOOD, border: "none", borderRadius: 10, padding: "0 16px", color: "#14141c", fontWeight: 700, cursor: "pointer" }}>
                  Simpan
                </button>
              </div>
            </div>

            <div style={{ marginTop: 20, paddingTop: 16, borderTop: "1px solid #24242f" }}>
              <label style={labelStyle}>IMPORT DARI FILE EXCEL KAMU</label>
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleExcelFile} style={{ display: "none" }} />
              <button onClick={() => fileInputRef.current?.click()}
                style={{ width: "100%", padding: "12px 0", borderRadius: 10, border: "1px solid #24242f", background: "#16161f", color: "#F5F1E8", fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
                <FileDown size={16} style={{ transform: "rotate(180deg)" }} /> Pilih File Excel (.xlsx)
              </button>
              <div style={{ fontSize: 11, color: "#8a8a9a", marginTop: 6 }}>
                Kolom yang dikenali otomatis: Nama Host, Tanggal, Jam Mulai, Sesi, GMV, Target — nama kolom fleksibel, tidak harus persis sama.
              </div>
              {excelMsg && <div style={{ marginTop: 8, fontSize: 12, color: "#8a8a9a" }}>{excelMsg}</div>}
            </div>
          </div>
        </div>
      )}

      {showPinPrompt && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 20 }}>
          <div style={{ background: "#1c1c26", borderRadius: 16, padding: 24, width: "85%", maxWidth: 320 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <h3 style={{ fontFamily: "Fredoka, sans-serif", fontSize: 18, margin: 0 }}>Mode Admin</h3>
              <button onClick={() => { setShowPinPrompt(false); setPinInput(""); setPinError(false); }} style={{ background: "none", border: "none", color: "#F5F1E8", cursor: "pointer" }}>
                <X size={20} />
              </button>
            </div>
            <input
              type="password"
              value={pinInput}
              onChange={(e) => { setPinInput(e.target.value); setPinError(false); }}
              placeholder="Masukkan PIN"
              style={selectStyle}
              onKeyDown={(e) => e.key === "Enter" && tryUnlock()}
            />
            {pinError && <div style={{ color: BAD, fontSize: 12, marginTop: 6 }}>PIN salah, coba lagi.</div>}
            <button onClick={tryUnlock}
              style={{ width: "100%", marginTop: 14, padding: "12px 0", borderRadius: 10, border: "none", background: ACCENT, color: "#14141c", fontWeight: 700, cursor: "pointer" }}>
              Buka
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

const labelStyle = { display: "block", fontSize: 11, letterSpacing: 1, color: "#8a8a9a", margin: "16px 0 6px", fontFamily: "'JetBrains Mono', monospace" };
const selectStyle = { width: "100%", padding: "12px 14px", borderRadius: 12, border: "1px solid #24242f", background: "#16161f", color: "#F5F1E8", fontSize: 14, fontFamily: "Inter", outline: "none", boxSizing: "border-box" };
const cardStyle = { flex: 1, background: "#16161f", border: "1px solid #24242f", borderRadius: 14, padding: 14 };
