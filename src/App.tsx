/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useRef, type ChangeEvent } from 'react';
import * as XLSX from 'xlsx';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// --- UTILS ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- CONSTANTS ---
const KABUPATEN_KOTA = [
  "Provinsi Sumatera Barat",
  "Kota Padang", "Kota Bukittinggi", "Kota Payakumbuh", "Kota Padang Panjang",
  "Kota Solok", "Kota Pariaman", "Kota Sawahlunto",
  "Kabupaten Agam", "Kabupaten Tanah Datar", "Kabupaten Padang Pariaman",
  "Kabupaten Pesisir Selatan", "Kabupaten Solok", "Kabupaten Solok Selatan",
  "Kabupaten Sijunjung", "Kabupaten Dharmasraya", "Kabupaten Lima Puluh Kota",
  "Kabupaten Pasaman", "Kabupaten Pasaman Barat", "Kabupaten Kepulauan Mentawai",
];

const KATEGORI = [
  "Festival Budaya", "Olahraga & Adventure", "Kuliner", "Seni & Pertunjukan",
  "Religi & Tradisi", "MICE", "Alam & Ekowisata", "Pameran & Expo",
];

const BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

const STATUS_BADGE: Record<string, { bg: string, color: string, border: string }> = {
  "Draft": { bg: "#1A1A2A", color: "#8A8AA0", border: "#2A2A4A" },
  "Diajukan": { bg: "#0A2218", color: "#4CAF82", border: "#0E4028" },
};

const ACCOUNTS = [
  { username: "admin.provinsi", password: "sumbarrancak", role: "provinsi", kabupatenKota: "Provinsi Sumatera Barat", nama: "Admin Dinas Provinsi Sumbar" },
  ...KABUPATEN_KOTA.map((k) => ({
    username: "admin." + k.toLowerCase().replace(/\s+/g, ".").replace(/\//g, ""),
    password: "sumbarrancak",
    role: "kabkota" as const,
    kabupatenKota: k,
    nama: `Admin Dispar ${k}`,
  })),
];

interface Event {
  id: number;
  namaEvent: string;
  kabupatenKota: string;
  kategori: string;
  tanggalMulai: string;
  tanggalSelesai?: string;
  lokasi?: string;
  deskripsi?: string;
  targetWisatawan?: string;
  kontakNama?: string;
  kontakHP?: string;
  kontakEmail?: string;
  anggaran?: string;
  status: "Draft" | "Diajukan";
  createdAt?: string;
}

const initialForm: Partial<Event> = {
  namaEvent: "", kabupatenKota: "", kategori: "", tanggalMulai: "",
  tanggalSelesai: "", lokasi: "", deskripsi: "", targetWisatawan: "",
  kontakNama: "", kontakHP: "", kontakEmail: "", anggaran: "", status: "Draft",
};

const STORAGE_KEY = "coe_sumbar_v2_events";
const SESSION_KEY = "coe_sumbar_session";
const SHEET_URL = import.meta.env.VITE_SHEET_URL || "";

// --- HELPERS ---
function formatDate(dateStr: string | undefined) {
  if (!dateStr) return "-";
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr; // Return as is if invalid
    return d.toLocaleDateString("id-ID", {
      day: "numeric",
      month: "long",
      year: "numeric",
    });
  } catch {
    return dateStr;
  }
}

async function sheetSaveEvent(event: Event) {
  try {
    // Note: Google Apps Script usually requires redirection and correct CORS Handling.
    // In many cases, sending as 'no-cors' doesn't send the body correctly if Content-Type is application/json.
    // However, if the user's script expects this, we'll try to refine the request.
    await fetch(SHEET_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain" }, // Google Scripts work better with text/plain to avoid pre-flight
      body: JSON.stringify({ action: "save", event }),
    });
  } catch (e) {
    console.warn("Sheet save failed:", e);
  }
}

async function sheetDeleteEvent(id: number) {
  try {
    await fetch(SHEET_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain" },
      body: JSON.stringify({ action: "delete", id }),
    });
  } catch (e) {
    console.warn("Sheet delete failed:", e);
  }
}

async function sheetGetAll(): Promise<Event[] | null> {
  try {
    const res = await fetch(SHEET_URL + "?action=getAll");
    const data = await res.json();
    return data.events || null;
  } catch (e) {
    console.warn("Sheet getAll failed:", e);
    return null;
  }
}

function getSeedData(): Event[] {
  return [
    {
      id: 1, namaEvent: "Pesona Budaya Hoyak Tabuik Piaman", kabupatenKota: "Kota Pariaman",
      kategori: "Festival Budaya", tanggalMulai: "2025-07-05", tanggalSelesai: "2025-07-14",
      lokasi: "Pantai Gandoriah, Pariaman", deskripsi: "Festival tradisi tabuik yang digelar setiap Muharram.",
      targetWisatawan: "5000", kontakNama: "Dedi Rahmat", kontakHP: "081234567890",
      kontakEmail: "pariwisata@pariaman.go.id", anggaran: "500000000", status: "Diajukan", createdAt: "2025-01-10"
    },
    {
      id: 2, namaEvent: "Tour de Singkarak", kabupatenKota: "Kabupaten Solok",
      kategori: "Olahraga & Adventure", tanggalMulai: "2025-10-15", tanggalSelesai: "2025-10-22",
      lokasi: "Danau Singkarak & sekitarnya", deskripsi: "Balap sepeda internasional mengelilingi alam Sumatera Barat.",
      targetWisatawan: "20000", kontakNama: "Roni Amir", kontakHP: "082345678901",
      kontakEmail: "pariwisata@solok.go.id", anggaran: "2000000000", status: "Diajukan", createdAt: "2025-02-01"
    },
    {
      id: 3, namaEvent: "Festival Rendang Dunia", kabupatenKota: "Kota Padang",
      kategori: "Kuliner", tanggalMulai: "2025-09-20", tanggalSelesai: "2025-09-22",
      lokasi: "GOR H. Agus Salim, Padang", deskripsi: "Festival kuliner rendang sebagai warisan budaya UNESCO.",
      targetWisatawan: "10000", kontakNama: "Sari Dewi", kontakHP: "083456789012",
      kontakEmail: "pariwisata@padang.go.id", anggaran: "750000000", status: "Draft", createdAt: "2025-03-15"
    },
  ];
}

// --- APP COMPONENT ---
export default function App() {
  const [user, setUser] = useState<{ role: string, kabupatenKota: string, nama: string, username: string } | null>(() => {
    try {
      const r = sessionStorage.getItem(SESSION_KEY);
      return r ? JSON.parse(r) : null;
    } catch { return null; }
  });

  const login = (u: any) => {
    sessionStorage.setItem(SESSION_KEY, JSON.stringify(u));
    setUser(u);
  };
  const logout = () => {
    sessionStorage.removeItem(SESSION_KEY);
    setUser(null);
  };

  if (!user) return <LoginScreen onLogin={login} />;
  return <MainApp user={user} onLogout={logout} />;
}

// --- LOGIN SCREEN ---
function LoginScreen({ onLogin }: { onLogin: (u: any) => void }) {
  const [uname, setUname] = useState("");
  const [pass, setPass] = useState("");
  const [err, setErr] = useState("");
  const [show, setShow] = useState(false);
  const [hint, setHint] = useState(false);

  const doLogin = () => {
    const acc = ACCOUNTS.find(a => a.username === uname.trim() && a.password === pass);
    if (!acc) { setErr("Username atau password salah."); return; }
    onLogin(acc);
  };

  return (
    <div className="min-h-screen flex items-center justify-center font-serif" 
      style={{ backgroundImage: `radial-gradient(ellipse at 30% 40%,rgba(196,160,60,.09) 0%,transparent 55%), radial-gradient(ellipse at 75% 70%,rgba(34,85,34,.18) 0%,transparent 50%)` }}>
      <div className="w-[380px] bg-[rgba(8,20,10,0.97)] border border-[rgba(196,160,60,0.2)] rounded-2xl p-9 shadow-[0_24px_60px_rgba(0,0,0,0.6)]">
        <div className="text-center mb-8">
          <div className="text-4xl mb-2.5">🏔</div>
          <div className="text-[9px] tracking-[3px] text-[#C4A03C] uppercase mb-1">Dinas Pariwisata</div>
          <div className="text-xl font-bold text-[#E8DCC8]">Sumatera Barat</div>
          <div className="text-[10px] text-[#6A5830] mt-1 tracking-widest">Calendar of Events · Sistem Pendataan</div>
        </div>

        <div className="mb-4">
          <label className="text-[11px] text-[#6A5830] block mb-1.5">Username</label>
          <input value={uname} onChange={e => { setUname(e.target.value); setErr(""); }}
            onKeyDown={e => e.key === "Enter" && doLogin()}
            placeholder="admin.provinsi / admin.kota.padang"
            className="input-field w-full" />
        </div>

        <div className="mb-2">
          <label className="text-[11px] text-[#6A5830] block mb-1.5">Password</label>
          <div className="relative">
            <input value={pass} onChange={e => { setPass(e.target.value); setErr(""); }}
              onKeyDown={e => e.key === "Enter" && doLogin()}
              type={show ? "text" : "password"} placeholder="••••••••"
              className="input-field w-full pr-10" />
            <button onClick={() => setShow(s => !s)}
              className="absolute right-2.5 top-1/2 -translate-y-1/2 bg-none border-none text-[#6A5830] cursor-pointer text-sm p-0">
              {show ? "🙈" : "👁"}
            </button>
          </div>
        </div>

        {err && <div className="text-[11px] text-[#E05A5A] mb-2.5 px-2.5 py-1.5 bg-[rgba(224,90,90,0.08)] rounded-md">{err}</div>}

        <button onClick={doLogin} className="btn-primary w-full py-2.5 mt-2">
          Masuk
        </button>

        <div className="mt-5 text-center">
          <button onClick={() => setHint(h => !h)}
            className="text-[10px] text-[#4A3820] bg-none border-none cursor-pointer underline">
            {hint ? "Sembunyikan" : "Lihat semua akun demo"}
          </button>
          {hint && (
            <div className="mt-2.5 bg-[rgba(255,255,255,0.02)] border border-[rgba(196,160,60,0.1)] rounded-lg p-3 text-left max-h-[260px] overflow-y-auto">
              <div className="text-[10px] text-[#6A5830] mb-1 font-bold">ADMIN PROVINSI</div>
              <div className="text-[10px] text-[#8A7860] font-mono mb-2.5 bg-[rgba(196,160,60,0.06)] px-2 py-1 rounded cursor-pointer"
                onClick={() => { setUname("admin.provinsi"); setPass("sumbarrancak"); setHint(false); }}>
                admin.provinsi &nbsp;·&nbsp; sumbarrancak
              </div>
              <div className="text-[10px] text-[#6A5830] mb-1 font-bold">
                ADMIN KAB/KOTA <span className="text-[#3A3020] font-normal">(password: sumbarrancak)</span>
              </div>
              {ACCOUNTS.filter(a => a.role === "kabkota").map(a => (
                <div key={a.username}
                  className="text-[10px] text-[#8A7860] font-mono px-2 py-1 rounded mb-0.5 bg-[rgba(255,255,255,0.02)] cursor-pointer hover:bg-[rgba(196,160,60,0.05)]"
                  onClick={() => { setUname(a.username); setPass("sumbarrancak"); setHint(false); }}>
                  {a.username}
                </div>
              ))}
              <div className="text-[9px] text-[#3A3020] mt-2 italic">✱ klik username untuk mengisi otomatis</div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// --- MAIN APP COMPONENT ---
function MainApp({ user, onLogout }: { user: any, onLogout: () => void }) {
  const isProvinsi = user.role === "provinsi";

  const [events, setEvents] = useState<Event[]>([]);
  const [view, setView] = useState("dashboard");
  const [form, setForm] = useState<Partial<Event>>(initialForm);
  const [editId, setEditId] = useState<number | null>(null);
  const [selEvent, setSelEvent] = useState<Event | null>(null);
  const [filterKab, setFilterKab] = useState("");
  const [filterKat, setFilterKat] = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const [filterBulan, setFilterBulan] = useState("");
  const [searchQ, setSearchQ] = useState("");
  const [calYear, setCalYear] = useState(2025);
  const [calMonth, setCalMonth] = useState(new Date().getMonth());
  const [toast, setToast] = useState<{ msg: string, type: string } | null>(null);
  const [delConfirm, setDelConfirm] = useState<number | null>(null);
  const [ajukanConfirm, setAjukanConfirm] = useState<number | null>(null);
  const [importConfirm, setImportConfirm] = useState<{ events: Event[], count: number } | null>(null);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiResult, setAiResult] = useState<string | null>(null);
  const importRef = useRef<HTMLInputElement>(null);

  const [sheetLoading, setSheetLoading] = useState(true);

  // Initial Load
  useEffect(() => {
    loadDataFromServer();
    
    // Auto-refresh every 10 seconds for real-time visibility
    const interval = setInterval(() => {
      loadDataFromServer(false);
    }, 10000);
    
    return () => clearInterval(interval);
  }, []);

  const loadDataFromServer = async (showLoading = true) => {
    if (showLoading) setSheetLoading(true);
    const sheetEvents = await sheetGetAll();
    if (sheetEvents && sheetEvents.length > 0) {
      setEvents(sheetEvents);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(sheetEvents));
    } else {
      const local = localStorage.getItem(STORAGE_KEY);
      if (local) setEvents(JSON.parse(local));
      else setEvents(getSeedData());
    }
    setSheetLoading(false);
  };

  const showToast = (msg: string, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  };

  const handleRefresh = async () => {
    setSheetLoading(true);
    await loadDataFromServer();
    showToast("Data diperbarui dari server! ✓");
  };

  const canEditEvent = (e: Event) => isProvinsi || (e.kabupatenKota === user.kabupatenKota && e.status === "Draft");

  const handleSubmit = async () => {
    const finalForm = !isProvinsi
      ? { ...form, kabupatenKota: user.kabupatenKota, status: form.status || "Draft" }
      : { ...form };

    if (!finalForm.namaEvent || !finalForm.kabupatenKota || !finalForm.tanggalMulai || !finalForm.kategori) {
      showToast("Harap isi semua field wajib (*)", "error");
      return;
    }

    if (editId) {
      const existing = events.find(e => e.id === editId);
      const updated = { ...finalForm, id: editId, createdAt: existing?.createdAt } as Event;
      setEvents(ev => ev.map(e => e.id === editId ? updated : e));
      await sheetSaveEvent(updated);
      showToast("Event berhasil diperbarui! ✓");
    } else {
      const newEvent = { ...finalForm, id: Date.now(), createdAt: new Date().toISOString().split("T")[0] } as Event;
      setEvents(ev => [...ev, newEvent]);
      await sheetSaveEvent(newEvent);
      showToast("Event baru berhasil disimpan! ✓");
    }
    setForm(initialForm); setEditId(null); setView("list");
  };

  const handleEdit = (e: Event) => {
    if (!canEditEvent(e)) return;
    setForm({ ...e }); setEditId(e.id); setView("form");
  };

  const handleDelete = async (id: number) => {
    setEvents(ev => ev.filter(e => e.id !== id));
    await sheetDeleteEvent(id);
    setDelConfirm(null);
    showToast("Event berhasil dihapus.", "info");
    if (view === "detail") setView("list");
  };

  const handleAjukan = async (id: number) => {
    let updatedEvent: Event | null = null;
    setEvents(ev => ev.map(e => {
      if (e.id === id) {
        updatedEvent = { ...e, status: "Diajukan" as const };
        return updatedEvent;
      }
      return e;
    }));
    
    if (updatedEvent) {
      await sheetSaveEvent(updatedEvent);
      if (selEvent?.id === id) setSelEvent(updatedEvent);
    }
    
    setAjukanConfirm(null);
    showToast("Event telah diajukan ke Provinsi! ✓");
  };

  const handleImportFile = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const parsed = JSON.parse(ev.target.result as string);
        const imported = parsed.events || (Array.isArray(parsed) ? parsed : null);
        if (!imported || !imported.length) { showToast("File tidak valid.", "error"); return; }
        setImportConfirm({ events: imported, count: imported.length });
      } catch { showToast("Gagal membaca file JSON.", "error"); }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  const confirmImport = (mode: "replace" | "merge") => {
    if (!importConfirm) return;
    if (mode === "replace") {
      setEvents(importConfirm.events);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(importConfirm.events));
      showToast(`${importConfirm.count} event di-restore.`);
    } else {
      const merged = [...events, ...importConfirm.events.filter(ie => !events.find(e => e.id === ie.id))];
      setEvents(merged);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(merged));
      showToast(`${importConfirm.count} event digabungkan.`);
    }
    setImportConfirm(null);
  };

  const exportToExcel = () => {
    const rows = events.map((e, i) => ({
      "No": i + 1, "Nama Event": e.namaEvent, "Provinsi / Kab / Kota": e.kabupatenKota,
      "Kategori": e.kategori, "Tgl Mulai": e.tanggalMulai, "Tgl Selesai": e.tanggalSelesai || "-",
      "Lokasi": e.lokasi || "-", "Deskripsi": e.deskripsi || "-",
      "Target Wisatawan": e.targetWisatawan ? parseInt(e.targetWisatawan) : 0,
      "Estimasi Anggaran": e.anggaran ? parseInt(e.anggaran) : 0,
      "Kontak": e.kontakNama || "-", "HP": e.kontakHP || "-",
      "Status": e.status, "Tgl Input": e.createdAt || "-"
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Calendar of Events");
    XLSX.writeFile(wb, `CoE_SumBar_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  // Filter Logic
  const filtered = events.filter(e => {
    const q = searchQ.toLowerCase();
    return (!q || e.namaEvent.toLowerCase().includes(q) || e.lokasi?.toLowerCase().includes(q))
      && (!filterKab || e.kabupatenKota === filterKab)
      && (!filterKat || e.kategori === filterKat)
      && (!filterStatus || e.status === filterStatus)
      && (!filterBulan || (e.tanggalMulai && new Date(e.tanggalMulai).getMonth() === parseInt(filterBulan)));
  });

  const getDIM = (y: number, m: number) => new Date(y, m + 1, 0).getDate();
  const getFD = (y: number, m: number) => new Date(y, m, 1).getDay();
  const eInDay = (day: number) => events.filter(e => {
    if (!e.tanggalMulai) return false;
    const s = new Date(e.tanggalMulai), en = e.tanggalSelesai ? new Date(e.tanggalSelesai) : s;
    const d = new Date(calYear, calMonth, day); return d >= s && d <= en;
  });

  // Stats
  const totalEvt = events.length;
  const totalTarget = events.reduce((s, e) => s + parseInt(e.targetWisatawan || "0"), 0);
  const diajukanCount = events.filter(e => e.status === "Diajukan").length;

  return (
    <div className="min-h-screen bg-[#0D1B0F] text-[#E8DCC8] relative overflow-x-hidden font-serif">
      <div className="fixed inset-0 z-0 pointer-events-none" style={{ backgroundImage: `radial-gradient(ellipse at 20% 50%,rgba(196,160,60,.07) 0%,transparent 60%), radial-gradient(ellipse at 80% 20%,rgba(34,85,34,.14) 0%,transparent 50%)` }} />

      {/* SIDEBAR */}
      <div className="fixed left-0 top-0 bottom-0 w-[232px] bg-[rgba(7,18,9,0.98)] border-r border-[rgba(196,160,60,0.15)] flex flex-col z-50">
        <div className="p-6 pb-4 border-b border-[rgba(196,160,60,0.1)]">
          <div className="text-[9px] tracking-[3px] text-[#C4A03C] uppercase mb-1">Dinas Pariwisata</div>
          <div className="text-base font-bold text-[#E8DCC8] leading-tight">Sumatera Barat</div>
          <div className="text-[9px] text-[#5A4820] mt-1 tracking-widest">Calendar of Events</div>
        </div>

        <div className="mx-3 mt-3 p-2.5 bg-[rgba(196,160,60,0.08)] border border-[rgba(196,160,60,0.2)] rounded-lg">
          <div className="text-[9px] text-[#C4A03C] tracking-widest uppercase mb-1">{isProvinsi ? '🏛 Admin Provinsi' : '🏘 Admin Kab/Kota'}</div>
          <div className="text-[11px] text-[#C8B890] font-semibold leading-tight">{user.nama}</div>
          {!isProvinsi && <div className="text-[10px] text-[#5A5040] mt-0.5">{user.kabupatenKota}</div>}
        </div>

        <nav className="p-2.5 flex flex-col gap-0.5 mt-2">
          <MenuBtn active={view === "dashboard"} icon="◈" label="Dashboard" onClick={() => setView("dashboard")} />
          <MenuBtn active={view === "form"} icon="✦" label="Tambah Event" onClick={() => { setForm(initialForm); setEditId(null); setView("form"); }} />
          <MenuBtn active={view === "list"} icon="≡" label="Daftar Event" onClick={() => setView("list")} />
          <MenuBtn active={view === "calendar"} icon="⊞" label="Kalender" onClick={() => setView("calendar")} />
        </nav>

        {isProvinsi && (
          <div className="px-2.5 pt-1 border-t border-[rgba(196,160,60,0.08)] mt-auto mb-20">
            <div className="text-[9px] tracking-[2px] text-[#2A2015] uppercase px-3 py-2 font-bold">Admin Tools</div>
            <SidebarAction icon="↻" label="Sync Server" onClick={handleRefresh} color="#4CAF82" />
            <SidebarAction icon="📊" label="Export Excel" onClick={exportToExcel} color="#4CAF82" />
            <SidebarAction icon="💾" label="Backup JSON" onClick={() => { 
                const blob = new Blob([JSON.stringify({events}, null, 2)], {type: 'application/json'});
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a'); a.href = url; a.download = 'backup_coe.json'; a.click();
              }} color="#7EB8F7" />
            <SidebarAction icon="📂" label="Import JSON" onClick={() => importRef.current?.click()} color="#E8A820" />
            <input ref={importRef} type="file" accept=".json" onChange={handleImportFile} className="hidden" />
          </div>
        )}

        <div className="mt-auto p-4 border-t border-[rgba(196,160,60,0.08)]">
          <button onClick={onLogout} className="flex items-center gap-2 w-full p-2.5 rounded-lg border border-[rgba(224,90,90,0.2)] bg-[rgba(224,90,90,0.05)] text-[#A05050] text-[11px] cursor-pointer hover:bg-[rgba(224,90,90,0.1)] transition-colors">
            <span>⬡</span> Keluar
          </button>
        </div>
      </div>

      {/* MAIN CONTENT */}
      <div className="ml-[232px] min-h-screen relative z-10">
        <Header title={view === "dashboard" ? "Dashboard Ikhtisar" : view === "form" ? (editId ? "Edit Event" : "Tambah Event") : view === "list" ? "Daftar Event" : "Kalender"} 
          onAdd={() => { setForm(initialForm); setEditId(null); setView("form"); }} 
          loading={sheetLoading}
          onRefresh={handleRefresh}
        />

        <main className="p-8">
          {view === "dashboard" && (
            <div className="animate-in fade-in duration-500">
              <div className="grid grid-cols-3 gap-4 mb-6">
                <StatCard label="Total Event" value={totalEvt} icon="✦" color="#C4A03C" />
                <StatCard label="Sudah Diajukan" value={diajukanCount} icon="✉" color="#4CAF82" />
                <StatCard label="Target Wisatawan" value={totalTarget.toLocaleString("id-ID")} icon="⊙" color="#7EB8F7" />
              </div>

              <div className="grid grid-cols-2 gap-5 mb-5">
                <div className="bg-white/5 border border-white/10 rounded-xl p-5">
                  <div className="text-[10px] text-[#C4A03C] tracking-widest uppercase mb-4">Event Terbaru</div>
                  {events.slice(-6).reverse().map(e => (
                    <div key={e.id} onClick={() => { setSelEvent(e); setView("detail"); }} className="py-2.5 border-b border-white/5 cursor-pointer hover:bg-white/5 transition-colors px-2 rounded-md">
                      <div className="flex justify-between items-center">
                        <span className="text-[13px] text-[#E8DCC8] font-semibold">{e.namaEvent}</span>
                        <StatusPill status={e.status} />
                      </div>
                      <div className="text-[10px] text-[#4A4030] mt-1">{e.kabupatenKota} · {e.tanggalMulai}</div>
                    </div>
                  ))}
                </div>
                <div className="bg-white/5 border border-white/10 rounded-xl p-5">
                  <div className="text-[10px] text-[#C4A03C] tracking-widest uppercase mb-4">Distribusi Kategori</div>
                  {KATEGORI.map(k => {
                    const count = events.filter(e => e.kategori === k).length;
                    const pct = totalEvt ? Math.round(count / totalEvt * 100) : 0;
                    return count > 0 ? (
                      <div key={k} className="mb-2.5">
                        <div className="flex justify-between text-[11px] text-[#7A6E5A] mb-1">
                          <span>{k}</span><span>{count}</span>
                        </div>
                        <div className="h-1 bg-white/5 rounded-full overflow-hidden">
                          <div className="h-full bg-gradient-to-r from-[#C4A03C] to-[#7EB8F7]" style={{ width: `${pct}%` }} />
                        </div>
                      </div>
                    ) : null;
                  })}
                </div>
              </div>
            </div>
          )}

          {view === "form" && (
            <div className="max-w-2xl animate-in slide-in-from-bottom-2 duration-300">
              <div className="bg-white/5 border border-[rgba(196,160,60,0.1)] rounded-xl p-6 shadow-xl">
                <h3 className="text-[#C4A03C] text-[10px] tracking-widest uppercase border-b border-[rgba(196,160,60,0.1)] pb-2 mb-6">📝 {editId ? 'Ubah' : 'Input'} Data Event</h3>
                
                <div className="grid grid-cols-2 gap-4">
                  <div className="col-span-2">
                    <label className="field-label">Nama Event *</label>
                    <input className="input-field w-full" value={form.namaEvent} onChange={e => setForm({...form, namaEvent: e.target.value})} />
                  </div>
                  <div>
                    <label className="field-label">Wilayah *</label>
                    <select className="input-field w-full" value={isProvinsi ? form.kabupatenKota : user.kabupatenKota} 
                      disabled={!isProvinsi} onChange={e => setForm({...form, kabupatenKota: e.target.value})}>
                      <option value="">Pilih Wilayah</option>
                      {KABUPATEN_KOTA.map(k => <option key={k} value={k}>{k}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="field-label">Kategori *</label>
                    <select className="input-field w-full" value={form.kategori} onChange={e => setForm({...form, kategori: e.target.value})}>
                      <option value="">Pilih Kategori</option>
                      {KATEGORI.map(k => <option key={k} value={k}>{k}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="field-label">Tgl Mulai *</label>
                    <input type="date" className="input-field w-full" value={form.tanggalMulai} onChange={e => setForm({...form, tanggalMulai: e.target.value})} />
                  </div>
                  <div>
                    <label className="field-label">Tgl Selesai</label>
                    <input type="date" className="input-field w-full" value={form.tanggalSelesai} onChange={e => setForm({...form, tanggalSelesai: e.target.value})} />
                  </div>
                  <div className="col-span-2">
                    <label className="field-label">Lokasi</label>
                    <input className="input-field w-full" value={form.lokasi} onChange={e => setForm({...form, lokasi: e.target.value})} />
                  </div>
                  <div className="col-span-2">
                    <label className="field-label">Deskripsi</label>
                    <textarea className="input-field w-full h-20 resize-none" value={form.deskripsi} onChange={e => setForm({...form, deskripsi: e.target.value})} />
                  </div>
                  <div>
                    <label className="field-label">Target Wisatawan</label>
                    <input type="number" className="input-field w-full" value={form.targetWisatawan} onChange={e => setForm({...form, targetWisatawan: e.target.value})} />
                  </div>
                  <div>
                    <label className="field-label">Estimasi Anggaran (Rp)</label>
                    <input type="number" className="input-field w-full" value={form.anggaran} onChange={e => setForm({...form, anggaran: e.target.value})} />
                  </div>
                </div>

                <div className="flex gap-3 mt-8">
                  <button onClick={handleSubmit} className="btn-primary px-8">Simpan</button>
                  <button onClick={() => setView("list")} className="px-6 py-2 rounded-lg border border-[rgba(196,160,60,0.2)] text-[#7A6E5A] text-sm">Batal</button>
                </div>
              </div>
            </div>
          )}

          {view === "list" && (
            <div className="animate-in fade-in duration-500">
              <div className="flex flex-wrap gap-2.5 mb-5 p-4 bg-white/5 border border-white/5 rounded-xl">
                <input className="input-field flex-grow" placeholder="🔍 Cari event..." value={searchQ} onChange={e => setSearchQ(e.target.value)} />
                <select className="input-field min-w-[150px]" value={filterKab} onChange={e => setFilterKab(e.target.value)}>
                  <option value="">Semua Wilayah</option>
                  {KABUPATEN_KOTA.map(k => <option key={k} value={k}>{k}</option>)}
                </select>
                <select className="input-field min-w-[150px]" value={filterKat} onChange={e => setFilterKat(e.target.value)}>
                  <option value="">Semua Kategori</option>
                  {KATEGORI.map(k => <option key={k} value={k}>{k}</option>)}
                </select>
              </div>

              {filtered.map(e => (
                <div key={e.id} onClick={() => { setSelEvent(e); setView("detail"); }} className="bg-white/5 border border-white/10 rounded-xl p-4 mb-2.5 flex items-center gap-4 cursor-pointer hover:border-[#C4A03C40] transition-all">
                  <div className="w-10 h-10 rounded-lg bg-[#C4A03C10] flex items-center justify-center text-lg shrink-0">
                    {e.kategori.includes("Budaya") ? "🎭" : e.kategori.includes("Olahraga") ? "🏆" : "✦"}
                  </div>
                  <div className="flex-grow min-w-0">
                    <div className="flex items-center gap-2 mb-0.5">
                      <span className="font-bold text-sm text-[#E8DCC8] truncate">{e.namaEvent}</span>
                      <StatusPill status={e.status} />
                    </div>
                    <div className="text-[11px] text-[#6A5840]">{e.kabupatenKota} · {e.kategori} · {e.tanggalMulai}</div>
                  </div>
                  <div className="flex gap-1.5" onClick={ev => ev.stopPropagation()}>
                    {canEditEvent(e) && (
                      <>
                        <button onClick={() => handleEdit(e)} className="p-2 text-[#C4A03C] hover:bg-white/5 rounded-md">✎</button>
                        <button onClick={() => setDelConfirm(e.id)} className="p-2 text-[#E05A5A] hover:bg-white/5 rounded-md">✕</button>
                      </>
                    )}
                  </div>
                </div>
              ))}
            </div>
          )}

          {view === "calendar" && (
            <div className="animate-in fade-in duration-500">
              <div className="flex items-center gap-4 mb-6">
                <button onClick={() => { if (calMonth === 0) { setCalMonth(11); setCalYear(y => y - 1); } else setCalMonth(m => m - 1); }}
                  className="btn-secondary !py-1 !px-3 hover:bg-[#C4A03C15]">‹</button>
                <div className="text-lg font-bold text-[#E8DCC8] min-w-[200px] text-center">
                  {BULAN[calMonth]} {calYear}
                </div>
                <button onClick={() => { if (calMonth === 11) { setCalMonth(0); setCalYear(y => y + 1); } else setCalMonth(m => m + 1); }}
                  className="btn-secondary !py-1 !px-3 hover:bg-[#C4A03C15]">›</button>
              </div>
              <div className="grid grid-cols-7 gap-1 mb-1">
                {["Min", "Sen", "Sel", "Rab", "Kam", "Jum", "Sab"].map(d => (
                  <div key={d} className="text-center text-[11px] text-[#4A4030] py-1 font-bold">{d}</div>
                ))}
              </div>
              <div className="grid grid-cols-7 gap-1">
                {Array(getFD(calYear, calMonth)).fill(null).map((_, i) => <div key={`empty-${i}`} className="min-h-[100px]" />)}
                {Array(getDIM(calYear, calMonth)).fill(null).map((_, i) => {
                  const day = i + 1;
                  const dayEvs = eInDay(day);
                  const isToday = new Date().getFullYear() === calYear && new Date().getMonth() === calMonth && new Date().getDate() === day;
                  return (
                    <div key={day} className={cn("min-h-[100px] rounded-lg p-2 transition-colors border", 
                      isToday ? "bg-[#C4A03C10] border-[#C4A03C40]" : "bg-white/5 border-white/5")}>
                      <div className={cn("text-sm mb-1 font-bold", isToday ? "text-[#C4A03C]" : "text-[#4A4030]")}>{day}</div>
                      <div className="flex flex-col gap-1">
                        {dayEvs.map(e => (
                          <div key={e.id} onClick={(ev) => { ev.stopPropagation(); setSelEvent(e); setView("detail"); }}
                            className="text-[8px] bg-[#C4A03C15] text-[#C4A03C] p-1 rounded border border-[#C4A03C20] truncate cursor-pointer hover:bg-[#C4A03C25] transition-colors"
                            title={e.namaEvent}>
                            {e.namaEvent}
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {view === "detail" && selEvent && (
            <div className="max-w-2xl animate-in fade-in zoom-in-95 duration-300">
               <button onClick={() => setView("list")} className="text-[11px] text-[#7A6E5A] mb-4 hover:text-[#C4A03C]">← Kembali ke daftar</button>
               <div className="bg-white/5 border border-white/10 rounded-xl p-6">
                  <div className="flex justify-between items-start mb-6">
                    <div>
                      <h2 className="text-xl font-bold text-[#E8DCC8] mb-1.5">{selEvent.namaEvent}</h2>
                      <StatusPill status={selEvent.status} />
                    </div>
                    {canEditEvent(selEvent) && selEvent.status === "Draft" && (
                      <button onClick={() => setAjukanConfirm(selEvent.id)} className="btn-primary text-[11px] px-4 py-2">✉ Ajukan ke Provinsi</button>
                    )}
                  </div>
                  <dl className="grid grid-cols-[140px_1fr] gap-x-4 gap-y-3.5 text-xs">
                    <dt className="text-[#4A4030]">Wilayah</dt><dd>{selEvent.kabupatenKota}</dd>
                    <dt className="text-[#4A4030]">Kategori</dt><dd>{selEvent.kategori}</dd>
                    <dt className="text-[#4A4030]">Tanggal</dt><dd>{formatDate(selEvent.tanggalMulai)} {selEvent.tanggalSelesai ? `s/d ${formatDate(selEvent.tanggalSelesai)}` : ''}</dd>
                    <dt className="text-[#4A4030]">Lokasi</dt><dd>{selEvent.lokasi || '-'}</dd>
                    <dt className="text-[#4A4030]">Deskripsi</dt><dd className="leading-relaxed opacity-80">{selEvent.deskripsi || '-'}</dd>
                    <dt className="text-[#4A4030]">Sasaran</dt><dd>{selEvent.targetWisatawan ? Number(selEvent.targetWisatawan).toLocaleString('id-ID') + ' orang' : '-'}</dd>
                    <dt className="text-[#4A4030]">Anggaran</dt><dd>Rp {Number(selEvent.anggaran || 0).toLocaleString('id-ID')}</dd>
                  </dl>
               </div>
            </div>
          )}
        </main>
      </div>

      {/* TOASTS & MODALS */}
      {toast && (
        <div className={cn("fixed bottom-6 right-6 z-[200] p-4 rounded-lg shadow-2xl border text-xs min-w-[200px]", 
          toast.type === "error" ? "bg-[#2A0808] border-[#E05A5A] text-[#FFD0D0]" : "bg-[#082214] border-[#4CAF82] text-[#D0FFD0]")}>
          {toast.msg}
        </div>
      )}

      {delConfirm && <ConfirmModal title="Hapus Event?" desc="Aksi ini tidak dapat dibatalkan." color="#E05A5A" onConfirm={() => handleDelete(delConfirm)} onCancel={() => setDelConfirm(null)} />}
      {ajukanConfirm && <ConfirmModal title="Ajukan Event?" desc="Kirim data ke Dinas Provinsi? Status akan menjadi 'Diajukan' dan tidak bisa diedit lagi." color="#4CAF82" onConfirm={() => handleAjukan(ajukanConfirm)} onCancel={() => setAjukanConfirm(null)} icon="✉" />}
      {importConfirm && (
        <div className="fixed inset-0 z-[300] bg-black/80 flex items-center justify-center p-4">
          <div className="bg-[#0B1A0D] border border-white/10 p-8 rounded-2xl max-w-sm w-full text-center">
            <div className="text-4xl mb-4">📂</div>
            <h4 className="text-lg font-bold mb-2">Import Data</h4>
            <p className="text-xs text-[#7A6E5A] mb-6">{importConfirm.count} event siap di-import.</p>
            <div className="flex flex-col gap-2">
              <button onClick={() => confirmImport("merge")} className="btn-secondary w-full">Gabungkan Data</button>
              <button onClick={() => confirmImport("replace")} className="btn-primary w-full !bg-[#A03030]">Ganti Semua Data</button>
              <button onClick={() => setImportConfirm(null)} className="text-[11px] text-[#4A4030] mt-2">Batal</button>
            </div>
          </div>
        </div>
      )}

      <style>{`
        .input-field {
          background: rgba(255,255,255,0.04);
          border: 1px solid rgba(196,160,60,0.16);
          color: #E8DCC8;
          padding: 9px 12px;
          border-radius: 8px;
          font-size: 13px;
          outline: none;
          transition: border-color 0.2s;
        }
        .input-field:focus { border-color: rgba(196,160,60,0.4); }
        .btn-primary {
          background: linear-gradient(135deg,#C4A03C,#A07828);
          color: #0D1B0F;
          border: none;
          border-radius: 8px;
          font-size: 13px;
          font-weight: 700;
          cursor: pointer;
          transition: transform 0.1s;
          padding: 10px 16px;
        }
        .btn-primary:active { transform: scale(0.98); }
        .btn-secondary {
          background: rgba(196,160,60,0.08);
          border: 1px solid rgba(196,160,60,0.2);
          color: #C4A03C;
          padding: 10px 16px;
          border-radius: 8px;
          font-size: 13px;
          cursor: pointer;
        }
        .field-label {
          display: block;
          font-size: 11px;
          color: #4A4030;
          margin-bottom: 6px;
          font-family: inherit;
        }
      `}</style>
    </div>
  );
}

// --- SUB COMPONENTS ---

function Header({ title, onAdd, onRefresh, loading }: any) {
  return (
    <header className="sticky top-0 z-40 px-8 py-5 flex justify-between items-center bg-[#071209]/90 backdrop-blur-md border-b border-white/5">
      <div>
        <h1 className="text-xl font-bold text-[#E8DCC8]">{title}</h1>
        <p className="text-[10px] text-[#4A4030] mt-0.5">{new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</p>
      </div>
      <div className="flex items-center gap-3">
        {loading && (
          <div className="flex items-center gap-2 text-[10px] text-[#4CAF82]">
            <div className="w-2.5 h-2.5 border-2 border-[#4CAF82] border-t-transparent rounded-full animate-spin"></div>
            Sync...
          </div>
        )}
        <button onClick={onRefresh} className="btn-secondary py-2 text-[11px]">↻ Refresh</button>
        <button onClick={onAdd} className="btn-primary py-2 px-4 shadow-lg">+ Event</button>
      </div>
    </header>
  );
}

function StatCard({ label, value, icon, color }: any) {
  return (
    <div className="bg-white/5 border border-white/10 rounded-xl p-5 hover:bg-white/10 transition-colors">
      <div className="text-2xl mb-2" style={{ color }}>{icon}</div>
      <div className="text-2xl font-bold">{value}</div>
      <div className="text-[10px] text-[#5A5040] mt-1 tracking-wider uppercase">{label}</div>
    </div>
  );
}

function MenuBtn({ active, icon, label, onClick }: any) {
  return (
    <button onClick={onClick} className={cn("flex items-center gap-3 w-full p-2.5 rounded-lg text-sm transition-all border-l-2", 
      active ? "bg-[#C4A03C15] text-[#C4A03C] border-[#C4A03C]" : "text-[#7A6E5A] border-transparent hover:bg-white/5")}>
      <span className="text-base">{icon}</span> {label}
    </button>
  );
}

function SidebarAction({ icon, label, onClick, color }: any) {
  return (
    <button onClick={onClick} className="flex items-center gap-3 w-full p-2.5 rounded-lg text-left transition-colors hover:bg-white/5">
      <span className="text-base">{icon}</span>
      <span className="text-[12px] font-semibold" style={{ color }}>{label}</span>
    </button>
  );
}

function StatusPill({ status }: { status: string }) {
  const config = STATUS_BADGE[status] || STATUS_BADGE.Draft;
  return (
    <span className="px-2 py-0.5 rounded text-[9px] font-bold border uppercase tracking-tighter"
      style={{ backgroundColor: config.bg, color: config.color, borderColor: config.border }}>
      {status}
    </span>
  );
}

function ConfirmModal({ title, desc, onConfirm, onCancel, color, icon = "⚠" }: any) {
  return (
    <div className="fixed inset-0 z-[300] bg-black/80 flex items-center justify-center p-4">
      <div className="bg-[#0B1A0D] border border-white/10 p-8 rounded-2xl max-w-xs w-full text-center">
        <div className="text-4xl mb-4">{icon}</div>
        <h4 className="text-lg font-bold mb-2">{title}</h4>
        <p className="text-[11px] text-[#4A4030] leading-relaxed mb-6">{desc}</p>
        <div className="flex gap-2 justify-center">
          <button onClick={onConfirm} className="btn-primary !px-5" style={{ background: color, color: '#fff' }}>Iya</button>
          <button onClick={onCancel} className="px-5 py-2 border border-white/10 rounded-lg text-xs">Batal</button>
        </div>
      </div>
    </div>
  );
}
