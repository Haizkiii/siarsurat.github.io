// Register Surat Kelurahan (offline, localStorage)
// Fokus: nomor urut per perihal per tahun, tidak saling campur.

const STORAGE_KEY = "sr_sim_v1";

const CLASS_LABEL = {
  SKTM: "Surat Keterangan Tidak Mampu",
  DOMISILI: "Surat Keterangan Domisili",
  SKCK: "Surat Keterangan Catatan Kepolisian",
  KELAHIRAN: "Surat Keterangan Kelahiran",
  KEMATIAN: "Surat Keterangan Kematian",
  DATANGPINDAH: "Surat Keterangan Datang/Pindah",
  SKM: "Surat Keterangan Menikah",
  WARIS: "Surat Keterangan Waris",
  USAHA: "Surat Keterangan Usaha",
  KELUAR: "Surat Keluar",
  LAIN: "Surat Keterangan Lainnya",
};

const CLASS_PREFIX = {
  SKTM: "460",
  DOMISILI: "470",
  SKCK: "474",
  KELAHIRAN: "474.1",
  KEMATIAN: "474.2",
  DATANGPINDAH: "474.3",
  SKM: "474.4",
  WARIS: "474.5",
  USAHA: "510",
  KELUAR: "800",
  LAIN: "900",
};

function pad(n, width = 4) {
  const s = String(n);
  return s.length >= width ? s : "0".repeat(width - s.length) + s;
}

function monthToRoman(month) {
  const romanNumerals = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"];
  return romanNumerals[month] || "";
}

function ymdToYear(ymd) {
  if (!ymd) return new Date().getFullYear();
  return Number(String(ymd).slice(0, 4));
}

function nowISO() {
  const d = new Date();
  return d.toISOString();
}

// Utility to dynamically load a script and wait until it's loaded
function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.async = true;
    s.onload = () => resolve();
    s.onerror = (e) => reject(e);
    document.head.appendChild(s);
  });
}
// Try multiple CDNs with a timeout per attempt
function loadScriptWithFallback(urls, timeoutMs = 8000) {
  return new Promise(async (resolve, reject) => {
    for (const url of urls) {
      try {
        await Promise.race([loadScript(url), new Promise((_, rej) => setTimeout(() => rej(new Error("timeout")), timeoutMs))]);
        // give browser a tick to register global
        await new Promise((r) => setTimeout(r, 50));
        if (typeof XLSX !== "undefined") return resolve();
        // if script loaded but XLSX not present, continue to next
      } catch (err) {
        console.warn(`Failed loading ${url}:`, err);
        // try next
      }
    }
    reject(new Error("All CDN attempts to load XLSX failed"));
  });
}
function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { counters: {}, letters: [] };
    const parsed = JSON.parse(raw);
    if (!parsed.counters) parsed.counters = {};
    if (!parsed.letters) parsed.letters = [];
    return parsed;
  } catch {
    return { counters: {}, letters: [] };
  }
}

function saveState(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function counterKey(classification, year) {
  return `${classification}:${year}`;
}

function formatRegisterNo(classification, year, month, regNo) {
  // Desired format: PREFIX/0001-Kel/I/2026 (number first, then -Kel, then month roman and year)
  return `${CLASS_PREFIX[classification]}/${pad(regNo)}-Kel/${monthToRoman(month)}/${year}`;
}

// Migrate existing stored letters to ensure `register_display` matches current format

// Atomic-ish generator (single-tab safe). For multi-user server, this must be transactional DB.
function generateNextNumber(state, classification, year) {
  const key = counterKey(classification, year);
  const last = state.counters[key] ?? 0;
  const next = last + 1;
  state.counters[key] = next;
  return next;
}

function addLetter({ classification, letterDate, fullName, birthPlaceDate, occupation, address }) {
  const state = loadState();
  const year = ymdToYear(letterDate);
  const month = Number(String(letterDate).slice(5, 7)); // Extract MM from YYYY-MM-DD
  const register_no = generateNextNumber(state, classification, year);

  const letter = {
    id: crypto.randomUUID(),
    classification,
    letter_date: letterDate,
    year,
    month,
    register_no,
    register_display: formatRegisterNo(classification, year, month, register_no),
    fullName: fullName.trim(),
    birthPlaceDate: birthPlaceDate.trim(),
    occupation: occupation.trim(),
    address: address.trim(),
    status: "AKTIF", // AKTIF | BATAL
    created_at: nowISO(),
    updated_at: nowISO(),
  };

  state.letters.unshift(letter); // newest first
  saveState(state);
  return letter;
}

function cancelLetter(id) {
  const state = loadState();
  const idx = state.letters.findIndex((x) => x.id === id);
  if (idx === -1) return false;
  state.letters[idx].status = "BATAL";
  state.letters[idx].updated_at = nowISO();
  saveState(state);
  return true;
}

function hardDeleteLetter(id) {
  const state = loadState();
  state.letters = state.letters.filter((x) => x.id !== id);
  saveState(state);
}

function exportPDF() {
  const state = loadState();
  if (!state.letters || state.letters.length === 0) {
    alert("Tidak ada data untuk diekspor.");
    return;
  }

  try {
    // Determine selected perihal from filterClass
    const sel = (filterClass && filterClass.value) || "ALL";
    const selLabel = sel === "ALL" ? "Semua Perihal" : CLASS_LABEL[sel] || sel;

    // Filter letters by selected classification
    const letters = state.letters.filter((l) => (sel === "ALL" ? true : l.classification === sel));

    // If no letters for this perihal, still generate PDF with NIHIL message
    const isNihil = letters.length === 0;

    // Resolve jsPDF constructor
    let jsPDFCtor = null;
    if (typeof window.jsPDF === "function") jsPDFCtor = window.jsPDF;
    else if (window.jspdf && typeof window.jspdf.jsPDF === "function") jsPDFCtor = window.jspdf.jsPDF;
    else if (typeof window.jspdf === "function") jsPDFCtor = window.jspdf;

    if (!jsPDFCtor) {
      alert("⚠️ Library PDF tidak terdeteksi. Refresh halaman dan coba lagi.");
      console.error("jsPDF not found on window object.");
      return;
    }

    const doc = new jsPDFCtor("l", "mm", "a4");
    const pageWidth = doc.internal.pageSize.getWidth();

    // Determine month/year range from exported letters
    const monthNames = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const monthYearSet = new Set();
    for (const l of letters) {
      const d = l.letter_date || "";
      if (/^\d{4}-\d{2}/.test(d)) monthYearSet.add(d.slice(0, 7)); // YYYY-MM
    }
    const monthYears = Array.from(monthYearSet).sort();

    // Prepare 3 header lines: line1="Register", line2=perihal name, line3="<MonthName> <Year>" (no word 'Tahun') or period
    let headerLine1 = "Register";
    let headerLine2 = selLabel;
    let headerLine3 = "";
    let filenameSuffix = "";

    if (monthYears.length === 1) {
      const [ystr, mstr] = monthYears[0].split("-");
      const mon = Number(mstr);
      const yr = Number(ystr);
      headerLine3 = `Bulan ${monthNames[mon]} ${yr}`; // e.g. "Bulan Desember 2025"
      filenameSuffix = `bulan-${monthNames[mon]}-tahun-${yr}`;
    } else if (monthYears.length > 1) {
      const start = monthYears[0].split("-");
      const end = monthYears[monthYears.length - 1].split("-");
      const sMon = Number(start[1]);
      const sYr = Number(start[0]);
      const eMon = Number(end[1]);
      const eYr = Number(end[0]);
      headerLine3 = `Bulan ${monthNames[sMon]} ${sYr} - ${monthNames[eMon]} ${eYr}`;
      filenameSuffix = `periode-${monthNames[sMon]}${sYr}-to-${monthNames[eMon]}${eYr}`;
    } else {
      // no dates available in letters -> use current month/year
      const now = new Date();
      headerLine3 = `Bulan ${monthNames[now.getMonth() + 1]} ${now.getFullYear()}`;
      filenameSuffix = `bulan-${monthNames[now.getMonth() + 1]}-tahun-${now.getFullYear()}`;
    }

    // Header (3 lines) - use bold sans-serif (Helvetica) as fallback for Arial Narrow Bold
    // Note: Arial Narrow Bold is proprietary; to use it we need the TTF and to embed it via addFileToVFS/addFont.
    // Using built-in Helvetica Bold for now.
    try {
      doc.setFont("helvetica", "bold");
    } catch (e) {
      // ignore if font not available
    }
    // Make header lines uppercase and set size 14
    doc.setFontSize(14);
    doc.text(String(headerLine1).toUpperCase(), pageWidth / 2, 14, { align: "center" });
    doc.text(String(headerLine2).toUpperCase(), pageWidth / 2, 20, { align: "center" });
    doc.text(String(headerLine3).toUpperCase(), pageWidth / 2, 26, { align: "center" });
    // Reset to normal for table body
    try {
      doc.setFont("helvetica", "normal");
    } catch (e) {}

    // Prepare table data with NO column for numbering
    let tableData;
    if (isNihil) {
      // Show NIHIL message in table if no data
      tableData = [["", "NIHIL", "", "", "", "", "", ""]];
    } else {
      tableData = letters.map((letter, index) => [
        String(index + 1), // NO column
        letter.register_display,
        letter.letter_date || "-",
        CLASS_LABEL[letter.classification] || letter.classification,
        letter.fullName,
        letter.birthPlaceDate || "-",
        letter.occupation || "-",
        letter.address || "-",
      ]);
    }

    // Add table using autoTable if available
    if (typeof doc.autoTable === "function") {
      doc.autoTable({
        head: [["NO", "No. Surat", "Tgl Surat", "Perihal", "Nama Lengkap", "Tempat, Tanggal Lahir", "Pekerjaan", "Alamat"].map((h) => String(h).toUpperCase())],
        body: tableData,
        startY: 32,
        margin: { left: 6, right: 6 },
        styles: { fontSize: 9, cellPadding: 2, font: "helvetica", halign: "center", lineColor: [0, 0, 0], lineWidth: 0.2 },
        headStyles: { fillColor: [33, 33, 33], textColor: 255, fontStyle: "bold", font: "helvetica", fontSize: 10, halign: "center", lineColor: [0, 0, 0], lineWidth: 0.2 },
        alternateRowStyles: { fillColor: [245, 245, 245], lineColor: [0, 0, 0], lineWidth: 0.2 },
        columnStyles: {
          0: { halign: "center" }, // NO
          1: { halign: "center" }, // No. Surat
          2: { halign: "center" }, // Tgl Surat
          3: { halign: "center" }, // Perihal
          4: { halign: "left" }, // Nama Lengkap
          5: { halign: "center" }, // Tempat, Tanggal Lahir
          6: { halign: "center" }, // Pekerjaan
          7: { halign: "left" }, // Alamat
        },
        didDrawPage: (data) => {},
      });
    } else {
      // Fallback: simple text list
      let y = 32;
      doc.setFontSize(9);
      for (const row of tableData) {
        const line = row.join(" | ");
        doc.text(line, 10, y);
        y += 6;
        if (y > doc.internal.pageSize.getHeight() - 20) {
          doc.addPage();
          y = 20;
        }
      }
    }

    const safeLabel = (sel === "ALL" ? "all" : sel).toLowerCase().replace(/[^a-z0-9]+/gi, "-");
    const filename = `Register_${safeLabel}_${filenameSuffix || new Date().toISOString().slice(0, 10)}.pdf`;
    doc.save(filename);
  } catch (error) {
    console.error("Error generating filtered PDF:", error);
    alert(`Gagal membuat PDF: ${error.message || String(error)}`);
  }
}

// importJSON removed per request

// UI
const $ = (q) => document.querySelector(q);
const tbody = $("#tbody");
const counterGrid = $("#counterGrid");
const stats = $("#stats");

const form = $("#letterForm");
const classificationEl = $("#classification");
const dateEl = $("#letterDate");
const fullNameEl = $("#fullName");
const birthPlaceDateEl = $("#birthPlaceDate");
const occupationEl = $("#occupation");
const addressEl = $("#address");
const previewEl = $("#previewNo");

const filterClass = $("#filterClass");
const filterYear = $("#filterYear");
const filterNo = $("#filterNo");

const btnExportPDF = $("#btnExportPDF");
const btnExportExcel = $("#btnExportExcel");
const btnReset = $("#btnReset");

// Dialog helpers
const dlg = $("#dlg");
const dlgTitle = $("#dlgTitle");
const dlgBody = $("#dlgBody");
const dlgFoot = $("#dlgFoot");
$("#dlgClose").addEventListener("click", () => dlg.close());

function showDialog({ title, bodyHTML, buttons }) {
  dlgTitle.textContent = title || "Konfirmasi";
  dlgBody.innerHTML = bodyHTML || "";
  dlgFoot.innerHTML = "";
  (buttons || []).forEach((b) => {
    const btn = document.createElement("button");
    btn.className = `btn ${b.variant || "secondary"}`;
    btn.textContent = b.text;
    btn.addEventListener("click", () => {
      if (b.onClick) b.onClick();
      if (!b.keepOpen) dlg.close();
    });
    dlgFoot.appendChild(btn);
  });
  dlg.showModal();
}

function updatePreview() {
  const cls = classificationEl.value;
  const dateValue = dateEl.value || new Date().toISOString().slice(0, 10);
  const year = ymdToYear(dateValue);
  const month = Number(String(dateValue).slice(5, 7));
  const state = loadState();
  const key = counterKey(cls, year);
  const next = (state.counters[key] ?? 0) + 1;
  previewEl.textContent = formatRegisterNo(cls, year, month, next);
}

function updateFormFields() {
  // Placeholder function - fields are managed via HTML form structure
}

function matchFilters(letter) {
  const cls = filterClass.value;
  const yr = filterYear.value ? Number(filterYear.value) : null;
  const q = (filterNo.value || "").trim().toLowerCase();

  if (cls && letter.classification !== cls) return false;
  if (yr && letter.year !== yr) return false;

  if (q) {
    const hay = letter.register_display.toLowerCase();
    if (!hay.includes(q)) return false;
  }
  return true;
}

function updateFilterClass() {
  // This function updates the filter class options based on available data
  // Currently all options are predefined in HTML, so no action needed here
  // This is a placeholder to prevent errors when called from render()
}

function render() {
  const state = loadState();
  updateFilterClass(); // Update filter options sesuai data yang ada
  const list = state.letters.filter(matchFilters);

  tbody.innerHTML = "";
  list.forEach((letter, index) => {
    const tr = document.createElement("tr");

    const statusBadge = letter.status === "AKTIF" ? `<span class="badge ok">● Aktif</span>` : `<span class="badge danger">● Batal</span>`;

    tr.innerHTML = `
      <td class="center">${index + 1}</td>
      <td class="mono"><b>${letter.register_display}</b></td>
      <td>${letter.letter_date || "-"}</td>
      <td>${CLASS_LABEL[letter.classification]}</td>
      <td>${escapeHtml(letter.fullName)}</td>
      <td>${escapeHtml(letter.birthPlaceDate || "-")}</td>
      <td>${escapeHtml(letter.occupation || "-")}</td>
      <td>${escapeHtml(letter.address || "-")}</td>
      <td>${statusBadge}</td>
      <td>
        <button class="btn secondary" data-act="copy" data-id="${letter.id}">Copy</button>
        ${letter.status === "AKTIF" ? `<button class="btn danger" data-act="cancel" data-id="${letter.id}">Batalkan</button>` : ""}
        <button class="btn secondary" data-act="delete" data-id="${letter.id}">Hapus</button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  const total = state.letters.length;
  const aktif = state.letters.filter((x) => x.status === "AKTIF").length;
  const batal = total - aktif;
  stats.textContent = `Total: ${total} • Aktif: ${aktif} • Batal: ${batal} • Ditampilkan: ${list.length}`;

  renderCounters(state);
  updatePreview();
}

function renderCounters(state) {
  const keys = Object.keys(state.counters).sort((a, b) => a.localeCompare(b));
  if (!keys.length) {
    counterGrid.innerHTML = `<div class="muted">Belum ada data. Counter akan muncul setelah kamu menyimpan surat.</div>`;
    return;
  }
  // Group by year
  const grouped = {};
  for (const k of keys) {
    const [cls, year] = k.split(":");
    if (!grouped[year]) grouped[year] = {};
    grouped[year][cls] = state.counters[k];
  }
  const years = Object.keys(grouped).sort((a, b) => Number(b) - Number(a));

  counterGrid.innerHTML = "";
  for (const year of years) {
    for (const cls of ["SKTM", "DOMISILI", "SKCK", "KELAHIRAN", "KEMATIAN", "DATANGPINDAH", "SKM", "WARIS", "USAHA", "KELUAR", "LAIN"]) {
      const last = grouped[year][cls] ?? 0;
      const div = document.createElement("div");
      div.className = "counterCard";
      // Show month roman for all years and place number before '-Kel' (e.g. 460/0001-Kel/I/2026)
      const monthRoman = monthToRoman(new Date().getMonth() + 1);
      const displayReg = `${CLASS_PREFIX[cls]}/${pad(last)}-Kel/${monthRoman}/${year}`;

      div.innerHTML = `
        <div class="k">${CLASS_LABEL[cls]}</div>
        <div class="v mono">${displayReg}</div>
        <div class="s">Nomor terakhir: <b>#${last}</b></div>
      `;
      counterGrid.appendChild(div);
    }
  }
}

function escapeHtml(s) {
  return String(s).replace(
    /[&<>"']/g,
    (m) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#039;",
      }[m])
  );
}

// Events
form.addEventListener("submit", (e) => {
  e.preventDefault();
  const cls = classificationEl.value;
  const dt = dateEl.value;
  const fullName = fullNameEl.value;
  const birthPlaceDate = birthPlaceDateEl.value;
  const occupation = occupationEl.value;
  const address = addressEl.value;

  if (!dt) return alert("Tanggal surat wajib diisi.");
  if (!fullName.trim()) return alert("Nama lengkap wajib diisi.");

  const letter = addLetter({ classification: cls, letterDate: dt, fullName, birthPlaceDate, occupation, address });

  fullNameEl.value = "";
  birthPlaceDateEl.value = "";
  occupationEl.value = "";
  addressEl.value = "";

  showDialog({
    title: "Berhasil",
    bodyHTML: `
      <div>Nomor register dibuat:</div>
      <div style="margin-top:10px" class="previewValue mono"><b>${letter.register_display}</b></div>
      <div class="small" style="margin-top:10px">Catatan: Pembatalan tidak mengubah urutan. Nomor tetap tercatat untuk histori.</div>
    `,
    buttons: [
      {
        text: "Copy Nomor",
        variant: "secondary",
        onClick: async () => {
          await navigator.clipboard.writeText(letter.register_display);
        },
        keepOpen: true,
      },
      { text: "Tutup", variant: "primary" },
    ],
  });

  render();
});

[classificationEl, dateEl].forEach((el) =>
  el.addEventListener("change", () => {
    updatePreview();
    updateFormFields();
  })
);

[filterClass, filterYear, filterNo].forEach((el) => el.addEventListener("input", render));

tbody.addEventListener("click", async (e) => {
  const btn = e.target.closest("button");
  if (!btn) return;
  const act = btn.dataset.act;
  const id = btn.dataset.id;
  const state = loadState();
  const letter = state.letters.find((x) => x.id === id);
  if (!letter) return;

  if (act === "copy") {
    await navigator.clipboard.writeText(letter.register_display);
    btn.textContent = "Copied!";
    setTimeout(() => (btn.textContent = "Copy"), 800);
    return;
  }

  if (act === "cancel") {
    showDialog({
      title: "Batalkan surat?",
      bodyHTML: `
        <div>Ini akan mengubah status menjadi <b>BATAL</b> tanpa mengubah urutan nomor.</div>
        <div style="margin-top:10px" class="small">Target: <span class="mono"><b>${letter.register_display}</b></span></div>
      `,
      buttons: [
        { text: "Batal", variant: "secondary" },
        {
          text: "Ya, Batalkan",
          variant: "danger",
          onClick: () => {
            cancelLetter(id);
            render();
          },
        },
      ],
    });
    return;
  }

  if (act === "delete") {
    showDialog({
      title: "Hapus data?",
      bodyHTML: `
        <div><b>Warning:</b> Ini menghapus data dari perangkat ini. Untuk simulasi boleh, tapi di sistem nyata biasanya pakai audit + status.</div>
        <div style="margin-top:10px" class="small">Target: <span class="mono"><b>${letter.register_display}</b></span></div>
      `,
      buttons: [
        { text: "Batal", variant: "secondary" },
        {
          text: "Ya, Hapus",
          variant: "danger",
          onClick: () => {
            hardDeleteLetter(id);
            render();
          },
        },
      ],
    });
    return;
  }
});

btnExportPDF.addEventListener("click", exportPDF);

async function exportExcel() {
  const state = loadState();
  if (!state.letters || state.letters.length === 0) {
    alert("Tidak ada data untuk diekspor.");
    return;
  }
  // Check if XLSX library is available; try to load it dynamically if missing
  if (typeof XLSX === "undefined") {
    const cdns = ["https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.min.js", "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js", "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js"];
    try {
      await loadScriptWithFallback(cdns, 8000);
    } catch (err) {
      alert("⚠️ Library Excel tidak terdeteksi dan gagal dimuat dari CDN.\nMohon periksa koneksi internet atau tambahkan file xlsx.min.js secara lokal ke proyek.");
      console.error("All attempts to load XLSX failed:", err);
      return;
    }
  }

  try {
    // Determine selected perihal
    const sel = (filterClass && filterClass.value) || "ALL";
    const selLabel = sel === "ALL" ? "Semua Perihal" : CLASS_LABEL[sel] || sel;

    // Filter letters by selected classification
    const letters = state.letters.filter((l) => (sel === "ALL" ? true : l.classification === sel));

    const isNihil = letters.length === 0;

    // Prepare month/year for header
    const monthNames = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const monthYearSet = new Set();
    for (const l of letters) {
      const d = l.letter_date || "";
      if (/^\d{4}-\d{2}/.test(d)) monthYearSet.add(d.slice(0, 7));
    }
    const monthYears = Array.from(monthYearSet).sort();

    let headerLine1 = "REGISTER";
    let headerLine2 = String(selLabel).toUpperCase();
    let headerLine3 = "";

    if (monthYears.length === 1) {
      const [ystr, mstr] = monthYears[0].split("-");
      const mon = Number(mstr);
      const yr = Number(ystr);
      headerLine3 = `BULAN ${monthNames[mon].toUpperCase()} ${yr}`;
    } else if (monthYears.length > 1) {
      const start = monthYears[0].split("-");
      const end = monthYears[monthYears.length - 1].split("-");
      const sMon = Number(start[1]);
      const sYr = Number(start[0]);
      const eMon = Number(end[1]);
      const eYr = Number(end[0]);
      headerLine3 = `BULAN ${monthNames[sMon].toUpperCase()} ${sYr} - ${monthNames[eMon].toUpperCase()} ${eYr}`;
    } else {
      const now = new Date();
      headerLine3 = `BULAN ${monthNames[now.getMonth() + 1].toUpperCase()} ${now.getFullYear()}`;
    }

    // Prepare sheet data
    const sheetData = [];

    // Add header rows (3 baris)
    sheetData.push([headerLine1]);
    sheetData.push([headerLine2]);
    sheetData.push([headerLine3]);
    sheetData.push([]); // Empty row for spacing

    // Add column header row
    sheetData.push(["NO", "No. Surat", "Tgl Surat", "Perihal", "Nama Lengkap", "Tempat, Tanggal Lahir", "Pekerjaan", "Alamat"]);

    // Add data rows
    if (isNihil) {
      sheetData.push(["", "NIHIL", "", "", "", "", "", ""]);
    } else {
      letters.forEach((letter, index) => {
        sheetData.push([
          index + 1,
          letter.register_display,
          letter.letter_date || "-",
          CLASS_LABEL[letter.classification] || letter.classification,
          letter.fullName,
          letter.birthPlaceDate || "-",
          letter.occupation || "-",
          letter.address || "-",
        ]);
      });
    }

    // Create workbook
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, selLabel.slice(0, 31)); // Sheet name max 31 chars

    // Set column widths
    ws["!cols"] = [
      { wch: 5 }, // NO
      { wch: 20 }, // No. Surat
      { wch: 12 }, // Tgl Surat
      { wch: 25 }, // Perihal
      { wch: 25 }, // Nama Lengkap
      { wch: 25 }, // Tempat, Tanggal Lahir
      { wch: 18 }, // Pekerjaan
      { wch: 25 }, // Alamat
    ];

    // Format header rows (rows 1-3) - bold, size 14, center aligned
    for (let row = 0; row < 3; row++) {
      for (let col = 0; col < 8; col++) {
        const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
        if (ws[cellRef]) {
          ws[cellRef].s = {
            font: { bold: true, size: 14 },
            alignment: { horizontal: "center", vertical: "center", wrapText: true },
          };
        }
      }
    }

    // Format column header row (row 5, after empty row 4)
    for (let col = 0; col < 8; col++) {
      const cellRef = XLSX.utils.encode_cell({ r: 4, c: col });
      if (ws[cellRef]) {
        ws[cellRef].s = {
          font: { bold: true, size: 10 },
          alignment: { horizontal: "center", vertical: "center" },
        };
      }
    }

    // Generate filename
    let filename = `Register_${selLabel.replace(/\s+/g, "_")}`;
    if (monthYears.length === 1) {
      const [ystr, mstr] = monthYears[0].split("-");
      const mon = Number(mstr);
      const yr = Number(ystr);
      filename += `_${monthNames[mon]}_${yr}`;
    } else if (monthYears.length > 1) {
      filename += `_periode`;
    }
    filename += `.xlsx`;

    XLSX.writeFile(wb, filename);
  } catch (error) {
    console.error("Error generating Excel:", error);
    alert(`Gagal membuat Excel: ${error.message || String(error)}`);
  }
}

btnExportExcel.addEventListener("click", exportExcel);

// Import JSON UI and behavior removed

btnReset.addEventListener("click", () => {
  showDialog({
    title: "Reset semua data?",
    bodyHTML: `<div>Ini menghapus semua counters dan data surat dari perangkat ini.</div>`,
    buttons: [
      { text: "Batal", variant: "secondary" },
      {
        text: "Ya, Reset",
        variant: "danger",
        onClick: () => {
          localStorage.removeItem(STORAGE_KEY);
          render();
        },
      },
    ],
  });
});

// Init defaults
(function init() {
  const today = new Date().toISOString().slice(0, 10);
  dateEl.value = today;
  // Migrate stored entries to new register format if needed
  updatePreview();
  render();
})();
