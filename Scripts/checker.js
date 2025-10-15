// ---------- Checker Functions ----------

/**
 * addResult versi fleksibel:
 * addResult(label, dataValue, refValue, isMatch, isQty=false, unitForRef = "", unitForData=undefined)
 * - unitForData default = unitForRef
 */
function addResult(
  check,
  value,
  ref,
  isMatch,
  isQty = false,
  unitForRef = "",
  unitForData = undefined
) {
  if (unitForData === undefined) unitForData = unitForRef; // backward-compat
  const tbody = document.querySelector("#resultTable tbody");
  const row = document.createElement("tr");
  row.innerHTML = `
            <td>${check}</td>
            <td>${formatValue(value, isQty, unitForData)}</td>
            <td>${formatValue(ref, isQty, unitForRef)}</td>
            <td>${isMatch ? "Sama" : "Beda"}</td>
          `;
  row.classList.add(isMatch ? "match" : "mismatch");
  tbody.appendChild(row);
}

// Filter view (dipanggil oleh event di app.js)
function applyFilter() {
  const filter = document.getElementById("filter").value;
  const rows = document.querySelectorAll("#resultTable tbody tr");
  rows.forEach((row) => {
    if (row.classList.contains("barang-header")) return;
    if (filter === "all") row.style.display = "";
    else if (filter === "sama")
      row.style.display = row.classList.contains("match") ? "" : "none";
    else if (filter === "beda")
      row.style.display = row.classList.contains("mismatch") ? "" : "none";
  });
}

// ---- Pengecekan utama (full logic dari file asli) ----
function checkAll(sheetPL, sheetINV, sheetsDATA, kurs) {
  document.querySelector("#resultTable tbody").innerHTML = "";

  function normalize(val) {
    if (val === null || val === undefined) return "";
    if (!isNaN(val)) return parseFloat(val);
    return String(val).trim();
  }
  function isEqual(v1, v2) {
    const n1 = normalize(v1),
      n2 = normalize(v2);
    if (typeof n1 === "number" && typeof n2 === "number")
      return Math.abs(n1 - n2) < 0.01;
    return String(n1) === String(n2);
  }

  // 🔹 fungsi baru: strict compare (case & spasi sensitif)
  function isEqualStrict(v1, v2) {
    if (v1 === null || v1 === undefined) v1 = "";
    if (v2 === null || v2 === undefined) v2 = "";
    return v1 === v2; // benar-benar exact match
  }

  // === Hitung total PL (termasuk kemasanUnit)
  const { kemasanSum, bruttoSum, nettoSum, kemasanUnit } =
    hitungKemasanNWGW(sheetPL);

  // === Hitung CIF dari INV
  const rangeINV = XLSX.utils.decode_range(sheetINV["!ref"]);
  const ptIdx = document.getElementById("ptSelect").value;
  const selectedPT = kontrakData[ptIdx]?.pt || "";

  let invCols;
  if (selectedPT.includes("Shoetown")) {
    // kalau Shoetown → pakai header STYLE
    invCols = findHeaderColumns(sheetINV, {
      kode: "STYLE",
      uraian: "ITEM NAME",
      qty: "QTY",
      cif: "AMOUNT",
      suratjalan: "SURAT JALAN",
    });
  } else {
    // default → pakai header MATERIAL CODE CUSTOMER
    invCols = findHeaderColumns(sheetINV, {
      kode: "MATERIAL CODE CUSTOMER",
      uraian: "ITEM NAME",
      qty: "QTY",
      cif: "AMOUNT",
      suratjalan: "SURAT JALAN",
    });
  }
  function findInvoiceNo(sheet) {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell = sheet[XLSX.utils.encode_cell({ r: R, c: C })];
        if (
          cell &&
          typeof cell.v === "string" &&
          cell.v.toUpperCase().includes("INVOICE NO")
        ) {
          // pisahkan baris demi baris kalau cell berisi multi-line
          let lines = cell.v.split(/\r?\n/);
          for (let line of lines) {
            if (line.toUpperCase().includes("INVOICE NO")) {
              let parts = line.split(":");
              if (parts.length > 1) {
                return parts[1].trim(); // hasil: "CPMI-2025-01621"
              }
            }
          }
        }
      }
    }
    return "";
  }

  let cifSum = 0;
  if (invCols.headerRow !== null && invCols.cif !== undefined) {
    for (let r = invCols.headerRow + 1; r <= rangeINV.e.r; r++) {
      const nomorSeri = getCellValue(sheetINV, "A" + (r + 1));
      if (!nomorSeri || isNaN(nomorSeri)) continue;
      cifSum +=
        parseFloat(
          getCellValue(sheetINV, XLSX.utils.encode_cell({ r, c: invCols.cif }))
        ) || 0;
    }
  }

  // Ambil valuta otomatis dari HEADER!CI2
  const valuta = getCellValue(sheetsDATA.HEADER, "CI2") || "";
  // === HEADER
  addResult(
    "CIF",
    getCellValue(sheetsDATA.HEADER, "BU2"),
    cifSum,
    isEqual(getCellValue(sheetsDATA.HEADER, "BU2"), cifSum),
    false,
    valuta
  );
  addResult(
    "Harga Penyerahan",
    getCellValue(sheetsDATA.HEADER, "BV2"),
    cifSum * kurs,
    isEqual(getCellValue(sheetsDATA.HEADER, "BV2"), cifSum * kurs)
  );
  // === Tambahan: Dasar Pengenaan Pajak (DPP)
  const hargaPenyerahan = getCellValue(sheetsDATA.HEADER, "BV2");
  const dasarPengenaanPajak = getCellValue(sheetsDATA.HEADER, "BY2");

  // Hitung pembanding: harga penyerahan * 11%
  const dppExpected = parseFloat(hargaPenyerahan || 0) * 0.11;

  // Bandingkan
  addResult(
    "PPN 11%",
    dasarPengenaanPajak,
    dppExpected,
    Math.abs(parseFloat(dasarPengenaanPajak || 0) - dppExpected) < 0.01
  );

  // === KEMASAN
  function mapUnit(unit) {
    if (!unit) return "";
    const u = String(unit).trim().toUpperCase();
    if (u.includes("POLYBAG")) return "BG";
    if (u.includes("BOX")) return "BX";
    if (u.includes("CARTON")) return "CT";
    return u; // fallback jika tidak cocok
  }
  const kemasanUnitData = getCellValue(sheetsDATA.KEMASAN, "C2"); // satuan dari Draft EXIM
  const kemasanQtyData = getCellValue(sheetsDATA.KEMASAN, "D2"); // angka dari Draft EXIM

  const kemasanUnitMapped = mapUnit(kemasanUnit); // mapping POLYBAG -> BG, BOX -> BX, CARTON -> CT
  const kemasanUnitDataMapped = mapUnit(kemasanUnitData); // mapping dari file data juga
  // fungsi mapping kemasan

  function normalizeUnit(u) {
    if (!u) return "";
    return String(u).trim().toUpperCase();
  }

  const angkaMatch = isEqual(kemasanQtyData, kemasanSum);
  const unitMatch =
    normalizeUnit(kemasanUnitDataMapped) === normalizeUnit(kemasanUnitMapped);

  addResult(
    "Total Kemasan",
    kemasanQtyData + " " + kemasanUnitDataMapped, // tampilkan draft dengan singkatan
    kemasanSum + " " + kemasanUnitMapped, // tampilkan hasil hitung dengan singkatan
    angkaMatch && unitMatch,
    true
  );
  addResult(
    "Brutto",
    getCellValue(sheetsDATA.HEADER, "CB2"),
    bruttoSum,
    isEqual(getCellValue(sheetsDATA.HEADER, "CB2"), bruttoSum),
    false,
    "KG"
  );
  addResult(
    "Netto",
    getCellValue(sheetsDATA.HEADER, "CC2"),
    nettoSum,
    isEqual(getCellValue(sheetsDATA.HEADER, "CC2"), nettoSum),
    false,
    "KG"
  );

  // === DOKUMEN
  const invInvoiceNo = findInvoiceNo(sheetINV);
  const plInvoiceNo = findInvoiceNo(sheetPL);
  addResult(
    "Invoice No.",
    getCellValue(sheetsDATA.DOKUMEN, "D2"), // dari DOKUMEN
    invInvoiceNo,
    isEqual(getCellValue(sheetsDATA.DOKUMEN, "D2"), invInvoiceNo)
  );
  addResult(
    "Packinglist No.",
    getCellValue(sheetsDATA.DOKUMEN, "D2"),
    plInvoiceNo,
    isEqual(getCellValue(sheetsDATA.DOKUMEN, "D2"), plInvoiceNo)
  );
  let invSuratJalan = "";
  if (invCols.suratjalan !== undefined && invCols.headerRow !== null) {
    invSuratJalan = getCellValue(
      sheetINV,
      XLSX.utils.encode_cell({
        r: invCols.headerRow + 1,
        c: invCols.suratjalan,
      })
    );
  }
  addResult(
    "Delivery Order",
    getCellValue(sheetsDATA.DOKUMEN, "D5"),
    invSuratJalan,
    isEqual(getCellValue(sheetsDATA.DOKUMEN, "D5"), invSuratJalan)
  );

  // ---------- Tambahkan fungsi helper baru ----------
  // Fungsi ini akan coba konversi input angka Excel (serial date) atau string menjadi YYYY-MM-DD
  function formatAsDate(val) {
    if (!val) return "";
    // jika val numeric (seperti 45901.00 dari Excel)
    if (!isNaN(val)) {
      const excelEpoch = new Date(1899, 11, 30); // base date Excel
      const d = new Date(excelEpoch.getTime() + val * 24 * 60 * 60 * 1000);
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const dd = String(d.getDate()).padStart(2, "0");
      return `${yyyy}-${mm}-${dd}`;
    }
    // kalau val string, coba parse langsung
    const d = new Date(val);
    if (!isNaN(d)) {
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const dd = String(d.getDate()).padStart(2, "0");
      return `${yyyy}-${mm}-${dd}`;
    }
    return val; // fallback, biarkan apa adanya
  }

  // === Nomor & Tanggal Kontrak: DATA (draft) vs FIELD (input dari dropdown)
  const draftNoKontrakRaw = getCellValue(sheetsDATA.DOKUMEN, "D4");
  const draftNoKontrak = formatAsDate(draftNoKontrakRaw);
  const fieldNoKontrak = document.getElementById("noKontrak").value;
  addResult(
    "Contract Number",
    draftNoKontrak,
    fieldNoKontrak,
    isEqual(draftNoKontrak, fieldNoKontrak)
  );

  const draftTglKontrakRaw = getCellValue(sheetsDATA.DOKUMEN, "E4");
  const draftTglKontrak = formatAsDate(draftTglKontrakRaw);
  const fieldTglKontrak = document.getElementById("tglKontrak").value;
  addResult(
    "Contract Date",
    draftTglKontrak,
    fieldTglKontrak,
    isEqual(draftTglKontrak, fieldTglKontrak)
  );

  // === BARANG
  const rangeBarang = XLSX.utils.decode_range(sheetsDATA.BARANG["!ref"]);
  const plCols = findHeaderColumns(sheetPL, { nw: "NW", gw: "GW" });

  let barangCounter = 1;
  for (let r = 1; r <= rangeBarang.e.r; r++) {
    const kodeBarang = getCellValue(sheetsDATA.BARANG, "D" + (r + 1));
    if (!kodeBarang) continue;

    const rowINV = (invCols.headerRow || 0) + r;
    const rowPL = (plCols.headerRow || 0) + r;

    // 🔹 Tambahkan baris judul barang collapsible
    const tbody = document.querySelector("#resultTable tbody");
    const rowHeader = document.createElement("tr");
    rowHeader.classList.add("fw-bold", "barang-header");
    rowHeader.setAttribute("data-target", "barang-" + barangCounter);
    rowHeader.innerHTML = `
            <td colspan="4">Barang ke-${barangCounter}</td>
          `;
    tbody.appendChild(rowHeader);

    // Kode Barang
    const invKodeVal =
      invCols.kode !== undefined
        ? getCellValue(
            sheetINV,
            XLSX.utils.encode_cell({ r: rowINV, c: invCols.kode })
          )
        : "";
    addResult("Code", kodeBarang, invKodeVal, isEqual(kodeBarang, invKodeVal));

    // Uraian
    const draftUraian = getCellValue(sheetsDATA.BARANG, "E" + (r + 1));
    const invUraian =
      invCols.uraian !== undefined
        ? getCellValue(
            sheetINV,
            XLSX.utils.encode_cell({ r: rowINV, c: invCols.uraian })
          )
        : "";
    addResult(
      "Name",
      draftUraian,
      invUraian,
      isEqualStrict(draftUraian, invUraian) // 🔹 ganti ke strict
    );

    // QTY Barang
    const draftQty = getCellValue(sheetsDATA.BARANG, "K" + (r + 1));

    // ambil QTY dari INV
    const invQty =
      invCols.qty !== undefined
        ? getCellValue(
            sheetINV,
            XLSX.utils.encode_cell({ r: rowINV, c: invCols.qty })
          )
        : "";

    // Unit dari Draft (file DATA, sheet BARANG, cell J2)
    const draftUnit = getCellValue(sheetsDATA.BARANG, "J2") || "NPR";

    // Unit dari dropdown (untuk INV & PL)
    const selectedUnit = document.getElementById("unitSelect").value || "NPR";

    // Cek angka & unit
    const qtyMatch = isEqual(draftQty, invQty);
    const unitMatch = draftUnit === selectedUnit;

    addResult(
      "Quantity",
      draftQty,
      invQty,
      qtyMatch && unitMatch, // Sama hanya jika angka & unit sama
      true,
      selectedUnit, // untuk INV/PL tampil sesuai dropdown
      draftUnit // untuk Draft tampil sesuai J2
    );

    // Netto/Item
    const draftNettoItem = getCellValue(sheetsDATA.BARANG, "T" + (r + 1));
    const plNettoItem =
      plCols.nw !== undefined
        ? getCellValue(
            sheetPL,
            XLSX.utils.encode_cell({ r: rowPL, c: plCols.nw })
          )
        : "";
    addResult(
      "NW",
      draftNettoItem,
      plNettoItem,
      isEqual(draftNettoItem, plNettoItem),
      false,
      "KG"
    );

    // Brutto/Item
    const draftBruttoItem = getCellValue(sheetsDATA.BARANG, "U" + (r + 1));
    const plBruttoItem =
      plCols.gw !== undefined
        ? getCellValue(
            sheetPL,
            XLSX.utils.encode_cell({ r: rowPL, c: plCols.gw })
          )
        : "";
    addResult(
      "GW",
      draftBruttoItem,
      plBruttoItem,
      isEqual(draftBruttoItem, plBruttoItem),
      false,
      "KG"
    );
    // 🔹 CIF per item
    const draftCIF = getCellValue(sheetsDATA.BARANG, "Z" + (r + 1)); // ambil dari draft EXIM
    const invCIF =
      invCols.cif !== undefined
        ? getCellValue(
            sheetINV,
            XLSX.utils.encode_cell({ r: rowINV, c: invCols.cif })
          )
        : "";
    addResult(
      "Amount",
      draftCIF,
      invCIF,
      isEqual(draftCIF, invCIF),
      false,
      "USD"
    );

    barangCounter++;
  }

  // 🔹 Aktifkan collapsible
  document.querySelectorAll(".barang-header").forEach((header) => {
    header.addEventListener("click", () => {
      let next = header.nextElementSibling;
      while (next && !next.classList.contains("barang-header")) {
        next.style.display = next.style.display === "none" ? "" : "none";
        next = next.nextElementSibling;
      }
    });
  });
}
