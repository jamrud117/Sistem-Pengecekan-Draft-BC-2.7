// ---------- util ----------

function readExcelFile(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      resolve(workbook);
    };
    reader.readAsArrayBuffer(file);
  });
}

function getCellValue(sheet, cell) {
  const c = sheet[cell];
  return c ? c.v : "";
}

function getCellValueRC(sheet, r, c) {
  const cell = sheet[XLSX.utils.encode_cell({ r, c })];
  return cell ? cell.v : "";
}

// Normalisasi kurs (contoh: "16.460,00" -> 16460)
function parseKurs(val) {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number") return val;
  let s = String(val).trim();
  s = s.replace(/\u00A0/g, ""); // non-breaking spaces
  // hapus simbol mata uang & spasi
  s = s.replace(/[^\d,\.\-]/g, "");
  if (s.indexOf(",") > -1 && s.indexOf(".") > -1) {
    // format "16.460,00"
    s = s.replace(/\./g, "").replace(",", ".");
  } else {
    s = s.replace(",", ".");
  }
  const n = parseFloat(s);
  return isNaN(n) ? "" : n;
}

// Format angka (QTY & kemasan integer, lainnya float)
function formatValue(val, isQty = false, unit = "") {
  if (val === null || val === undefined || val === "") return "";
  const n = Number(val);
  let result = val;
  if (!isNaN(n)) {
    result = isQty ? Math.round(n) : n.toFixed(2);
  }
  return unit ? result + " " + unit : result;
}

function cleanNumber(val) {
  if (!val) return "";
  return String(val)
    .replace(/.*?:\s*/i, "")
    .trim();
}

// Deteksi jenis file: DATA, PL, atau INV
function detectFileType(workbook) {
  if (workbook.SheetNames.includes("HEADER")) return "DATA";
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet || !sheet["!ref"]) return "INV";
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "string") {
        const val = cell.v.toString().toUpperCase();
        if (
          val.includes("KEMASAN") ||
          val.includes("GW") ||
          val.includes("NW")
        ) {
          return "PL";
        }
      }
    }
  }
  return "INV";
}

// Cari kolom berdasarkan header (tidak diubah)
function findHeaderColumns(sheet, headers) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  let found = {},
    headerRow = null;
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "string") {
        const val = cell.v.toString().trim().toUpperCase();
        for (const key in headers) {
          if (val.includes(headers[key].toUpperCase())) {
            found[key] = c;
          }
        }
      }
    }
    if (Object.keys(found).length > 0) {
      headerRow = r;
      break;
    }
  }
  return { ...found, headerRow };
}

// Hitung total dari PL + deteksi satuan kemasan
function hitungKemasanNWGW(sheet) {
  if (!sheet || !sheet["!ref"]) {
    return { kemasanSum: 0, bruttoSum: 0, nettoSum: 0, kemasanUnit: "" };
  }
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  let colKemasan = null,
    colGW = null,
    colNW = null,
    headerRow = null,
    kemasanUnit = "";

  // cari kolom & headerRow
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "string") {
        const val = cell.v.toString().toUpperCase();
        if (val.includes("KEMASAN")) colKemasan = c;
        if (val.includes("GW")) colGW = c;
        if (val.includes("NW")) colNW = c;
      }
    }
    if (colKemasan !== null && colGW !== null && colNW !== null) {
      headerRow = r;
      break;
    }
  }

  function extractUnit(text) {
    if (!text) return "";
    const s = String(text).trim();
    const m = s.match(/KEMASAN\s*(.*)/i);
    if (m && m[1] && m[1].trim()) return m[1].trim().toUpperCase();
    const matches = s.match(/[A-Za-z()\/\-\s]{2,}/g);
    if (!matches) return "";
    let candidate = matches[matches.length - 1].trim();
    candidate = candidate.replace(/^QTY\s*/i, "");
    candidate = candidate.replace(/^\(/, "").replace(/\)$/, "");
    return candidate.toUpperCase();
  }

  if (colKemasan !== null && headerRow !== null) {
    const headerCell = getCellValueRC(sheet, headerRow, colKemasan);
    kemasanUnit = extractUnit(headerCell);
    if (!kemasanUnit) {
      const below =
        headerRow + 1 <= range.e.r
          ? getCellValueRC(sheet, headerRow + 1, colKemasan)
          : "";
      if (below && typeof below === "string" && !/\d/.test(String(below))) {
        kemasanUnit = extractUnit(below);
      }
    }
    if (!kemasanUnit) {
      const above =
        headerRow - 1 >= range.s.r
          ? getCellValueRC(sheet, headerRow - 1, colKemasan)
          : "";
      if (above && typeof above === "string" && !/\d/.test(String(above))) {
        kemasanUnit = extractUnit(above);
      }
    }
  }

  // cari dataStartRow
  let dataStartRow = headerRow !== null ? headerRow + 1 : range.s.r;
  let foundDataStart = false;
  for (let rr = dataStartRow; rr <= range.e.r; rr++) {
    const serial = getCellValueRC(sheet, rr, 0); // kolom A -> c=0
    if (serial !== "" && !isNaN(Number(serial))) {
      dataStartRow = rr;
      foundDataStart = true;
      break;
    }
  }
  if (!foundDataStart) {
    dataStartRow = headerRow !== null ? headerRow + 1 : range.s.r + 1;
  }

  // akumulasi totals dari dataStartRow ke bawah
  let totalKemasan = 0,
    totalGW = 0,
    totalNW = 0;
  if (colKemasan !== null && colGW !== null && colNW !== null) {
    for (let r = dataStartRow; r <= range.e.r; r++) {
      const serial = getCellValueRC(sheet, r, 0);
      if (serial === "" || isNaN(Number(serial))) {
        continue;
      }
      const kemVal = parseInt(getCellValueRC(sheet, r, colKemasan)) || 0;
      const gwVal = parseFloat(getCellValueRC(sheet, r, colGW)) || 0;
      const nwVal = parseFloat(getCellValueRC(sheet, r, colNW)) || 0;

      totalKemasan += kemVal;
      totalGW += gwVal;
      totalNW += nwVal;
    }
  }

  return {
    kemasanSum: totalKemasan,
    bruttoSum: totalGW,
    nettoSum: totalNW,
    kemasanUnit: kemasanUnit,
  };
}
