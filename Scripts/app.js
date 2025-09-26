// Event bindings
document.getElementById("btnCheck").addEventListener("click", processFiles);
document.getElementById("ptSelect").addEventListener("change", updateKontrak);
document.getElementById("filter").addEventListener("change", applyFilter);

function updateKontrak() {
  const idx = document.getElementById("ptSelect").value;
  if (idx === "") {
    document.getElementById("noKontrak").value = "";
    document.getElementById("tglKontrak").value = "";
    return;
  }
  document.getElementById("noKontrak").value = kontrakData[idx].no;
  document.getElementById("tglKontrak").value = kontrakData[idx].tgl;
}

function processFiles() {
  const files = document.getElementById("files").files;
  if (files.length !== 3) {
    alert("Harap upload tepat 3 file (Draft, INV, PL)");
    return;
  }

  let sheetPL, sheetINV, sheetsDATA;
  Promise.all([...files].map((f) => readExcelFile(f))).then((workbooks) => {
    workbooks.forEach((wb) => {
      const type = detectFileType(wb);
      if (type === "PL") sheetPL = wb.Sheets[wb.SheetNames[0]];
      if (type === "INV") sheetINV = wb.Sheets[wb.SheetNames[0]];
      if (type === "DATA") {
        // asumsi sheet names di file DATA ada HEADER, DOKUMEN, KEMASAN, BARANG
        sheetsDATA = {
          HEADER: wb.Sheets["HEADER"],
          DOKUMEN: wb.Sheets["DOKUMEN"],
          KEMASAN: wb.Sheets["KEMASAN"],
          BARANG: wb.Sheets["BARANG"],
        };
      }
    });

    if (!sheetPL || !sheetINV || !sheetsDATA) {
      alert("Tidak bisa mendeteksi DRAFT / INV / PL, harap periksa file anda.");
      return;
    }

    // parsing kurs agar input type=number tidak error
    const kursCell = getCellValue(sheetsDATA.HEADER, "BW2");
    const kursParsed = parseKurs(kursCell) || 1;
    document.getElementById("kurs").value = kursParsed;

    checkAll(sheetPL, sheetINV, sheetsDATA, kursParsed);
  });
}
