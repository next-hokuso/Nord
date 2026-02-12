function exportCsv(json, fileName) {

  const data = JSON.parse(json);

  // ヘッダ行
  let headerCells = [];
  for (let i = 0; i < data.reportList[0].columnsTop.length; i++) {
    const col = data.reportList[0].columnsTop[i];
    headerCells.push(`"${col.label ?? ""}"`);
  }

  let lines = [];
  lines.push(headerCells.join(","));

  // 行データ
  for (let t = 0; t < data.reportList[0].tables.length; t++) {
    const table = data.reportList[0].tables[t];

    for (let r = 0; r < table.rows.length; r++) {
      const row = table.rows[r];

      let rowCells = [];
      for (let c = 0; c < row.cells.length; c++) {
        const cell = row.cells[c];
        rowCells.push(`"${cell.value ?? ""}"`);
      }

      lines.push(rowCells.join(","));
    }
  }

  // CRLF 改行
  const csvBody = lines.join("\r\n");

  // UTF-8 BOM
  const bom = "\uFEFF";
  const csv = bom + csvBody;

  download(csv, `${fileName}.csv`, "text/csv;charset=utf-8;");
}


function download(content, fileName, mime) {
  const blob = new Blob([content], { type: mime });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(a.href);
}
