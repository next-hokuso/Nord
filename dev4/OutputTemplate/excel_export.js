window.ExportExcelFromUnity = function (json, fileName) {
  //console.log("=== ExportExcel called ===");

  const data = JSON.parse(json);
  //console.log("parsed data:", data);

  const ws_data = [];
  const merges = [];

  let colmax = 0;

  const wb = XLSX.utils.book_new();


  data.reportList[0].tables.forEach(table => {
  table.rows.forEach((row, rowIndex) => {
    //console.log(`row[${rowIndex}]`, row);

    const rowArray = [];

    row.cells.forEach((cell, colIndex) => {
      //console.log(` cell r:${rowIndex} c:${colIndex}`, cell);

        // 合体セルのスキップ
        if (cell.skip) {
          //colIndex++;
          //return;
        }

      // セル設定
      rowArray[colIndex] =
        cell.value !== undefined && cell.value !== null
          ? cell.value
          : "";

      // 行合体設定(縦)
      if (cell.rowspan && cell.rowspan > 1) {
        merges.push({
          s: { r: rowIndex, c: colIndex },
          e: { r: rowIndex + cell.rowspan - 1, c: colIndex }
        });
      }
      // 列合体設定(横)
      if (cell.colspan && cell.colspan > 1) {
        merges.push({
          s: { r: rowIndex, c: colIndex },
          e: { r: rowIndex, c: colIndex + cell.colspan - 1 }
        });
      }

       // 列数の保存
       const span = cell.colspan || 1;
       const end = colIndex + span;
       if (end > colmax) colmax = end;
    });

    //console.log("rowArray:", rowArray);
    ws_data.push(rowArray);
  });
});

  const ws = XLSX.utils.aoa_to_sheet(ws_data);

  data.reportList[0].tables.forEach(table => {
  table.rows.forEach((row, rowIndex) => {
    row.cells.forEach((cell, colIndex) => {

        // 合体セルのスキップ
        if (cell.skip) {
          colIndex++;
          return;
        }

      // セル設定 align
      applyAlignment(ws, rowIndex, colIndex, cell);
    });
  });
});


  //console.log("final ws_data:", ws_data);
  //console.log("final merges:", merges);

  // セル合体
  ws["!merges"] = merges;

  // セル列サイズ設定
  //ws["!cols"] = data.reportList[0].columnsTop.map(c => ({ wch: mmToWch(c.widthExcel) }));
  ws["!cols"] = data.reportList[0].columnsTop.map(c => ({ wch: c.widthExcel }));

  //--- 罫線適用 ------
// 明細範囲
const startRow = 1;                  // 2 行
const endRow   = ws_data.length - 1; // 最終行
const startCol = 1;                  // B列
const endCol   = colmax;             // 最後の列

  // 合体セルの罫線を引くのに必須らしい
  const range = XLSX.utils.encode_range({
    s: { r:startRow, c:startCol },
    e: { r: endRow +1, c:endCol+1 }
  });
  ws["!ref"] = range;


  //console.log("endRow :", endRow);
  //console.log("endCol :", endCol);

// ① 全体に細罫線
applyBorder(ws, startRow, endRow, startCol, endCol);

// ② 外枠を太線
applyOuterBorder(ws, startRow, endRow, startCol, endCol);


// ③ ヘッダー行の下線を太線
const headerRow = 0;
applyHeaderBorder(ws, headerRow, startCol, endCol);
applyHeaderBorder(ws, headerRow + 1, startCol, endCol);

// ４合体セル 各セルで引く必要があるためループ
if(ws["merges"]){
 ws["!merges"].forEach(m => {
   applyMergeWithBorder(
     ws, m
   );
 });
}
debugMergeBorders(ws);
  //--- ↑ここまで

  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, `${fileName}.xlsx`);
};



// mm の wch 変換 完全一致は厳しいため近似値を設定
function mmToWch(mm) {
  const px = mm * 96 / 25.4;
  const excelWidth = (px - 5) / 7;

  return excelWidth;
  // return Math.round(mm * 0.35);
}





// 罫線の処理
function setBorder(ws, r, c, side, style = "thin") {
  const addr = XLSX.utils.encode_cell({ r, c });
  if (!ws[addr]) return;

  ws[addr].s = ws[addr].s || {};
  ws[addr].s.border = ws[addr].s.border || {};
  ws[addr].s.border[side] = { style };
}


/* テーブル全体に罫線を引く */
function applyBorder(ws, startRow, endRow, startCol, endCol) {
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (ws[addr]){
       ws[addr].s = ws[addr].s || {};
       ws[addr].s.border = {
         top:    { style: "thin" },
         bottom: { style: "thin" },
         left:   { style: "thin" },
         right:  { style: "thin" }
       };
       //console.log("r", r, "c", c);
      }
    }
  }
}


/* 外枠だけ太くする（帳票必須) */
function applyOuterBorder(ws, startRow, endRow, startCol, endCol) {
 for (let c = startCol; c <= endCol; c++) {
  setBorder(ws, startRow, c, "top", "medium");
  setBorder(ws, endRow, c, "bottom", "medium");
 }
 for (let r = startRow; r <= endRow; r++) {
  setBorder(ws, r, startCol, "left", "medium");
  setBorder(ws, r, endCol, "right", "medium");
 }
}



/* ヘッダー行だけ太罫線 */
function applyHeaderBorder(ws, headerRow, startCol, endCol) {
 for (let c = startCol; c <= endCol; c++) {
  setBorder(ws, headerRow, c, "top", "medium");
  setBorder(ws, headerRow, c, "bottom", "medium");
 }
}

// merge セル外周用
function applyMergedBorder(ws, startRow, startCol, endRow, endCol) {
 for (let c = startCol; c <= endCol; c++) {
  setBorder(ws, startRow, c, "top", "medium");
  setBorder(ws, endRow, c, "bottom", "medium");
 }

 for (let r = startRow; r <= endRow; r++) {
  setBorder(ws, r, startCol, "left", "medium");
  setBorder(ws, r, endCol, "right", "medium");
 }
}
// merge セル外周用
function applyMergeWithBorder(ws, merge, style = "thin") {
  const { s, e } = merge;

  // 上下
  for (let c = s.c; c <= e.c; c++) {
    setBorder(ws, s.r, c, "top", style);
    setBorder(ws, e.r, c, "bottom", style);
  }

  // 左右
  for (let r = s.r; r <= e.r; r++) {
    setBorder(ws, r, s.c, "left", style);
    setBorder(ws, r, e.c, "right", style);
  }
}






// 罫線デバッグ
function debugMergeBorders(ws) {
  if (!ws["!merges"]) return;

  ws["!merges"].forEach((m, i) => {
    const missing = [];

    for (let c = m.s.c; c <= m.e.c; c++) {
      if (!hasBorder(ws, m.s.r, c, "top")) missing.push(`top r${m.s.r} c${c}`);
      if (!hasBorder(ws, m.e.r, c, "bottom")) missing.push(`bottom r${m.e.r} c${c}`);
    }

    for (let r = m.s.r; r <= m.e.r; r++) {
      if (!hasBorder(ws, r, m.s.c, "left")) missing.push(`left r${r} c${m.s.c}`);
      if (!hasBorder(ws, r, m.e.c, "right")) missing.push(`right r${r} c${m.e.c}`);
    }

    if (missing.length > 0) {
      console.warn(`⚠ merge[${i}] border missing:`, missing);
    }
  });
}

function hasBorder(ws, r, c, side) {
  const addr = XLSX.utils.encode_cell({ r, c });
  return ws[addr]?.s?.border?.[side];
}



// Alignment
function applyAlignment(ws, r, c, cell) {
  const addr = XLSX.utils.encode_cell({ r, c });
  if(!ws[addr]) return;
  ws[addr].s = ws[addr].s || {};
  ws[addr].s.alignment = ws[addr].s.alignment || {};

  if (cell.align) {
    ws[addr].s.alignment.horizontal = cell.align;
  }

  if (cell.valign) {
    ws[addr].s.alignment.vertical =
      cell.valign === "middle" ? "center" : cell.valign;
  }
}



