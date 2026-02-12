/* ------------------------------------------------
共通帳票エンジン

帳票固有処理なし

rowspan / colspan 完全対応

JSON駆動
--------------------------------------------------- */
let currentCss = null;
let currentData = null;

window.PrintFromUnity = async function (json, template, fileName) {
  // console.log("PrintFromUnity called");

  /* ---------- 1. HTMLテンプレ読込 ---------- */
  const html = await fetch(template + '.html').then(r => r.text());
  document.getElementById('report-root').innerHTML = html;

  if (currentCss) {
    document.head.removeChild(currentCss);
  }  

  currentCss = document.createElement('link');
  currentCss.rel = 'stylesheet';
  currentCss.href = template + '.css';
  document.head.appendChild(currentCss);

  // 表示切替
  document.body.classList.add("printing");

  /* ---------- 2. DOM生成待ち ---------- */
  /* await waitDomReady();*/
  /* const data = JSON.parse(json);*/
  const data = JSON.parse(json);
  currentData = data;

  /* ---------- 3. 帳票のリストを単位でページ生成 ---------- */
  data.reportList.forEach((report, index) => {

    /* sheet生成（＝1ページ） */
    const sheet = document.createElement("div");
    sheet.className = "sheet";

    const root = document.getElementById('report-root');
    root.appendChild(sheet);

    /* HTMLテンプレを流し込む */
    sheet.innerHTML = html;

    /* テキスト反映 */
    setTextIfExists(sheet, "report-title", report.title);
    setTextIfExists(sheet, "customerName", report.customerName);

    /* テーブル生成 */
    if (report.tables) {
      report.tables.forEach(t => {
        buildTable(sheet, t.tbodyId, t.rows);
      });
    }
  });

  /* ---------- 4. 印刷 ---------- */
  document.getElementById('report-root').style.display = 'block';
  beforePrint();

  /* PDFファイル名を指定 */
  const oldTitle = document.title;
  document.title = fileName ?? "帳票";

  window.print();

  /* 印刷後にファイル名を戻す */
  document.title = oldTitle;
};




/* ===============================
  DOM生成待ち
=============================== */
function waitDomReady() {
  return new Promise(resolve => {
   const check = () => {
    if (document.body && document.body.children.length > 0) {
      resolve();
    } else {
      setTimeout(check, 10);
    }
   };
   check();
  });
}

/* ===============================
 テキスト反映（存在チェック付）
=============================== */
function setTextIfExists(root, id, value) {
  if (!value) return;
  const el = root.querySelector("#" + id);
  if (el) el.innerText = value;
}

/* ===============================
 汎用テーブルビルダー
=============================== */
function buildTable(root, tbodyId, rows)
{
  const tbody = root.querySelector("#" + tbodyId);
  if (!tbody) return;

  rows.forEach(row => {
    const tr = document.createElement("tr");

    row.cells.forEach(cell => {
      if (cell.skip) return;

      const td = document.createElement("td");
      td.innerText = cell.value ?? "";

      if (cell.rowspan) td.rowSpan = cell.rowspan;
      if (cell.colspan) td.colSpan = cell.colspan;
      if (cell.class) td.className = cell.class;

      if (cell.style) {
        Object.keys(cell.style).forEach(key => {
          td.style[key] = cell.style[key];
        });
      }

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
    console.log("BuildTableCount");
  });
}


