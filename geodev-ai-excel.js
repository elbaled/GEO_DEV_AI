/* ============================================================================
   GeoDev Survey & GIS Excel Workstation (V4.0 PRO - Smart AI Command Engine)
   - تنفيذ فوري للأوامر باللهجة المصرية: "بينلي المنسوب بس"، "احذف الأعمدة الفاضية"،
     "هات كود MNH بس"، "اجمع 0.5 على Z"، "احسب المساحة"، "احسب المسافة والانحراف"
   - لوحة CAD تفاعلية بالكامل (Zoom / Pan باللمس والماوس)
   - تصدير Excel نظيف بدون سطور عناوين (PENZD / PNZ / PNEZ) + DXF + Surfer + ArcGIS + SDR
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevAIExcelV4Init) return;
  window.GeoDevAIExcelV4Init = true;

  function ensureSheetJS(callback) {
    if (window.XLSX) {
      if (callback) callback();
      return;
    }
    const script = document.createElement("script");
    script.src = "https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js";
    script.onload = () => { if (callback) callback(); };
    document.head.appendChild(script);
  }

  let columns = ["P", "E", "N", "Z", "Code"];
  let rows = [
    { P: "1", E: 312450.120, N: 3110210.450, Z: 24.510, Code: "BM1" },
    { P: "2", E: 312475.840, N: 3110225.180, Z: 24.680, Code: "RD" },
    { P: "3", E: 312510.300, N: 3110205.900, Z: 24.820, Code: "RD" },
    { P: "4", E: 312495.600, N: 3110175.400, Z: 24.390, Code: "MANHOLE" },
    { P: "5", E: 312455.200, N: 3110180.100, Z: 24.450, Code: "CORNER" }
  ];

  let historyStack = [];
  let activeSubTab = "grid";
  let crsMode = "UTM36N";
  let showLabelsOnCanvas = true;
  let connectLinesOnCanvas = true;
  let searchFilter = "";

  let cadZoom = 1;
  let cadPanX = 0;
  let cadPanY = 0;
  let isDraggingCanvas = false;
  let dragStartX = 0;
  let dragStartY = 0;
  let initialPinchDist = 0;
  let initialPinchZoom = 1;

  function resetCadView() {
    cadZoom = 1;
    cadPanX = 0;
    cadPanY = 0;
    drawCadPreview();
  }

  function saveSnapshot() {
    historyStack.push(JSON.stringify({ columns, rows }));
    if (historyStack.length > 35) historyStack.shift();
  }

  function undoLastAction() {
    if (!historyStack.length) {
      showExcelLog("لا توجد خطوات سابقة للتراجع عنها.", "warn");
      return;
    }
    const prev = JSON.parse(historyStack.pop());
    columns = prev.columns;
    rows = prev.rows;
    renderExcelTable();
    drawCadPreview();
    showExcelLog("↩️ تم التراجع واستعادة البيانات السابقة بنجاح.", "success");
  }

  function showExcelLog(msg, type = "info") {
    const box = document.getElementById("gdExcelAiLog");
    if (!box) return;
    const palette = {
      success: { bg: "rgba(22,163,74,0.16)", border: "#22c55e", text: "#4ade80" },
      warn: { bg: "rgba(234,179,8,0.16)", border: "#eab308", text: "#facc15" },
      error: { bg: "rgba(220,38,38,0.16)", border: "#ef4444", text: "#f87171" },
      info: { bg: "rgba(56,189,248,0.14)", border: "#38bdf8", text: "#38bdf8" }
    };
    const c = palette[type] || palette.info;
    box.style.background = c.bg;
    box.style.borderColor = c.border;
    box.style.color = c.text;
    box.textContent = msg;
  }

  function injectStyles() {
    const old = document.getElementById("gdAiExcelStyles");
    if (old) old.remove();

    const style = document.createElement("style");
    style.id = "gdAiExcelStyles";
    style.textContent = `
      #gdAiExcelStudio {
        direction: rtl;
        text-align: right;
        background: linear-gradient(165deg, #141210, #1e1b18);
        border: 2px solid #e09f24;
        border-radius: 20px;
        padding: 16px;
        margin: 18px 0;
        color: #f5f5f4;
        box-shadow: 0 18px 44px rgba(0, 0, 0, 0.5);
      }
      .gd-ex-top-bar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 10px;
        padding-bottom: 12px;
        border-bottom: 1px solid #38332e;
        margin-bottom: 14px;
      }
      .gd-ex-top-bar h3 {
        margin: 0;
        font-size: 18.5px;
        font-weight: 900;
        color: #fbbf24;
      }
      .gd-ex-top-bar p {
        margin: 4px 0 0;
        font-size: 12px;
        color: #a8a29e;
      }
      .gd-ex-pro-pill {
        background: linear-gradient(135deg, #e09f24, #b45309);
        color: #141210;
        font-weight: 900;
        font-size: 11px;
        padding: 5px 12px;
        border-radius: 999px;
        letter-spacing: 0.5px;
      }

      .gd-ex-ai-panel {
        background: #100e0c;
        border: 1.5px solid #38bdf8;
        border-radius: 15px;
        padding: 13px;
        margin-bottom: 14px;
      }
      .gd-ex-ai-row {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
      }
      .gd-ex-input {
        flex: 1;
        min-width: 210px;
        padding: 10px 12px;
        border-radius: 10px;
        border: 1.5px solid #3f3a36;
        background: #211e1b;
        color: #ffffff;
        font-size: 13px;
        font-weight: 700;
        font-family: inherit;
      }
      .gd-ex-input:focus {
        outline: none;
        border-color: #fbbf24;
      }
      .gd-ex-ai-btn {
        padding: 10px 16px;
        border-radius: 10px;
        border: 0;
        background: linear-gradient(135deg, #0284c7, #1d4ed8);
        color: #ffffff;
        font-size: 13px;
        font-weight: 900;
        cursor: pointer;
        font-family: inherit;
      }
      .gd-ex-chips {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin-top: 8px;
      }
      .gd-ex-chip {
        background: #26221f;
        border: 1px solid #44403c;
        color: #e7e5e4;
        padding: 5px 9px;
        border-radius: 8px;
        font-size: 11px;
        font-weight: 700;
        cursor: pointer;
        font-family: inherit;
      }
      .gd-ex-chip:hover {
        border-color: #fbbf24;
        color: #fbbf24;
      }
      #gdExcelAiLog {
        margin-top: 9px;
        padding: 8px 12px;
        border-radius: 9px;
        border: 1px solid #38bdf8;
        background: rgba(56, 189, 248, 0.12);
        color: #38bdf8;
        font-size: 12px;
        font-weight: 800;
        line-height: 1.6;
      }

      .gd-ex-mode-tabs {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 8px;
        margin-bottom: 13px;
      }
      .gd-ex-mode-tab {
        padding: 10px 8px;
        border-radius: 11px;
        border: 1.5px solid #3f3a36;
        background: #24201d;
        color: #d6d3d1;
        font-size: 12.5px;
        font-weight: 800;
        cursor: pointer;
        font-family: inherit;
        text-align: center;
      }
      .gd-ex-mode-tab.active {
        background: rgba(224, 159, 36, 0.18);
        border-color: #e09f24;
        color: #fbbf24;
      }

      .gd-ex-toolbar {
        display: flex;
        flex-wrap: wrap;
        gap: 7px;
        margin-bottom: 12px;
        align-items: center;
      }
      .gd-ex-btn {
        padding: 8px 11px;
        border-radius: 10px;
        border: 1px solid #44403c;
        background: #292524;
        color: #f5f5f4;
        font-size: 11.5px;
        font-weight: 800;
        cursor: pointer;
        font-family: inherit;
        display: inline-flex;
        align-items: center;
        gap: 5px;
      }
      .gd-ex-btn:hover {
        border-color: #e09f24;
        color: #fbbf24;
      }
      .gd-ex-btn.gold {
        background: linear-gradient(135deg, #e09f24, #b45309);
        color: #141210;
        border: 0;
        font-weight: 900;
      }
      .gd-ex-btn.danger {
        background: rgba(153, 27, 27, 0.35);
        border-color: #ef4444;
        color: #fca5a5;
      }

      .gd-ex-stats {
        display
