/* ============================================================================
   GeoDev Survey & GIS Excel Workstation (V2.0 PRO)
   - تكامل كامل مع: AutoCAD (DXF Import/Export, SCR, LISP), Civil 3D,
     Golden Software Surfer (.dat/.bln), ArcGIS Pro & QGIS (.geojson / WKT CSV),
     Google Earth (.kml), Total Station (Sokkia .sdr/.cdr & Leica .gsi), Excel (.xlsx)
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevAIExcelV2Init) return;
  window.GeoDevAIExcelV2Init = true;

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
  let activeSubTab = "grid"; // "grid" | "cad_import" | "calc_tools"
  let crsMode = "UTM36N";    // "UTM36N" | "UTM35N" | "WGS84"
  let showLabelsOnCanvas = true;
  let connectLinesOnCanvas = true;
  let searchFilter = "";

  function saveSnapshot() {
    historyStack.push(JSON.stringify({ columns, rows }));
    if (historyStack.length > 30) historyStack.shift();
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
    showExcelLog("تم التراجع عن آخر عملية بنجاح.", "success");
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

      /* شريط الذكاء الاصطناعي */
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

      /* تبويبات بيئة العمل */
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

      /* شريط الاستيراد والأدوات */
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

      /* بطاقات الإحصائيات الهندسية */
      .gd-ex-stats {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(115px, 1fr));
        gap: 7px;
        margin-bottom: 12px;
      }
      .gd-ex-stat {
        background: #100e0c;
        border: 1px solid #2e2a26;
        border-radius: 10px;
        padding: 8px;
        text-align: center;
      }
      .gd-ex-stat span {
        display: block;
        font-size: 10px;
        color: #a8a29e;
      }
      .gd-ex-stat strong {
        font-size: 13px;
        color: #fbbf24;
        font-weight: 900;
      }

      /* الجدول التفاعلي */
      .gd-ex-grid-wrap {
        max-height: 340px;
        overflow: auto;
        border: 1.5px solid #44403c;
        border-radius: 12px;
        background: #0e0d0b;
        margin-bottom: 14px;
      }
      .gd-ex-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 12.5px;
        text-align: center;
      }
      .gd-ex-table th {
        position: sticky;
        top: 0;
        z-index: 5;
        background: #292524;
        color: #fbbf24;
        padding: 9px 6px;
        border: 1px solid #44403c;
        font-weight: 900;
        white-space: nowrap;
      }
      .gd-ex-table td {
        border: 1px solid #26221f;
        padding: 0;
        background: #141210;
      }
      .gd-ex-table tr:nth-child(even) td {
        background: #191614;
      }
      .gd-ex-cell {
        width: 100%;
        min-width: 76px;
        box-sizing: border-box;
        border: 0;
        background: transparent;
        color: #f5f5f4;
        padding: 8px 6px;
        text-align: center;
        font-size: 12.5px;
        font-weight: 700;
        font-family: inherit;
      }
      .gd-ex-cell:focus {
        outline: none;
        background: rgba(224, 159, 36, 0.16);
      }

      /* لوحة استيراد نص الأوتوكاد أو أدوات الحساب */
      .gd-ex-subpanel {
        background: #12100e;
        border: 1px solid #3f3a36;
        border-radius: 13px;
        padding: 13px;
        margin-bottom: 14px;
      }
      .gd-ex-textarea {
        width: 100%;
        min-height: 115px;
        box-sizing: border-box;
        background: #1c1917;
        border: 1.5px solid #44403c;
        border-radius: 10px;
        color: #f5f5f4;
        padding: 10px;
        font-family: monospace;
        font-size: 12px;
        margin: 8px 0;
        direction: ltr;
      }

      /* شاشة المعاينة CAD */
      .gd-ex-canvas-box {
        background: #06080d;
        border: 1.5px solid #334155;
        border-radius: 14px;
        padding: 10px;
        margin-bottom: 14px;
      }
      .gd-ex-canvas-head {
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 7px;
        margin-bottom: 8px;
        font-size: 12px;
        color: #38bdf8;
        font-weight: 800;
      }
      #gdExcelCadCanvas {
        width: 100%;
        height: 245px;
        display: block;
        border-radius: 10px;
        background: #04060a;
      }

      /* شبكة التصدير للبرامج الهندسية */
      .gd-ex-export-section-title {
        font-size: 13px;
        font-weight: 900;
        color: #fbbf24;
        margin: 12px 0 7px;
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .gd-ex-export-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(145px, 1fr));
        gap: 8px;
      }
      .gd-ex-exp-btn {
        padding: 10px 9px;
        border-radius: 11px;
        border: 1.5px solid #44403c;
        background: #211e1b;
        color: #f5f5f4;
        font-size: 12px;
        font-weight: 800;
        cursor: pointer;
        font-family: inherit;
        text-align: center;
        transition: all 0.2s ease;
      }
      .gd-ex-exp-btn:hover {
        border-color: #fbbf24;
        transform: translateY(-1px);
      }
      .gd-ex-exp-btn small {
        display: block;
        font-size: 10px;
        color: #a8a29e;
        margin-top: 2px;
      }
    `;
    document.head.appendChild(style);
  }

  function renderStudioHTML() {
    return `
      <div class="gd-ex-top-bar">
        <div>
          <h3>📊 محطة إكسيل المساحة ونظم المعلومات الجغرافية (AI Survey Excel)</h3>
          <p>استيراد وتصدير مباشر مع AutoCAD DXF، Surfer، ArcGIS Pro، QGIS، وأجهزة التوتال ستيشن SDR / GSI</p>
        </div>
        <span class="gd-ex-pro-pill">WORKSTATION PRO</span>
      </div>

      <!-- شريط أوامر الذكاء الاصطناعي -->
      <div class="gd-ex-ai-panel">
        <div class="gd-ex-ai-row">
          <input type="text" id="gdExcelAiPrompt" class="gd-ex-input"
            placeholder="اكتب أمرك بالعربي: اكتب عمرو 1 2 3 | اجمع 0.5 على Z | احسب المساحة | احسب المسافة والانحراف...">
          <button type="button" id="gdExcelAiExecBtn" class="gd-ex-ai-btn">⚡ تنفيذ بالـ AI</button>
        </div>
        <div class="gd-ex-chips">
          <button type="button" class="gd-ex-chip" data-ai-cmd="اكتب في ملف الإكسيل عمرو 1 2 3">✍️ اكتب عمرو 1 2 3</button>
          <button type="button" class="gd-ex-chip" data-ai-cmd="احسب المسافة والانحراف">📐 المسافة والانحراف</button>
          <button type="button" class="gd-ex-chip" data-ai-cmd="احسب المساحة والمحيط">🗺️ المساحة بالفدان والمتر</button>
          <button type="button" class="gd-ex-chip" data-ai-cmd="اجمع 0.50 على Z">➕ إزاحة Z (+0.50)</button>
          <button type="button" class="gd-ex-chip" data-ai-cmd="بدل الشرق والشمال">🔄 تبديل E ⇄ N</button>
          <button type="button" class="gd-ex-chip" data-ai-cmd="احذف النقاط المكررة">🧹 حذف المكرر</button>
        </div>
        <div id="gdExcelAiLog">جاهز لاستيراد ملفات (XLSX, DXF, SDR, CDR, GSI, CSV) أو تصديرها مباشرة للأوتوكاد و ArcGIS Pro و Surfer.</div>
      </div>

      <!-- شريط الاستيراد الموحد -->
      <div class="gd-ex-toolbar">
        <label class="gd-ex-btn gold" style="cursor:pointer;">
          📂 استيراد ملف (Excel / DXF / SDR / CDR / GSI / CSV)
          <input type="file" id="gdExcelFileInput" accept=".xlsx,.xls,.csv,.txt,.dxf,.sdr,.cdr,.gsi" style="display:none;">
        </label>
        <button type="button" class="gd-ex-btn" id="gdExAddRowBtn">➕ نقطة جديدة</button>
        <button type="button" class="gd-ex-btn" id="gdExAddColBtn">📑 عمود جديد</button>
        <button type="button" class="gd-ex-btn" id="gdExSwapENBtn">🔄 E ⇄ N</button>
        <button type="button" class="gd-ex-btn" id="gdExUndoBtn">↩️ تراجع</button>
        <button type="button" class="gd-ex-btn danger" id="gdExClearBtn">🗑️ تفريغ</button>
      </div>

      <!-- تبويبات بيئة العمل -->
      <div class="gd-ex-mode-tabs">
        <button type="button" class="gd-ex-mode-tab active" data-ex-subtab="grid">📋 جدول الإكسيل والبيانات</button>
        <button type="button" class="gd-ex-mode-tab" data-ex-subtab="cad_import">📐 لصق من الأوتوكاد (LIST)</button>
        <button type="button" class="gd-ex-mode-tab" data-ex-subtab="calc_tools">🧮 إزاحة وحفر وردم (Cut/Fill)</button>
      </div>

      <!-- لوحة لصق إحداثيات من الأوتوكاد مباشرة -->
      <div id="gdExSubpanelCadImport" class="gd-ex-subpanel" style="display:none;">
        <strong style="color:#fbbf24;font-size:13px;">استخراج الإحداثيات من الأوتوكاد عبر أمر LIST أو نسخ النصوص:</strong>
        <p style="margin:4px 0;font-size:11.5px;color:#a8a29e;">
          في الأوتوكاد حدد النقاط أو البولي لاين واكتب أمر <code>LIST</code> ثم انسخ النص الظاهر في نافذة الأوامر والصقه هنا، وسيستخرج النظام كل إحداثيات <code>X, Y, Z</code> تلقائياً:
        </p>
        <textarea id="gdAcadListPasteArea" class="gd-ex-textarea" placeholder="at point  X= 312450.120  Y= 3110210.450  Z= 24.510&#10;at point  X= 312475.840  Y= 3110225.180  Z= 24.680"></textarea>
        <button type="button" class="gd-ex-btn gold" id="gdParseAcadListBtn" style="width:100%;justify-content:center;">
          ⚡ استخراج الإحداثيات وإضافتها لجدول الإكسيل الآن
        </button>
      </div>

      <!-- لوحة الأدوات الهندسية (إزاحة الإحداثيات + حساب الحفر والردم) -->
      <div id="gdExSubpanelCalcTools" class="gd-ex-subpanel" style="display:none;">
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;">
          <div style="background:#1c1917;padding:10px;border-radius:10px;border:1px solid #38332e;">
            <strong style="color:#38bdf8;font-size:12.5px;">1. إزاحة الإحداثيات (Coordinate Shift / Datum)</strong>
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin:8px 0;">
              <input type="number" step="any" id="gdShiftE" class="gd-ex-input" style="min-width:0;text-align:center;" placeholder="ΔE (0)">
              <input type="number" step="any" id="gdShiftN" class="gd-ex-input" style="min-width:0;text-align:center;" placeholder="ΔN (0)">
              <input type="number" step="any" id="gdShiftZ" class="gd-ex-input" style="min-width:0;text-align:center;" placeholder="ΔZ (0)">
            </div>
            <button type="button" class="gd-ex-btn" id="gdApplyShiftBtn" style="width:100%;justify-content:center;">تطبيق الإزاحة على النقاط</button>
          </div>

          <div style="background:#1c1917;padding:10px;border-radius:10px;border:1px solid #38332e;">
            <strong style="color:#4ade80;font-size:12.5px;">2. حساب الحفر والردم (Cut &amp; Fill vs Design)</strong>
            <div style="margin:8px 0;">
              <input type="number" step="any" id="gdDesignLevelInput" class="gd-ex-input" style="width:100%;text-align:center;" placeholder="أدخل منسوب التصميم (مثال: 24.50)">
            </div>
            <button type="button" class="gd-ex-btn gold" id="gdCalcCutFillBtn" style="width:100%;justify-content:center;">حساب عمود Cut / Fill</button>
          </div>
        </div>
      </div>

      <!-- الإحصائيات المساحية الحية -->
      <div class="gd-ex-stats" id="gdExStatsBar"></div>

      <!-- جدول الإكسيل -->
      <div class="gd-ex-grid-wrap">
        <table class="gd-ex-table" id="gdExcelDataTable"></table>
      </div>

      <!-- شاشة المعاينة الهندسية -->
      <div class="gd-ex-canvas-box">
        <div class="gd-ex-canvas-head">
          <span>🖥️ شاشة المعاينة الهندسية الحية (Live CAD / GIS Plot)</span>
          <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;">
            <input type="text" id="gdExSearchInput" class="gd-ex-input" style="max-width:130px;padding:4px 8px;font-size:11px;" placeholder="🔍 بحث...">
            <button type="button" class="gd-ex-btn" style="padding:4px 8px;font-size:11px;" id="gdToggleLabelsBtn">🏷️ النصوص</button>
            <button type="button" class="gd-ex-btn" style="padding:4px 8px;font-size:11px;" id="gdToggleLinesBtn">📏 الخطوط</button>
            <select id="gdCrsSelect" class="gd-ex-btn" style="padding:4px 8px;font-size:11px;">
              <option value="UTM36N" selected>UTM Zone 36N (مصر)</option>
              <option value="UTM35N">UTM Zone 35N (غرب مصر)</option>
              <option value="WGS84">WGS84 (Lat / Lon)</option>
            </select>
          </div>
        </div>
        <canvas id="gdExcelCadCanvas"></canvas>
      </div>

      <!-- قسم التصدير لبرامج CAD & Surfer -->
      <div class="gd-ex-export-section-title">📐 تصدير مباشر لبرامج الأوتوكاد والسيرفر (CAD &amp; Surfer):</div>
      <div class="gd-ex-export-grid">
        <button type="button" class="gd-ex-exp-btn" data-export-fmt="dxf" style="border-color:#ef4444;">
          📕 لوحة AutoCAD (.dxf)
          <small>تفتح مباشرة في AutoCAD بطبقات</small>
        </button>
        <button type="button" class="gd-ex-exp-btn" data-export-fmt="csv" style="border-color:#38bdf8;">
          📘 Civil 3D (.csv)
          <small>صيغة PENZD جاهزة للاستيراد</small>
        </button>
        <button type="button" class="gd-ex-exp-btn" data-export-fmt="surfer" style="border-color:#f59e0b;">
          🏔️ برنامج Surfer (.dat)
          <small>XYZ جاهز للكنتور والكميات</small>
        </button>
        <button type="button" class="gd-ex-exp-btn" data-export-fmt="lsp" style="border-color:#a855f7;">
          ⚡ AutoCAD LISP (.lsp)
          <small>أمر GD_PLOT_PTS</small>
        </button>
      </div>

      <!-- قسم التصدير لبرامج GIS وأجهزة التوتال ستيشن -->
      <div class="gd-ex-export-section-title">🌍 تصدير مباشر لـ ArcGIS Pro و QGIS وأجهزة التوتال ستيشن:</div>
      <div class="gd-ex-export-grid">
        <button type="button" class="gd-ex-exp-btn" data-export-fmt="geojson" style="border-color:#10b981;">
          🗺️ ArcGIS Pro / QGIS (.geojson)
          <small>طبقة جغرافية ثلاثية الأبعاد</small>
        </button>
        <button type="button" class="gd-ex-exp-btn" data-export-fmt="wkt" style="border-color:#06b6d4;">
          📊 ArcGIS Pro WKT (.csv)
          <small>جدول بحقل POINT Z جاهز</small>
        </button>
        <button type="button" class="gd-ex-exp-btn" data-export-fmt="kml" style="border-color:#f97316;">
          🛰️ Google Earth (.kml)
          <small>عرض النقاط على القمر الصناعي</small>
        </button>
        <button type="button" class="gd-ex-exp-btn" data-export-fmt="sdr" style="border-color:#eab308;">
          📡 Sokkia SDR / CDR (.sdr)
          <small>للتوتال ستيشن (التوقيع)</small>
        </button>
        <button type="button" class="gd-ex-exp-btn" data-export-fmt="gsi" style="border-color:#84cc16;">
          📡 Leica GSI (.gsi)
          <small>ملف جهاز لايكا جاهز</small>
        </button>
        <button type="button" class="gd-ex-exp-btn" data-export-fmt="xlsx" style="border-color:#22c55e;">
          📗 ملف Excel (.xlsx)
          <small>جدول إكسيل منسق كامل</small>
        </button>
      </div>
    `;
  }

  /* ============================================================================
     دوال التحليل الهندسي وتحويل الإحداثيات
     ============================================================================ */
  function getCoordKeys() {
    const pKey = columns.find(c => /^(p|pt|point|id|no|num|رقم|النقطة)$/i.test(c.trim())) || columns[0];
    const eKey = columns.find(c => /^(e|east|easting|x|lon|long|شرق)$/i.test(c.trim())) || columns[1];
    const nKey = columns.find(c => /^(n|north|northing|y|lat|شمال)$/i.test(c.trim())) || columns[2];
    const zKey = columns.find(c => /^(z|elev|elevation|rl|level|منسوب)$/i.test(c.trim())) || columns[3];
    const cKey = columns.find(c => /^(code|desc|description|note|كود|وصف)$/i.test(c.trim())) || columns[4];
    return { pKey, eKey, nKey, zKey, cKey };
  }

  function computeSurveyStats() {
    const { eKey, nKey, zKey } = getCoordKeys();
    const validPts = rows
      .map(r => ({
        e: parseFloat(r[eKey]),
        n: parseFloat(r[nKey]),
        z: parseFloat(r[zKey])
      }))
      .filter(p => !isNaN(p.e) && !isNaN(p.n));

    if (!validPts.length) {
      return { count: rows.length, minZ: "-", maxZ: "-", perimeter: 0, areaM2: 0 };
    }

    const zVals = validPts.map(p => p.z).filter(v => !isNaN(v));
    const minZ = zVals.length ? Math.min(...zVals).toFixed(3) : "-";
    const maxZ = zVals.length ? Math.max(...zVals).toFixed(3) : "-";

    let perimeter = 0;
    for (let i = 0; i < validPts.length - 1; i++) {
      perimeter += Math.hypot(validPts[i + 1].e - validPts[i].e, validPts[i + 1].n - validPts[i].n);
    }

    let area2 = 0;
    if (validPts.length >= 3) {
      for (let i = 0; i < validPts.length; i++) {
        const j = (i + 1) % validPts.length;
        area2 += validPts[i].e * validPts[j].n - validPts[j].e * validPts[i].n;
      }
    }
    return { count: rows.length, minZ, maxZ, perimeter, areaM2: Math.abs(area2) / 2 };
  }

  function projectToWGS84(easting, northing) {
    if (crsMode === "WGS84" || (Math.abs(easting) <= 180 && Math.abs(northing) <= 90)) {
      return { lon: easting, lat: northing };
    }
    const zoneNumber = crsMode === "UTM35N" ? 35 : 36;
    const a = 6378137.0;
    const e = 0.0818191908426215;
    const e1sq = 0.00673949674227643;
    const k0 = 0.9996;

    const x = easting - 500000.0;
    const y = northing;
    const lonOrigin = (zoneNumber - 1) * 6 - 180 + 3;

    const M = y / k0;
    const mu = M / (a * (1 - Math.pow(e, 2) / 4 - 3 * Math.pow(e, 4) / 64 - 5 * Math.pow(e, 6) / 256));
    const e1 = (1 - Math.sqrt(1 - e * e)) / (1 + Math.sqrt(1 - e * e));
    const J1 = (3 * e1) / 2 - (27 * Math.pow(e1, 3)) / 32;
    const J2 = (21 * Math.pow(e1, 2)) / 16 - (55 * Math.pow(e1, 4)) / 32;
    const J3 = (151 * Math.pow(e1, 3)) / 96;
    const J4 = (1097 * Math.pow(e1, 4)) / 512;

    const fp = mu + J1 * Math.sin(2 * mu) + J2 * Math.sin(4 * mu) + J3 * Math.sin(6 * mu) + J4 * Math.sin(8 * mu);
    const C1 = e1sq * Math.pow(Math.cos(fp), 2);
    const T1 = Math.pow(Math.tan(fp), 2);
    const R1 = (a * (1 - e * e)) / Math.pow(1 - Math.pow(e * Math.sin(fp), 2), 1.5);
    const N1 = a / Math.sqrt(1 - Math.pow(e * Math.sin(fp), 2));
    const D = x / (N1 * k0);

    const Q1 = (N1 * Math.tan(fp)) / R1;
    const Q2 = (D * D) / 2;
    const Q3 = ((5 + 3 * T1 + 10 * C1 - 4 * C1 * C1 - 9 * e1sq) * Math.pow(D, 4)) / 24;
    const Q4 = ((61 + 90 * T1 + 298 * C1 + 45 * T1 * T1 - 252 * e1sq - 3 * C1 * C1) * Math.pow(D, 6)) / 720;
    const lat = fp - Q1 * (Q2 - Q3 + Q4);

    const Q5 = D;
    const Q6 = ((1 + 2 * T1 + C1) * Math.pow(D, 3)) / 6;
    const Q7 = ((5 - 2 * C1 + 28 * T1 - 3 * C1 * C1 + 8 * e1sq + 24 * T1 * T1) * Math.pow(D, 5)) / 120;
    const lon = (lonOrigin * Math.PI) / 180 + (Q5 - Q6 + Q7) / Math.cos(fp);

    return { lat: (lat * 180) / Math.PI, lon: (lon * 180) / Math.PI };
  }

  /* ============================================================================
     استيراد الملفات (DXF, XLSX, SDR/CDR, GSI, CSV, TXT, AutoCAD LIST)
     ============================================================================ */
  function parseDxfContent(dxfText) {
    const lines = dxfText.split(/\r?\n/).map(l => l.trim());
    const extracted = [];
    let ptCounter = 1;

    for (let i = 0; i < lines.length - 1; i++) {
      if (lines[i] === "0" && (lines[i + 1] === "POINT" || lines[i + 1] === "VERTEX")) {
        let x = null, y = null, z = 0, layer = "CAD";
        for (let j = i + 2; j < Math.min(i + 40, lines.length - 1); j += 2) {
          const code = lines[j];
          const val = lines[j + 1];
          if (code === "0") break;
          if (code === "8") layer = val;
          if (code === "10") x = parseFloat(val);
          if (code === "20") y = parseFloat(val);
          if (code === "30") z = parseFloat(val) || 0;
        }
        if (x !== null && y !== null && !isNaN(x) && !isNaN(y)) {
          extracted.push({ P: String(ptCounter++), E: x, N: y, Z: z, Code: layer });
        }
      } else if (lines[i] === "0" && lines[i + 1] === "LWPOLYLINE") {
        let layer = "POLY", curX = null;
        for (let j = i + 2; j < Math.min(i + 400, lines.length - 1); j += 2) {
          const code = lines[j];
          const val = lines[j + 1];
          if (code === "0") break;
          if (code === "8") layer = val;
          if (code === "10") curX = parseFloat(val);
          if (code === "20" && curX !== null) {
            const curY = parseFloat(val);
            if (!isNaN(curX) && !isNaN(curY)) {
              extracted.push({ P: String(ptCounter++), E: curX, N: curY, Z: 0, Code: layer });
            }
            curX = null;
          }
        }
      }
    }
    return extracted;
  }

  function parseAutoCADListText(rawText) {
    const regex = /X\s*=\s*([+-]?\d+(?:\.\d+)?)\s*Y\s*=\s*([+-]?\d+(?:\.\d+)?)(?:\s*Z\s*=\s*([+-]?\d+(?:\.\d+)?))?/gi;
    const found = [];
    let match;
    let idx = rows.length + 1;
    while ((match = regex.exec(rawText)) !== null) {
      found.push({
        P: String(idx++),
        E: parseFloat(match[1]),
        N: parseFloat(match[2]),
        Z: match[3] !== undefined ? parseFloat(match[3]) : 0,
        Code: "ACAD_LIST"
      });
    }
    return found;
  }

  function handleFileUpload(file) {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    saveSnapshot();

    if (ext === "xlsx" || ext === "xls") {
      ensureSheetJS(() => {
        const reader = new FileReader();
        reader.onload = e => {
          try {
            const wb = window.XLSX.read(e.target.result, { type: "array" });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const jsonRows = window.XLSX.utils.sheet_to_json(sheet, { defval: "" });
            if (!jsonRows.length) return showExcelLog("ملف الإكسيل فارغ.", "error");
            columns = Object.keys(jsonRows[0]);
            rows = jsonRows;
            renderExcelTable();
            drawCadPreview();
            showExcelLog(`📗 تم استيراد ${rows.length} نقطة من ملف Excel (${file.name}).`, "success");
          } catch (err) {
            showExcelLog("خطأ في قراءة Excel: " + err.message, "error");
          }
        };
        reader.readAsArrayBuffer(file);
      });
      return;
    }

    const reader = new FileReader();
    reader.onload = e => {
      const text = String(e.target.result || "");

      if (ext === "dxf") {
        const pts = parseDxfContent(text);
        if (!pts.length) return showExcelLog("لم يتم العثور على عناصر POINT أو POLYLINE داخل ملف DXF.", "error");
        columns = ["P", "E", "N", "Z", "Code"];
        rows = pts;
        renderExcelTable();
        drawCadPreview();
        showExcelLog(`📐 تم استيراد ${rows.length} نقطة مباشرة من لوحة AutoCAD DXF (${file.name})!`, "success");
        return;
      }

      const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
      if (!lines.length) return;

      if (ext === "sdr" || ext === "cdr") {
        const parsed = [];
        lines.forEach(line => {
          if (line.startsWith("08KI") || line.startsWith("08TP")) {
            const ptId = line.slice(4, 20).trim();
            const nVal = parseFloat(line.slice(20, 36).trim());
            const eVal = parseFloat(line.slice(36, 52).trim());
            const zVal = parseFloat(line.slice(52, 68).trim());
            const code = line.slice(68).trim();
            if (!isNaN(eVal) && !isNaN(nVal)) {
              parsed.push({ P: ptId || String(parsed.length + 1), E: eVal, N: nVal, Z: isNaN(zVal) ? 0 : zVal, Code: code });
            }
          }
        });
        if (parsed.length) {
          columns = ["P", "E", "N", "Z", "Code"];
          rows = parsed;
          renderExcelTable();
          drawCadPreview();
          showExcelLog(`📡 تم استيراد وتحويل ${rows.length} نقطة من ملف التوتال ستيشن (${file.name})!`, "success");
          return;
        }
      }

      if (ext === "gsi") {
        const parsed = [];
        lines.forEach(line => {
          const mP = line.match(/11\d{4}\+([^\s]+)/);
          const mE = line.match(/81\d{4}\+([+-]?\d+)/);
          const mN = line.match(/82\d{4}\+([+-]?\d+)/);
          const mZ = line.match(/83\d{4}\+([+-]?\d+)/);
          if (mE && mN) {
            parsed.push({
              P: mP ? mP[1].replace(/^0+/, "") || "1" : String(parsed.length + 1),
              E: parseFloat(mE[1]) / 1000,
              N: parseFloat(mN[1]) / 1000,
              Z: mZ ? parseFloat(mZ[1]) / 1000 : 0,
              Code: "LEICA"
            });
          }
        });
        if (parsed.length) {
          columns = ["P", "E", "N", "Z", "Code"];
          rows = parsed;
          renderExcelTable();
          drawCadPreview();
          showExcelLog(`📡 تم استيراد ${rows.length} نقطة من ملف Leica GSI!`, "success");
          return;
        }
      }

      // CSV / TXT عام
      const sep = lines[0].includes("\t") ? "\t" : lines[0].includes(";") ? ";" : ",";
      const firstCells = lines[0].split(sep).map(s => s.trim());
      const hasHeader = isNaN(parseFloat(firstCells[1]));
      if (hasHeader) {
        columns = firstCells;
        rows = lines.slice(1).map(ln => {
          const vals = ln.split(sep);
          const obj = {};
          columns.forEach((c, idx) => (obj[c] = (vals[idx] || "").trim()));
          return obj;
        });
      } else {
        columns = ["P", "E", "N", "Z", "Code"].slice(0, Math.max(3, firstCells.length));
        rows = lines.map((ln, rIdx) => {
          const vals = ln.split(sep);
          const obj = {};
          columns.forEach((c, idx) => (obj[c] = (vals[idx] || (idx === 0 ? String(rIdx + 1) : "")).trim()));
          return obj;
        });
      }
      renderExcelTable();
      drawCadPreview();
      showExcelLog(`✅ تم استيراد ${rows.length} نقطة من ${file.name}.`, "success");
    };
    reader.readAsText(file);
  }

  /* ============================================================================
     التصدير المباشر لجميع برامج الهندسة و GIS وأجهزة التوتال ستيشن
     ============================================================================ */
  function downloadBlob(content, filename, mime) {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function exportData(format) {
    if (!rows.length) return showExcelLog("الجدول فارغ، أضف نقاطاً أولاً.", "error");
    const { pKey, eKey, nKey, zKey, cKey } = getCoordKeys();

    // 1. تصدير ملف AutoCAD DXF رسمي حقيقي
    if (format === "dxf") {
      const dxf = ["0", "SECTION", "2", "ENTITIES"];
      const polyPts = [];

      rows.forEach(r => {
        const e = parseFloat(r[eKey]);
        const n = parseFloat(r[nKey]);
        const z = parseFloat(r[zKey]) || 0;
        const p = String(r[pKey] || "");
        const c = String(r[cKey] || "");
        if (isNaN(e) || isNaN(n)) return;

        polyPts.push({ e, n, z });

        // كيان النقطة POINT
        dxf.push("0", "POINT", "8", "GD_POINTS", "10", e.toFixed(4), "20", n.toFixed(4), "30", z.toFixed(4));
        // نص رقم النقطة والكود
        dxf.push("0", "TEXT", "8", "GD_TEXT_ID", "10", (e + 0.35).toFixed(4), "20", (n + 0.35).toFixed(4), "30", z.toFixed(4), "40", "0.55", "1", `${p} ${c}`.trim());
        // نص المنسوب Z
        dxf.push("0", "TEXT", "8", "GD_TEXT_ELEV", "10", (e + 0.35).toFixed(4), "20", (n - 0.35).toFixed(4), "30", z.toFixed(4), "40", "0.45", "1", z.toFixed(3));
      });

      // رسم خط متصل POLYLINE يربط النقاط
      if (polyPts.length >= 2) {
        dxf.push("0", "POLYLINE", "8", "GD_POLYLINE", "66", "1", "70", "8");
        polyPts.forEach(pt => {
          dxf.push("0", "VERTEX", "8", "GD_POLYLINE", "10", pt.e.toFixed(4), "20", pt.n.toFixed(4), "30", pt.z.toFixed(4));
        });
        dxf.push("0", "SEQEND");
      }

      dxf.push("0", "ENDSEC", "0", "EOF");
      downloadBlob(dxf.join("\r\n") + "\r\n", `GeoDev_AutoCAD_${Date.now()}.dxf`, "application/dxf");
      showExcelLog("📕 تم تصدير لوحة AutoCAD DXF (.dxf) بطبقات النقاط والأرقام والمناسيب والخطوط!", "success");
      return;
    }

    // 2. تصدير برنامج Golden Software Surfer (.dat)
    if (format === "surfer") {
      const datLines = ["Easting,Northing,Elevation,PointID"];
      rows.forEach(r => {
        const e = parseFloat(r[eKey]);
        const n = parseFloat(r[nKey]);
        const z = parseFloat(r[zKey]) || 0;
        const p = String(r[pKey] || "");
        if (!isNaN(e) && !isNaN(n)) {
          datLines.push(`${e.toFixed(4)},${n.toFixed(4)},${z.toFixed(4)},${p}`);
        }
      });
      downloadBlob(datLines.join("\r\n") + "\r\n", `GeoDev_Surfer_${Date.now()}.dat`, "text/plain;charset=utf-8;");
      showExcelLog("🏔️ تم تصدير ملف Golden Software Surfer (.dat) جاهز لإنشاء الخرائط الكنتورية وحساب الكميات!", "success");
      return;
    }

    // 3. تصدير ArcGIS Pro WKT CSV (POINT Z)
    if (format === "wkt") {
      const wktHeader = [...columns, "GEOMETRY_WKT"].join(",");
      const wktRows = rows.map(r => {
        const e = parseFloat(r[eKey]);
        const n = parseFloat(r[nKey]);
        const z = parseFloat(r[zKey]) || 0;
        const geom = !isNaN(e) && !isNaN(n) ? `"POINT Z (${e.toFixed(4)} ${n.toFixed(4)} ${z.toFixed(4)})"` : '""';
        const vals = columns.map(col => `"${String(r[col] ?? "").replace(/"/g, '""')}"`);
        return [...vals, geom].join(",");
      });
      downloadBlob("\uFEFF" + [wktHeader, ...wktRows].join("\n"), `GeoDev_ArcGISPro_WKT_${Date.now()}.csv`, "text/csv;charset=utf-8;");
      showExcelLog("📊 تم تصدير جدول ArcGIS Pro بحقل GEOMETRY_WKT ثلاثي الأبعاد!", "success");
      return;
    }

    // 4. تصدير ArcGIS Pro / QGIS GeoJSON
    if (format === "geojson") {
      const features = rows
        .map(r => {
          const e = parseFloat(r[eKey]);
          const n = parseFloat(r[nKey]);
          if (isNaN(e) || isNaN(n)) return null;
          const ll = projectToWGS84(e, n);
          return {
            type: "Feature",
            properties: { ...r, CRS_Source: crsMode },
            geometry: {
              type: "Point",
              coordinates: [Number(ll.lon.toFixed(7)), Number(ll.lat.toFixed(7)), parseFloat(r[zKey]) || 0]
            }
          };
        })
        .filter(Boolean);

      downloadBlob(JSON.stringify({ type: "FeatureCollection", features }, null, 2), `GeoDev_ArcGIS_QGIS_${Date.now()}.geojson`, "application/geo+json;charset=utf-8;");
      showExcelLog("🗺️ تم تصدير طبقة GeoJSON جاهزة للفتح مباشرة في ArcGIS Pro و QGIS!", "success");
      return;
    }

    // 5. تصدير Google Earth KML
    if (format === "kml") {
      const marks = rows
        .map(r => {
          const e = parseFloat(r[eKey]);
          const n = parseFloat(r[nKey]);
          const z = parseFloat(r[zKey]) || 0;
          if (isNaN(e) || isNaN(n)) return "";
          const ll = projectToWGS84(e, n);
          return `<Placemark><name>${String(r[pKey] || "")}</name><description>E=${e} | N=${n} | Z=${z} | ${String(r[cKey] || "")}</description><Point><coordinates>${ll.lon.toFixed(7)},${ll.lat.toFixed(7)},${z}</coordinates></Point></Placemark>`;
        })
        .join("\n");
      const kml = `<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>GeoDev Survey</name>\n${marks}\n</Document></kml>`;
      downloadBlob(kml, `GeoDev_GoogleEarth_${Date.now()}.kml`, "application/vnd.google-earth.kml+xml;charset=utf-8;");
      showExcelLog("🛰️ تم تصدير ملف Google Earth (.kml) بنجاح!", "success");
      return;
    }

    // 6. تصدير Sokkia SDR33 / CDR
    if (format === "sdr") {
      const sdrLines = ["00NMSDR33 V04-04.02  GEODEV STAKEOUT"];
      rows.forEach(r => {
        const p = String(r[pKey] || "").padEnd(16, " ").slice(0, 16);
        const n = (parseFloat(r[nKey]) || 0).toFixed(3).padStart(16, " ");
        const e = (parseFloat(r[eKey]) || 0).toFixed(3).padStart(16, " ");
        const z = (parseFloat(r[zKey]) || 0).toFixed(3).padStart(16, " ");
        const c = String(r[cKey] || "").padEnd(16, " ").slice(0, 16);
        sdrLines.push(`08KI${p}${n}${e}${z}${c}`);
      });
      downloadBlob(sdrLines.join("\r\n") + "\r\n", `GeoDev_Sokkia_${Date.now()}.sdr`, "text/plain;charset=utf-8;");
      showExcelLog("📡 تم تصدير ملف Sokkia SDR (.sdr) جاهز للرفع على جهاز التوتال ستيشن!", "success");
      return;
    }

    // 7. تصدير Leica GSI
    if (format === "gsi") {
      const gsiLines = rows.map((r, idx) => {
        const p = String(r[pKey] || idx + 1).padStart(16, "0");
        const eMm = Math.round((parseFloat(r[eKey]) || 0) * 1000);
        const nMm = Math.round((parseFloat(r[nKey]) || 0) * 1000);
        const zMm = Math.round((parseFloat(r[zKey]) || 0) * 1000);
        const fmt = val => (val >= 0 ? "+" : "-") + String(Math.abs(val)).padStart(16, "0");
        return `11${String(idx + 1).padStart(4, "0")}+${p} 81..10${fmt(eMm)} 82..10${fmt(nMm)} 83..10${fmt(zMm)}`;
      });
      downloadBlob(gsiLines.join("\r\n") + "\r\n", `GeoDev_Leica_${Date.now()}.gsi`, "text/plain;charset=utf-8;");
      showExcelLog("📡 تم تصدير ملف Leica GSI (.gsi) جاهز لأجهزة لايكا!", "success");
      return;
    }

    // 8. تصدير Civil 3D CSV
    if (format === "csv") {
      const csvLines = rows.map(r => [r[pKey] || "", r[eKey] || "", r[nKey] || "", r[zKey] || "0", r[cKey] || ""].join(","));
      downloadBlob("\uFEFF" + csvLines.join("\n"), `GeoDev_Civil3D_PENZD_${Date.now()}.csv`, "text/csv;charset=utf-8;");
      showExcelLog("📘 تم تصدير ملف Civil 3D (PENZD .csv) بنجاح!", "success");
      return;
    }

    // 9. تصدير AutoLISP
    if (format === "lsp") {
      const ptsLisp = rows
        .map(r => {
          const e = parseFloat(r[eKey]);
          const n = parseFloat(r[nKey]);
          const z = parseFloat(r[zKey]) || 0;
          const p = String(r[pKey] || "").replace(/"/g, "");
          const c = String(r[cKey] || "").replace(/"/g, "");
          if (isNaN(e) || isNaN(n)) return "";
          return `  (gd-add-pt "${p}" ${e} ${n} ${z} "${c}")`;
        })
        .filter(Boolean)
        .join("\n");

      const lispCode = `;;; GeoDev AutoLISP Point Plotter - Command: GD_PLOT_PTS
(defun c:GD_PLOT_PTS (/ oldosm)
  (setq oldosm (getvar "OSMODE"))
  (setvar "OSMODE" 0)
  (setvar "PDMODE" 34)
  (setvar "PDSIZE" 0.5)
  (command "_.-LAYER" "_M" "GD_POINTS" "_C" "3" "" "")
  (command "_.-LAYER" "_M" "GD_LABELS" "_C" "2" "" "")
  (defun gd-add-pt (pid x y z code)
    (entmake (list '(0 . "POINT") '(8 . "GD_POINTS") (list 10 x y z)))
    (entmake (list '(0 . "TEXT") '(8 . "GD_LABELS") (list 10 (+ x 0.4) (+ y 0.4) z) '(40 . 0.55) (cons 1 (strcat pid " | Z=" (rtos z 2 3) " " code))))
  )
${ptsLisp}
  (setvar "OSMODE" oldosm)
  (command "_.ZOOM" "_E")
  (princ "\\n[GeoDev] Plotted all points successfully!")
  (princ)
)
`;
      downloadBlob(lispCode, `GeoDev_PlotPoints_${Date.now()}.lsp`, "text/plain;charset=utf-8;");
      showExcelLog("⚡ تم تصدير ملف AutoLISP (.lsp)! اكتب الأمر GD_PLOT_PTS داخل الأوتوكاد.", "success");
      return;
    }

    // 10. تصدير Excel XLSX
    if (format === "xlsx") {
      ensureSheetJS(() => {
        const ws = window.XLSX.utils.json_to_sheet(rows, { header: columns });
        const wb = window.XLSX.utils.book_new();
        window.XLSX.utils.book_append_sheet(wb, ws, "GeoDev_Survey");
        window.XLSX.writeFile(wb, `GeoDev_Survey_${Date.now()}.xlsx`);
        showExcelLog("📗 تم تصدير ملف Excel (.xlsx) بنجاح!", "success");
      });
    }
  }

  /* ============================================================================
     محرك الأوامر الذكية بالعربي والإنجليزي
     ============================================================================ */
  function executeAICommand(rawPrompt) {
    const cmd = String(rawPrompt || "").trim();
    if (!cmd) return;
    saveSnapshot();
    const { pKey, eKey, nKey, zKey, cKey } = getCoordKeys();

    if (/تبديل|بدل\s*الشرق/i.test(cmd)) {
      rows.forEach(r => {
        const tmp = r[eKey];
        r[eKey] = r[nKey];
        r[nKey] = tmp;
      });
      renderExcelTable();
      drawCadPreview();
      showExcelLog("🔄 تم تبديل إحداثيات الشرق (E) والشمال (N) لجميع النقاط.", "success");
      return;
    }

    const writeMatch = cmd.match(/(?:اكتب|حط|سجل|أضف|اضف|ضع)\s*(?:لي\s*)?(?:في\s*(?:ملف\s*)?(?:الإكسيل|الاكسيل|الجدول)\s*)?(.+)/i);
    if (writeMatch && !/عمود|نقطة|نقطه|صف|مكرر|مسافة|مساحه|مساحة/.test(writeMatch[1])) {
      const payload = writeMatch[1].trim();
      const tokens = payload.split(/[\s,،]+/).filter(Boolean);
      const colName = "AI_Note";
      if (!columns.includes(colName)) columns.push(colName);
      if (tokens.length > 1) {
        while (rows.length < tokens.length) {
          const obj = {};
          columns.forEach(c => (obj[c] = ""));
          obj[pKey] = String(rows.length + 1);
          rows.push(obj);
        }
        tokens.forEach((tok, i) => {
          rows[i][colName] = tok;
          if (cKey && !rows[i][cKey]) rows[i][cKey] = tok;
        });
      } else {
        rows.forEach(r => (r[colName] = payload));
      }
      renderExcelTable();
      drawCadPreview();
      showExcelLog(`✅ تم تنفيذ أمرك وكتابة «${payload}» داخل جدول الإكسيل بنجاح.`, "success");
      return;
    }

    const mathMatch = cmd.match(/(اجمع|زود|اطرح|نقص|اضرب)\s*([+-]?\d+(?:\.\d+)?)\s*(?:على|من|في)\s*(?:كل\s*)?(?:عمود\s*|المناسيب\s*)?([a-zA-Z0-9_\u0600-\u06FF]+)/i);
    if (mathMatch) {
      const op = mathMatch[1];
      const num = parseFloat(mathMatch[2]);
      const targetRaw = mathMatch[3].trim();
      let targetCol = columns.find(c => c.toLowerCase() === targetRaw.toLowerCase());
      if (!targetCol) {
        if (/منسوب|مناسيب|z/i.test(targetRaw)) targetCol = zKey;
        else if (/شرق|e/i.test(targetRaw)) targetCol = eKey;
        else if (/شمال|n/i.test(targetRaw)) targetCol = nKey;
      }
      if (!targetCol) return showExcelLog(`لم أجد العمود «${targetRaw}».`, "error");
      rows.forEach(r => {
        const cur = parseFloat(r[targetCol]);
        if (!isNaN(cur)) {
          let res = cur;
          if (/اجمع|زود/.test(op)) res = cur + num;
          else if (/اطرح|نقص/.test(op)) res = cur - num;
          else if (/اضرب/.test(op)) res = cur * num;
          r[targetCol] = Number(res.toFixed(3));
        }
      });
      renderExcelTable();
      drawCadPreview();
      showExcelLog(`✅ تم تعديل عمود «${targetCol}» (${op} ${num}) لجميع النقاط.`, "success");
      return;
    }

    if (/مساف|انحراف|azimuth|dist/i.test(cmd)) {
      if (!columns.includes("Dist_m")) columns.push("Dist_m");
      if (!columns.includes("Azimuth_deg")) columns.push("Azimuth_deg");
      for (let i = 0; i < rows.length; i++) {
        if (i === 0) {
          rows[i].Dist_m = "0.000";
          rows[i].Azimuth_deg = "-";
          continue;
        }
        const dE = (parseFloat(rows[i][eKey]) || 0) - (parseFloat(rows[i - 1][eKey]) || 0);
        const dN = (parseFloat(rows[i][nKey]) || 0) - (parseFloat(rows[i - 1][nKey]) || 0);
        let az = (Math.atan2(dE, dN) * 180) / Math.PI;
        if (az < 0) az += 360;
        rows[i].Dist_m = Math.hypot(dE, dN).toFixed(3);
        rows[i].Azimuth_deg = az.toFixed(4) + "°";
      }
      renderExcelTable();
      showExcelLog("✅ تم حساب المسافة والانحراف وإضافتهما كأعمدة جديدة في الجدول.", "success");
      return;
    }

    if (/مساح|محيط|area|perim/i.test(cmd)) {
      const st = computeSurveyStats();
      const feddan = st.areaM2 / 4200.83;
      const qirat = (feddan % 1) * 24;
      const sahm = (qirat % 1) * 24;
      showExcelLog(`📐 المساحة = ${st.areaM2.toFixed(2)} م² (${Math.floor(feddan)} فدان، ${Math.floor(qirat)} قيراط، ${sahm.toFixed(1)} سهم) | المحيط = ${st.perimeter.toFixed(2)} م`, "success");
      return;
    }

    if (/مكرر|تكرار|duplicate/i.test(cmd)) {
      const before = rows.length;
      const seen = new Set();
      rows = rows.filter(r => {
        const key = `${r[eKey]}_${r[nKey]}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
      renderExcelTable();
      drawCadPreview();
      showExcelLog(`🧹 تم حذف ${before - rows.length} نقطة مكررة.`, "success");
      return;
    }

    if (!columns.includes("AI_Note")) columns.push("AI_Note");
    rows.forEach(r => (r.AI_Note = cmd));
    renderExcelTable();
    showExcelLog(`✅ تم تسجيل «${cmd}» في الجدول.`, "info");
  }

  /* ============================================================================
     رسم الجدول وشاشة CAD Canvas
     ============================================================================ */
  function renderExcelTable() {
    const table = document.getElementById("gdExcelDataTable");
    const statsBar = document.getElementById("gdExStatsBar");
    if (!table) return;

    const st = computeSurveyStats();
    if (statsBar) {
      statsBar.innerHTML = `
        <div class="gd-ex-stat"><span>عدد النقاط</span><strong>${st.count}</strong></div>
        <div class="gd-ex-stat"><span>أقل منسوب (Min Z)</span><strong>${st.minZ}</strong></div>
        <div class="gd-ex-stat"><span>أعلى منسوب (Max Z)</span><strong>${st.maxZ}</strong></div>
        <div class="gd-ex-stat"><span>المحيط / الطول</span><strong>${st.perimeter.toFixed(2)} م</strong></div>
        <div class="gd-ex-stat"><span>المساحة المغلقة</span><strong>${st.areaM2.toFixed(2)} م²</strong></div>
      `;
    }

    const q = searchFilter.trim().toLowerCase();
    const filtered = rows
      .map((r, originalIdx) => ({ r, originalIdx }))
      .filter(({ r }) => !q || columns.some(c => String(r[c] ?? "").toLowerCase().includes(q)));

    table.innerHTML = `
      <thead>
        <tr>
          <th>#</th>
          ${columns.map(col => `<th>${col}</th>`).join("")}
          <th>حذف</th>
        </tr>
      </thead>
      <tbody>
        ${filtered.map(({ r, originalIdx }) => `
          <tr>
            <td style="background:#211e1b;color:#a8a29e;font-weight:800;padding:6px 8px;">${originalIdx + 1}</td>
            ${columns.map(col => `
              <td>
                <input type="text" class="gd-ex-cell"
                  data-row="${originalIdx}" data-col="${col}"
                  value="${String(r[col] ?? "").replace(/"/g, "&quot;")}">
              </td>
            `).join("")}
            <td>
              <button type="button" style="background:transparent;border:0;color:#f87171;cursor:pointer;font-weight:900;padding:6px 9px;"
                data-del-row="${originalIdx}">✕</button>
            </td>
          </tr>
        `).join("")}
      </tbody>
    `;
  }

  function drawCadPreview() {
    const canvas = document.getElementById("gdExcelCadCanvas");
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    const rect = canvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(300, rect.width * dpr);
    canvas.height = 245 * dpr;
    ctx.scale(dpr, dpr);

    const W = rect.width || 320;
    const H = 245;

    ctx.fillStyle = "#04060a";
    ctx.fillRect(0, 0, W, H);
    ctx.strokeStyle = "rgba(56, 189, 248, 0.08)";
    ctx.lineWidth = 1;
    for (let x = 0; x < W; x += 28) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, H); ctx.stroke(); }
    for (let y = 0; y < H; y += 28) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(W, y); ctx.stroke(); }

    const { pKey, eKey, nKey, zKey } = getCoordKeys();
    const pts = rows
      .map(r => ({
        p: String(r[pKey] || ""),
        e: parseFloat(r[eKey]),
        n: parseFloat(r[nKey]),
        z: parseFloat(r[zKey])
      }))
      .filter(pt => !isNaN(pt.e) && !isNaN(pt.n));

    if (!pts.length) {
      ctx.fillStyle = "#64748b";
      ctx.font = "bold 13px sans-serif";
      ctx.textAlign = "center";
      ctx.fillText("لا توجد إحداثيات E / N صالحة للرسم", W / 2, H / 2);
      return;
    }

    const minE = Math.min(...pts.map(p => p.e));
    const maxE = Math.max(...pts.map(p => p.e));
    const minN = Math.min(...pts.map(p => p.n));
    const maxN = Math.max(...pts.map(p => p.n));
    const spanE = Math.max(maxE - minE, 1);
    const spanN = Math.max(maxN - minN, 1);
    const pad = 34;
    const scale = Math.min((W - pad * 2) / spanE, (H - pad * 2) / spanN);

    const toCanvas = (e, n) => ({
      x: pad + (e - minE) * scale + (W - pad * 2 - spanE * scale) / 2,
      y: H - (pad + (n - minN) * scale + (H - pad * 2 - spanN * scale) / 2)
    });

    if (connectLinesOnCanvas && pts.length >= 2) {
      ctx.beginPath();
      ctx.strokeStyle = "#38bdf8";
      ctx.lineWidth = 1.6;
      pts.forEach((pt, idx) => {
        const c = toCanvas(pt.e, pt.n);
        if (idx === 0) ctx.moveTo(c.x, c.y);
        else ctx.lineTo(c.x, c.y);
      });
      ctx.stroke();
    }

    pts.forEach(pt => {
      const c = toCanvas(pt.e, pt.n);
      ctx.beginPath();
      ctx.arc(c.x, c.y, 4, 0, Math.PI * 2);
      ctx.fillStyle = "#fbbf24";
      ctx.fill();
      ctx.strokeStyle = "#04060a";
      ctx.lineWidth = 1.2;
      ctx.stroke();

      if (showLabelsOnCanvas) {
        ctx.fillStyle = "#f8fafc";
        ctx.font = "bold 10.5px sans-serif";
        ctx.textAlign = "left";
        const zTxt = !isNaN(pt.z) ? ` (${pt.z.toFixed(2)})` : "";
        ctx.fillText(`${pt.p}${zTxt}`, c.x + 6, c.y - 4);
      }
    });
  }

  function findSurveyingSection() {
    const navButtons = Array.from(document.querySelectorAll("nav button, [data-t]"));
    const survBtn = navButtons.find(btn => (btn.textContent || "").trim() === "المساحة");
    if (survBtn && survBtn.dataset && survBtn.dataset.t) {
      const secByTab = document.getElementById("s-" + survBtn.dataset.t);
      if (secByTab) return secByTab;
    }
    return document.getElementById("s-surv") || document.getElementById("s-surveying") || null;
  }

  function mountAIExcelStudio() {
    const targetSec = findSurveyingSection();
    if (!targetSec) return;

    const existing = document.getElementById("gdAiExcelStudio");
    if (existing) {
      if (existing.parentElement !== targetSec) targetSec.prepend(existing);
      return;
    }

    const container = document.createElement("div");
    container.id = "gdAiExcelStudio";
    container.innerHTML = renderStudioHTML();
    targetSec.prepend(container);

    renderExcelTable();
    setTimeout(drawCadPreview, 80);
  }

  function bindEvents() {
    document.addEventListener("click", e => {
      if (e.target.closest("#gdExcelAiExecBtn")) {
        const inp = document.getElementById("gdExcelAiPrompt");
        if (inp) executeAICommand(inp.value);
        return;
      }

      const chip = e.target.closest("[data-ai-cmd]");
      if (chip) {
        const cmd = chip.dataset.aiCmd;
        const inp = document.getElementById("gdExcelAiPrompt");
        if (inp) inp.value = cmd;
        executeAICommand(cmd);
        return;
      }

      const subTab = e.target.closest("[data-ex-subtab]");
      if (subTab) {
        activeSubTab = subTab.dataset.exSubtab;
        document.querySelectorAll("[data-ex-subtab]").forEach(b => b.classList.toggle("active", b.dataset.exSubtab === activeSubTab));
        const pCad = document.getElementById("gdExSubpanelCadImport");
        const pCalc = document.getElementById("gdExSubpanelCalcTools");
        if (pCad) pCad.style.display = activeSubTab === "cad_import" ? "block" : "none";
        if (pCalc) pCalc.style.display = activeSubTab === "calc_tools" ? "block" : "none";
        return;
      }

      if (e.target.closest("#gdParseAcadListBtn")) {
        const txt = document.getElementById("gdAcadListPasteArea")?.value || "";
        const extracted = parseAutoCADListText(txt);
        if (!extracted.length) {
          showExcelLog("لم يتم العثور على إحداثيات X= ... Y= ... في النص الملصق.", "error");
          return;
        }
        saveSnapshot();
        rows.push(...extracted);
        renderExcelTable();
        drawCadPreview();
        showExcelLog(`📐 تم استخراج وإضافة ${extracted.length} نقطة من نص الأوتوكاد!`, "success");
        return;
      }

      if (e.target.closest("#gdApplyShiftBtn")) {
        const dE = parseFloat(document.getElementById("gdShiftE")?.value) || 0;
        const dN = parseFloat(document.getElementById("gdShiftN")?.value) || 0;
        const dZ = parseFloat(document.getElementById("gdShiftZ")?.value) || 0;
        if (!dE && !dN && !dZ) return;
        saveSnapshot();
        const { eKey, nKey, zKey } = getCoordKeys();
        rows.forEach(r => {
          if (!isNaN(parseFloat(r[eKey]))) r[eKey] = Number((parseFloat(r[eKey]) + dE).toFixed(3));
          if (!isNaN(parseFloat(r[nKey]))) r[nKey] = Number((parseFloat(r[nKey]) + dN).toFixed(3));
          if (!isNaN(parseFloat(r[zKey]))) r[zKey] = Number((parseFloat(r[zKey]) + dZ).toFixed(3));
        });
        renderExcelTable();
        drawCadPreview();
        showExcelLog(`✅ تم تطبيق إزاحة الإحداثيات (ΔE=${dE}, ΔN=${dN}, ΔZ=${dZ}) بنجاح.`, "success");
        return;
      }

      if (e.target.closest("#gdCalcCutFillBtn")) {
        const designLvl = parseFloat(document.getElementById("gdDesignLevelInput")?.value);
        if (isNaN(designLvl)) return showExcelLog("أدخل منسوب التصميم أولاً.", "error");
        saveSnapshot();
        const { zKey } = getCoordKeys();
        if (!columns.includes("Cut_Fill")) columns.push("Cut_Fill");
        let totalCut = 0, totalFill = 0;
        rows.forEach(r => {
          const z = parseFloat(r[zKey]);
          if (!isNaN(z)) {
            const diff = z - designLvl;
            if (diff > 0) {
              r.Cut_Fill = `Cut +${diff.toFixed(3)}`;
              totalCut += diff;
            } else if (diff < 0) {
              r.Cut_Fill = `Fill ${diff.toFixed(3)}`;
              totalFill += Math.abs(diff);
            } else {
              r.Cut_Fill = "0.000";
            }
          }
        });
        renderExcelTable();
        showExcelLog(`🧮 تم حساب الحفر والردم مقارنة بمنسوب ${designLvl} م (متوسط فروق الحفر: ${totalCut.toFixed(2)} م | الردم: ${totalFill.toFixed(2)} م).`, "success");
        return;
      }

      if (e.target.closest("#gdExSwapENBtn")) {
        executeAICommand("تبديل الشرق والشمال");
        return;
      }

      if (e.target.closest("#gdExAddRowBtn")) {
        saveSnapshot();
        const { pKey } = getCoordKeys();
        const obj = {};
        columns.forEach(c => (obj[c] = ""));
        obj[pKey] = String(rows.length + 1);
        rows.push(obj);
        renderExcelTable();
        drawCadPreview();
        return;
      }

      if (e.target.closest("#gdExAddColBtn")) {
        const colName = prompt("أدخل اسم العمود الجديد (مثال: Offset أو Note):");
        if (!colName || !colName.trim()) return;
        const clean = colName.trim();
        if (!columns.includes(clean)) {
          saveSnapshot();
          columns.push(clean);
          rows.forEach(r => (r[clean] = ""));
          renderExcelTable();
        }
        return;
      }

      if (e.target.closest("#gdExUndoBtn")) { undoLastAction(); return; }

      if (e.target.closest("#gdExClearBtn")) {
        if (!confirm("هل تريد مسح جميع النقاط من الجدول؟")) return;
        saveSnapshot();
        rows = [];
        renderExcelTable();
        drawCadPreview();
        return;
      }

      const delRowBtn = e.target.closest("[data-del-row]");
      if (delRowBtn) {
        saveSnapshot();
        rows.splice(parseInt(delRowBtn.dataset.delRow, 10), 1);
        renderExcelTable();
        drawCadPreview();
        return;
      }

      if (e.target.closest("#gdToggleLabelsBtn")) { showLabelsOnCanvas = !showLabelsOnCanvas; drawCadPreview(); return; }
      if (e.target.closest("#gdToggleLinesBtn")) { connectLinesOnCanvas = !connectLinesOnCanvas; drawCadPreview(); return; }

      const expBtn = e.target.closest("[data-export-fmt]");
      if (expBtn) exportData(expBtn.dataset.exportFmt);
    });

    document.addEventListener("input", e => {
      if (e.target.classList.contains("gd-ex-cell")) {
        const rIdx = parseInt(e.target.dataset.row, 10);
        const col = e.target.dataset.col;
        if (rows[rIdx]) {
          rows[rIdx][col] = e.target.value;
          drawCadPreview();
        }
        return;
      }
      if (e.target.id === "gdExSearchInput") {
        searchFilter = e.target.value;
        renderExcelTable();
      }
    });

    document.addEventListener("change", e => {
      if (e.target.id === "gdExcelFileInput" && e.target.files?.[0]) {
        handleFileUpload(e.target.files[0]);
        e.target.value = "";
        return;
      }
      if (e.target.id === "gdCrsSelect") {
        crsMode = e.target.value;
        showExcelLog(`تم ضبط نظام الإحداثيات الجغرافي إلى: ${crsMode}`, "info");
      }
    });

    document.addEventListener("keydown", e => {
      if (e.target.id === "gdExcelAiPrompt" && e.key === "Enter") {
        e.preventDefault();
        executeAICommand(e.target.value);
      }
    });
  }

  function init() {
    ensureSheetJS();
    injectStyles();
    bindEvents();
    mountAIExcelStudio();

    const obs = new MutationObserver(() => {
      mountAIExcelStudio();
    });
    if (document.body) {
      obs.observe(document.body, { childList: true, subtree: true });
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
