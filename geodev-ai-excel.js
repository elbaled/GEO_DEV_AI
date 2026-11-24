/* ============================================================================
   GeoDev AI Survey & GIS Excel Workstation (V8.0 - As-Built & Official Report)
   - استيراد وتنظيف وتصنيف ملفات المساحة (Excel / CSV / DXF / SDR / STR / GSI)
   - معاينة CAD حية بالرموز + Zoom + 3 سمات لونية + رسم أسهم ترحيل As-Built
   - أداة مقارنة نقاط التصميم (Design) بالرفع الجديد (As-Built) وحساب التفاوتات
   - زر تصدير التقرير الهندسي بنافذة إدخال بيانات المشروع والمهندس والسماحية (PDF & Excel)
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevAiExcelV8Init) return;
  window.GeoDevAiExcelV8Init = true;

  const THEME_STORAGE_KEY = "gd_excel_station_theme_v1";
  const THEMES = ["blue", "gold", "light"];
  const THEME_LABELS = {
    blue: "السمة: الكاد الأزرق",
    gold: "السمة: الذهبي الداكن",
    light: "السمة: النهاري الفاتح"
  };

  let points = [
    { p: "1", e: 312450.12, n: 3110240.55, z: 24.50, code: "EG", cat: "EG" },
    { p: "2", e: 312475.84, n: 3110225.18, z: 24.68, code: "MNH", cat: "MNH" },
    { p: "3", e: 312510.30, n: 3110205.90, z: 24.82, code: "MNH", cat: "MNH" },
    { p: "4", e: 312495.60, n: 3110175.40, z: 24.39, code: "LP", cat: "LP" },
    { p: "5", e: 312480.15, n: 3110190.25, z: 24.42, code: "TREE", cat: "TREE" },
    { p: "6", e: 312462.40, n: 3110182.10, z: 24.45, code: "EG", cat: "EG" }
  ];

  let designPoints = []; // نقاط التصميم / التوقيع الأصلي لمقارنة As-Built
  let asBuiltComparisons = []; // نتائج المقارنة التفصيلية
  let undoStack = [];
  let activeFilter = "ALL";
  let showLabels = true;
  let connectByCode = false;
  let cadZoom = 1.0;
  let cadPanX = 0;
  let cadPanY = 0;
  let isDraggingCad = false;
  let dragStartX = 0;
  let dragStartY = 0;
  let currentTheme = localStorage.getItem(THEME_STORAGE_KEY) || "blue";
  if (!THEMES.includes(currentTheme)) currentTheme = "blue";

  /* ============================================================================
     تصنيف الأكواد الهندسية
     ============================================================================ */
  function classifyCode(rawCode) {
    const c = String(rawCode || "").trim().toUpperCase();
    if (!c || c === "EG" || c === "NGL" || c === "GND" || c.includes("ارض") || c.includes("أرض")) {
      return { cat: "EG", label: "أرض طبيعية (EG)", color: "#94a3b8" };
    }
    if (c.startsWith("MNH") || c.startsWith("MH") || c.includes("SEW") || c.includes("MANHOLE") || c.includes("مطبق")) {
      return { cat: "MNH", label: "مطبق / صرف صحي (MNH)", color: "#f97316" };
    }
    if (c.startsWith("LP") || c.startsWith("LIGHT") || c.startsWith("POLE") || c.includes("عمود") || c.includes("انارة") || c.includes("إنارة")) {
      return { cat: "LP", label: "عمود إنارة (LP)", color: "#facc15" };
    }
    if (c.startsWith("TR") || c.includes("TREE") || c.includes("PALM") || c.includes("شجر") || c.includes("نخل")) {
      return { cat: "TREE", label: "أشجار / تخضير (TREE)", color: "#22c55e" };
    }
    if (c.startsWith("COL") || c.startsWith("AX") || c.startsWith("C1") || c.startsWith("C2") || c.includes("عمود خرسان") || c.includes("قاعد")) {
      return { cat: "COL", label: "أعمدة / قواعد خرسانية (COL)", color: "#38bdf8" };
    }
    if (c.startsWith("RD") || c.startsWith("CRB") || c.startsWith("SW") || c.includes("طريق") || c.includes("بردور")) {
      return { cat: "ROAD", label: "طرق وأرصفة (ROAD)", color: "#a855f7" };
    }
    return { cat: "OTHER", label: `كود مخصص (${c})`, color: "#e879f9" };
  }

  function pushUndo() {
    undoStack.push(JSON.stringify(points));
    if (undoStack.length > 15) undoStack.shift();
  }

  /* ============================================================================
     استيراد الملفات الذكي (تجاهل العناوين والبدء من النقطة 1)
     ============================================================================ */
  function parseCoordinateText(text) {
    const lines = String(text || "").split(/\r?\n/);
    const parsed = [];
    let autoIdx = 1;

    for (let raw of lines) {
      const line = raw.trim();
      if (!line) continue;

      // دعم صيغة DXF البسيطة
      if (line === "EOF" || line === "SECTION" || line === "ENTITIES") continue;

      // دعم صيغة Leica GSI
      if (line.startsWith("*") || line.includes("11....+") || line.includes("81..")) {
        const pM = line.match(/11[^\s+]*\+([0-9A-Za-z]+)/);
        const eM = line.match(/81[^\s+]*\+([+-]?\d+)/);
        const nM = line.match(/82[^\s+]*\+([+-]?\d+)/);
        const zM = line.match(/83[^\s+]*\+([+-]?\d+)/);
        if (eM && nM) {
          const eVal = parseFloat(eM[1]) / 1000;
          const nVal = parseFloat(nM[1]) / 1000;
          const zVal = zM ? parseFloat(zM[1]) / 1000 : 0;
          const pId = pM ? pM[1].replace(/^0+/, "") || String(autoIdx) : String(autoIdx);
          const info = classifyCode("EG");
          parsed.push({ p: pId, e: eVal, n: nVal, z: zVal, code: "EG", cat: info.cat });
          autoIdx++;
        }
        continue;
      }

      // تقسيم الأعمدة (CSV / TXT / TSV / STR)
      const cols = line.split(/[,;\t|]+/).map(s => s.trim()).filter(Boolean);
      if (cols.length < 2) continue;

      // تخطي سطور العناوين غير الرقمية
      const numCols = cols.filter(c => !isNaN(parseFloat(c)) && isFinite(c));
      if (numCols.length < 2) continue;

      let p = String(autoIdx), e = 0, n = 0, z = 0, code = "EG";

      if (cols.length >= 4 && !isNaN(parseFloat(cols[1])) && !isNaN(parseFloat(cols[2]))) {
        p = cols[0];
        e = parseFloat(cols[1]);
        n = parseFloat(cols[2]);
        z = !isNaN(parseFloat(cols[3])) ? parseFloat(cols[3]) : 0;
        code = cols[4] || (isNaN(parseFloat(cols[3])) ? cols[3] : "EG");
      } else if (cols.length === 3) {
        if (!isNaN(parseFloat(cols[0])) && !isNaN(parseFloat(cols[1])) && !isNaN(parseFloat(cols[2]))) {
          p = String(autoIdx);
          e = parseFloat(cols[0]);
          n = parseFloat(cols[1]);
          z = parseFloat(cols[2]);
        } else {
          p = cols[0];
          e = parseFloat(cols[1]);
          n = parseFloat(cols[2]);
        }
      } else if (cols.length === 2) {
        e = parseFloat(cols[0]);
        n = parseFloat(cols[1]);
      }

      if (isNaN(e) || isNaN(n)) continue;
      const info = classifyCode(code);
      parsed.push({ p, e, n, z: isNaN(z) ? 0 : z, code: code || "EG", cat: info.cat });
      autoIdx++;
    }
    return parsed;
  }

  /* ============================================================================
     حساب المقارنة الهندسية بين نقاط التصميم (Design) والرفع الجديد (As-Built)
     ============================================================================ */
  function computeAsBuiltComparison(toleranceMm = 20) {
    if (!designPoints.length || !points.length) {
      asBuiltComparisons = [];
      return [];
    }

    const results = [];
    points.forEach((asPt, idx) => {
      // البحث أولاً بنفس رقم/اسم النقطة، وإلا بأقرب نقطة مكانياً
      let match = designPoints.find(dp => String(dp.p).trim().toLowerCase() === String(asPt.p).trim().toLowerCase());
      if (!match) {
        let bestDist = Infinity;
        designPoints.forEach(dp => {
          const d = Math.hypot(asPt.e - dp.e, asPt.n - dp.n);
          if (d < bestDist) {
            bestDist = d;
            match = dp;
          }
        });
      }

      if (match) {
        const dE_m = asPt.e - match.e;
        const dN_m = asPt.n - match.n;
        const dZ_m = (asPt.z || 0) - (match.z || 0);
        const dXY_m = Math.hypot(dE_m, dN_m);

        const dE_mm = dE_m * 1000;
        const dN_mm = dN_m * 1000;
        const dZ_mm = dZ_m * 1000;
        const dXY_mm = dXY_m * 1000;

        let az = (Math.atan2(dE_m, dN_m) * 180) / Math.PI;
        if (az < 0) az += 360;

        const isPass = dXY_mm <= toleranceMm;
        results.push({
          idx: idx + 1,
          p: asPt.p,
          code: asPt.code,
          designE: match.e,
          designN: match.n,
          designZ: match.z || 0,
          actualE: asPt.e,
          actualN: asPt.n,
          actualZ: asPt.z || 0,
          dE_mm,
          dN_mm,
          dXY_mm,
          dZ_mm,
          azimuth: az,
          status: isPass ? "مقبول" : "تجاوز السماحية",
          isPass
        });
      }
    });

    asBuiltComparisons = results;
    return results;
  }

  /* ============================================================================
     تنسيقات المحطة (CSS) بالسمات الثلاث
     ============================================================================ */
  function injectStyles() {
    if (document.getElementById("gdAiExcelV8Styles")) return;
    const style = document.createElement("style");
    style.id = "gdAiExcelV8Styles";
    style.textContent = `
      #gdAiExcelStudio {
        --ex-bg: #0b1628;
        --ex-card: #10203a;
        --ex-panel: #08101e;
        --ex-border: #1e3a5f;
        --ex-accent: #38bdf8;
        --ex-gold: #fbbf24;
        --ex-text: #f1f5f9;
        --ex-muted: #94a3b8;
        --ex-row-bg: #f8fafc;
        --ex-row-tx: #0f172a;
        background: var(--ex-bg);
        border: 1.5px solid var(--ex-border);
        border-radius: 16px;
        padding: 14px;
        margin-bottom: 16px;
        color: var(--ex-text);
        direction: rtl;
        text-align: right;
        box-shadow: 0 12px 34px rgba(0, 0, 0, 0.45);
      }

      #gdAiExcelStudio[data-ex-theme="gold"] {
        --ex-bg: #171411;
        --ex-card: #221d18;
        --ex-panel: #100e0c;
        --ex-border: #574323;
        --ex-accent: #fbbf24;
        --ex-gold: #f59e0b;
        --ex-text: #fafaf9;
        --ex-muted: #a8a29e;
      }

      #gdAiExcelStudio[data-ex-theme="light"] {
        --ex-bg: #f1f5f9;
        --ex-card: #ffffff;
        --ex-panel: #e2e8f0;
        --ex-border: #cbd5e1;
        --ex-accent: #0284c7;
        --ex-gold: #d97706;
        --ex-text: #0f172a;
        --ex-muted: #475569;
        --ex-row-bg: #ffffff;
        --ex-row-tx: #0f172a;
      }

      .gd-ex-top {
        display: flex; align-items: center; justify-content: space-between;
        flex-wrap: wrap; gap: 8px; padding-bottom: 10px;
        border-bottom: 1px solid var(--ex-border); margin-bottom: 12px;
      }
      .gd-ex-title {
        font-size: 16px; font-weight: 900; color: var(--ex-accent); margin: 0;
      }
      .gd-ex-sub {
        font-size: 11.5px; color: var(--ex-muted); margin-top: 2px;
      }
      .gd-ex-header-actions {
        display: flex; align-items: center; gap: 6px; flex-wrap: wrap;
      }
      .gd-theme-bulb-btn {
        padding: 6px 11px; border-radius: 999px;
        border: 1.5px solid var(--ex-accent); background: var(--ex-card);
        color: var(--ex-accent); font-size: 11.5px; font-weight: 900;
        cursor: pointer; font-family: inherit;
      }

      .gd-ex-cmd-box {
        display: flex; gap: 6px; margin-bottom: 10px;
      }
      .gd-ex-cmd-box input {
        flex: 1; padding: 10px 12px; border-radius: 11px;
        border: 1.5px solid var(--ex-border); background: var(--ex-card);
        color: var(--ex-text); font-size: 13px; margin: 0;
      }
      .gd-ex-btn {
        padding: 8px 12px; border-radius: 10px; border: 1px solid var(--ex-border);
        background: var(--ex-card); color: var(--ex-text); font-size: 12px;
        font-weight: 800; cursor: pointer; font-family: inherit; white-space: nowrap;
      }
      .gd-ex-btn.primary {
        background: var(--ex-accent); color: #090d16; border-color: var(--ex-accent); font-weight: 900;
      }
      .gd-ex-btn.gold {
        background: linear-gradient(135deg, #f59e0b, #d97706); color: #141210; border-color: #fbbf24; font-weight: 900;
      }
      .gd-ex-btn.green {
        background: linear-gradient(135deg, #16a34a, #15803d); color: #ffffff; border-color: #4ade80; font-weight: 900;
      }

      .gd-ex-chips {
        display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 10px;
      }
      .gd-ex-chip {
        padding: 5px 10px; border-radius: 999px; border: 1px solid var(--ex-border);
        background: var(--ex-panel); color: var(--ex-text); font-size: 11.5px;
        font-weight: 700; cursor: pointer; font-family: inherit;
      }
      .gd-ex-chip.active {
        background: var(--ex-accent); color: #090d16; border-color: var(--ex-accent); font-weight: 900;
      }

      .gd-ex-stats {
        display: grid; grid-template-columns: repeat(auto-fit, minmax(95px, 1fr));
        gap: 6px; margin-bottom: 10px;
      }
      .gd-ex-stat {
        background: var(--ex-card); border: 1px solid var(--ex-border);
        border-radius: 10px; padding: 7px 8px; text-align: center;
      }
      .gd-ex-stat span { display: block; font-size: 10.5px; color: var(--ex-muted); }
      .gd-ex-stat b { font-size: 14px; color: var(--ex-accent); }

      /* شريط أداة مقارنة التصميم بالرفع الجديد (As-Built vs Design) */
      .gd-asbuilt-bar {
        background: var(--ex-panel); border: 1.5px solid #fbbf24;
        border-radius: 12px; padding: 10px 12px; margin-bottom: 12px;
        display: flex; align-items: center; justify-content: space-between;
        flex-wrap: wrap; gap: 8px;
      }

      .gd-ex-table-wrap {
        max-height: 270px; overflow: auto; border-radius: 12px;
        border: 1px solid var(--ex-border); background: var(--ex-panel);
        margin-bottom: 12px;
      }
      .gd-ex-table {
        width: 100%; border-collapse: collapse; font-size: 12px; min-width: 560px;
      }
      .gd-ex-table th {
        position: sticky; top: 0; background: var(--ex-card); color: var(--ex-accent);
        padding: 8px 6px; border-bottom: 1px solid var(--ex-border); font-weight: 900; z-index: 2;
      }
      .gd-ex-table td {
        padding: 6px; text-align: center; border-bottom: 1px solid var(--ex-border);
      }
      .gd-ex-cell {
        width: 100%; background: var(--ex-row-bg); color: var(--ex-row-tx);
        border: 1px solid #cbd5e1; border-radius: 7px; padding: 5px 6px;
        font-weight: 800; font-size: 12px; text-align: center; margin: 0;
      }

      .gd-cad-wrap {
        border: 1.5px solid var(--ex-border); border-radius: 12px;
        background: #060a12; padding: 10px; margin-bottom: 12px;
      }
      .gd-cad-bar {
        display: flex; align-items: center; justify-content: space-between;
        flex-wrap: wrap; gap: 6px; margin-bottom: 8px;
      }
      #gdCadCanvas {
        width: 100%; height: 260px; display: block; border-radius: 9px;
        background: #05080f; border: 1px solid #1e293b; cursor: grab; touch-action: none;
      }

      .gd-export-grid {
        display: grid; grid-template-columns: repeat(auto-fit, minmax(145px, 1fr));
        gap: 7px; margin-top: 8px;
      }
      .gd-export-card {
        background: var(--ex-card); border: 1px solid var(--ex-border);
        border-radius: 10px; padding: 9px; text-align: right; cursor: pointer;
        color: var(--ex-text); font-family: inherit; transition: transform 0.15s ease;
      }
      .gd-export-card:hover { transform: translateY(-2px); border-color: var(--ex-accent); }
      .gd-export-card b { display: block; font-size: 12px; color: var(--ex-accent); margin-bottom: 2px; }
      .gd-export-card span { display: block; font-size: 10.5px; color: var(--ex-muted); }
    `;
    document.head.appendChild(style);
  }

  /* ============================================================================
     بناء واجهة محطة إكسيل المساحة
     ============================================================================ */
  function renderStation() {
    const surveySec = document.getElementById("s-survey");
    if (!surveySec) return;

    let host = document.getElementById("gdAiExcelStudio");
    if (!host) {
      host = document.createElement("div");
      host.id = "gdAiExcelStudio";
      surveySec.prepend(host);
    }
    host.setAttribute("data-ex-theme", currentTheme);

    const todayStr = new Date().toISOString().slice(0, 10);

    host.innerHTML = `
      <div class="gd-ex-top">
        <div>
          <h3 class="gd-ex-title">محطة إكسيل المساحة ونظم المعلومات الجغرافية (Survey & GIS Excel)</h3>
          <div class="gd-ex-sub">استيراد وتصنيف ومقارنة نقاط التصميم بالرفع الجديد وتصدير التقارير الهندسية و Civil 3D و AutoCAD</div>
        </div>
        <div class="gd-ex-header-actions">
          <button type="button" id="gdOpenReportModalBtn" class="gd-ex-btn green">
            تصدير التقرير الهندسي (PDF / Excel)
          </button>
          <button type="button" id="gdCycleThemeBulbBtn" class="gd-theme-bulb-btn">
            ${THEME_LABELS[currentTheme]}
          </button>
        </div>
      </div>

      <!-- شريط الأوامر الذكية -->
      <div class="gd-ex-cmd-box">
        <input id="gdExcelAiInput" type="text" placeholder="أدخل أمراً هندسياً: صنف الأكواد تلقائياً | افحص أخطاء الرفع | احسب المساحة | اجمع 0.50 على Z...">
        <button type="button" id="gdExcelAiExecBtn" class="gd-ex-btn primary">تنفيذ الأمر</button>
      </div>

      <!-- أزرار الاختصارات السريعة -->
      <div class="gd-ex-chips">
        <button type="button" class="gd-ex-chip" data-ai-cmd="classify">+ تصنيف الأكواد (MNH / LP / EG / COL)</button>
        <button type="button" class="gd-ex-chip" data-ai-cmd="qaqc">+ فحص الجودة (QA/QC)</button>
        <button type="button" class="gd-ex-chip" data-ai-cmd="area">+ المساحة بالفدان والمتر</button>
        <button type="button" class="gd-ex-chip" data-ai-cmd="shiftZ">+ إزاحة المنسوب (+0.50 Z)</button>
        <button type="button" class="gd-ex-chip" data-ai-cmd="swapEN">+ تبديل E ↔ N</button>
      </div>

      <!-- أزرار الاستيراد والتحكم -->
      <div class="gd-ex-chips">
        <button type="button" id="gdImportSurveyBtn" class="gd-ex-btn gold">استيراد ملف الرفع (Excel / CSV / DXF / SDR / STR / GSI)</button>
        <input id="gdImportSurveyFile" type="file" hidden accept=".csv,.txt,.tsv,.dxf,.sdr,.str,.gsi,.dat,.xls,.xlsx">
        <button type="button" id="gdAddPointRowBtn" class="gd-ex-btn">+ إضافة نقطة</button>
        <button type="button" id="gdUndoExcelBtn" class="gd-ex-btn">↩ تراجع</button>
        <button type="button" id="gdClearExcelBtn" class="gd-ex-btn" style="color:#f87171;">✕ تفريغ الجدول</button>
      </div>

      <!-- أداة مقارنة نقاط التصميم (القديمة) بنقاط الرفع الجديد (As-Built vs Design) -->
      <div class="gd-asbuilt-bar">
        <div>
          <strong style="font-size:13px;color:#fbbf24;">مقارنة التوقيع التصميمي بالرفع التنفيذي (As-Built vs Design)</strong>
          <div style="font-size:11px;color:var(--ex-muted);" id="gdAsBuiltStatusText">
            ${designPoints.length
              ? `تم تحميل ${designPoints.length} نقطة تصميمية — جاهز لحساب التفاوتات بالمليمتر`
              : "ارفع ملف نقاط التصميم/الأعمدة الأصلي لمقارنته بنقاط الرفع الجديد وحساب فروق الإزاحة"}
          </div>
        </div>
        <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;">
          <button type="button" id="gdImportDesignBtn" class="gd-ex-btn gold">رفع نقاط التصميم (القديمة)</button>
          <input id="gdImportDesignFile" type="file" hidden accept=".csv,.txt,.tsv,.dxf,.sdr,.str,.gsi,.dat">
          <button type="button" id="gdRunAsBuiltBtn" class="gd-ex-btn primary" ${designPoints.length ? "" : "disabled"}>
            فحص التفاوتات والترحيل
          </button>
          ${designPoints.length ? `<button type="button" id="gdClearDesignBtn" class="gd-ex-btn" style="color:#f87171;">مسح التصميم</button>` : ""}
        </div>
      </div>

      <!-- فلاتر الفئات -->
      <div class="gd-ex-chips" id="gdCatFilterBar"></div>

      <!-- بطاقات الإحصائيات -->
      <div class="gd-ex-stats" id="gdExcelStatsGrid"></div>

      <!-- جدول الإحداثيات والتصنيف -->
      <div class="gd-ex-table-wrap">
        <table class="gd-ex-table">
          <thead>
            <tr>
              <th>التصنيف الهندسي</th>
              <th>P</th>
              <th>E (شرق)</th>
              <th>N (شمال)</th>
              <th>Z (منسوب)</th>
              <th>Code</th>
              <th>تفاوت XY</th>
              <th>حذف</th>
            </tr>
          </thead>
          <tbody id="gdExcelTableBody"></tbody>
        </table>
      </div>

      <!-- معاينة الكاد الحية بالرموز الهندسية وأسهم الترحيل -->
      <div class="gd-cad-wrap">
        <div class="gd-cad-bar">
          <strong style="font-size:12.5px;color:#38bdf8;">معاينة CAD بالرموز الهندسية وأسهم ترحيل As-Built</strong>
          <div style="display:flex;gap:5px;flex-wrap:wrap;">
            <button type="button" id="gdToggleLabelsBtn" class="gd-ex-chip ${showLabels ? "active" : ""}">النصوص</button>
            <button type="button" id="gdToggleConnectBtn" class="gd-ex-chip ${connectByCode ? "active" : ""}">التوصيل حسب الكود</button>
            <button type="button" id="gdCadZoomInBtn" class="gd-ex-chip">+ تقريب</button>
            <button type="button" id="gdCadZoomOutBtn" class="gd-ex-chip">- تبعيد</button>
            <button type="button" id="gdCadZoomResetBtn" class="gd-ex-chip">توسيط الكل</button>
          </div>
        </div>
        <canvas id="gdCadCanvas"></canvas>
      </div>

      <!-- مركز التصدير والحفظ الاحترافي -->
      <div>
        <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:6px;">
          <strong style="font-size:13px;color:var(--ex-accent);">مركز التصدير والحفظ الاحترافي (Civil 3D / AutoCAD / Total Station / GIS)</strong>
          <select id="gdExportPointOrder" style="width:auto;margin:0;padding:6px 10px;font-size:12px;">
            <option value="PENZD">PENZD — (نقطة، شرق E، شمال N، منسوب Z، وصف D)</option>
            <option value="PNEZD">PNEZD — (نقطة، شمال N، شرق E، منسوب Z، وصف D)</option>
            <option value="PENZ">PENZ — (بدون وصف)</option>
          </select>
        </div>

        <div class="gd-export-grid">
          <button type="button" class="gd-export-card" data-export-fmt="civil3d">
            <b>Civil 3D (.csv)</b>
            <span>بدون BOM — يفتح بالـ Point Style مباشرة</span>
          </button>
          <button type="button" class="gd-export-card" data-export-fmt="txt">
            <b>ملف نصي قياسي (.txt)</b>
            <span>صيغة نصية نظيفة للأجهزة والبرامج</span>
          </button>
          <button type="button" class="gd-export-card" data-export-fmt="excel">
            <b>ملف Excel نظيف (.csv)</b>
            <span>يبدأ من أول نقطة مباشرة بلا عناوين</span>
          </button>
          <button type="button" class="gd-export-card" data-export-fmt="dxf">
            <b>لوحة AutoCAD (.dxf)</b>
            <span>طبقات ورموز هندسية جاهزة</span>
          </button>
          <button type="button" class="gd-export-card" data-export-fmt="sokkia_str">
            <b>توتال ستيشن Sokkia (.str)</b>
            <span>صيغة SOKKIA للتوقيع المساحي</span>
          </button>
          <button type="button" class="gd-export-card" data-export-fmt="leica_gsi">
            <b>توتال ستيشن Leica (.gsi)</b>
            <span>ملف لايكا جاهز للرفع على الجهاز</span>
          </button>
          <button type="button" class="gd-export-card" data-export-fmt="surfer_dat">
            <b>برنامج Surfer (.dat)</b>
            <span>XYZ للكنتور وحساب الكميات</span>
          </button>
          <button type="button" class="gd-export-card" data-export-fmt="geojson">
            <b>ArcGIS Pro / QGIS (.geojson)</b>
            <span>طبقة GIS ثلاثية الأبعاد</span>
          </button>
        </div>
      </div>
    `;

    refreshAllViews();
  }

  /* ============================================================================
     تحديث الفلاتر والإحصائيات والجدول ورسم الكاد
     ============================================================================ */
  function refreshAllViews() {
    renderFiltersAndStats();
    renderTableRows();
    drawCadCanvas();
  }

  function renderFiltersAndStats() {
    const counts = { ALL: points.length, MNH: 0, EG: 0, LP: 0, TREE: 0, COL: 0, ROAD: 0, OTHER: 0 };
    points.forEach(pt => {
      const cat = classifyCode(pt.code).cat;
      pt.cat = cat;
      counts[cat] = (counts[cat] || 0) + 1;
    });

    const filterBar = document.getElementById("gdCatFilterBar");
    if (filterBar) {
      const btns = [
        ["ALL", `الكل (${counts.ALL})`],
        ["MNH", `المطابق والصرف (${counts.MNH})`],
        ["EG", `الأرض الطبيعية (${counts.EG})`],
        ["LP", `أعمدة الإنارة (${counts.LP})`],
        ["TREE", `الأشجار (${counts.TREE})`],
        ["COL", `الأعمدة الخرسانية (${counts.COL})`]
      ];
      filterBar.innerHTML = btns.map(([k, lbl]) =>
        `<button type="button" class="gd-ex-chip ${activeFilter === k ? "active" : ""}" data-cat-filter="${k}">${lbl}</button>`
      ).join("");
    }

    // حساب المساحة التقريبية
    let areaM2 = 0;
    if (points.length >= 3) {
      for (let i = 0; i < points.length; i++) {
        const j = (i + 1) % points.length;
        areaM2 += points[i].e * points[j].n - points[j].e * points[i].n;
      }
      areaM2 = Math.abs(areaM2) / 2;
    }

    const statsGrid = document.getElementById("gdExcelStatsGrid");
    if (statsGrid) {
      const rejectedCount = asBuiltComparisons.filter(c => !c.isPass).length;
      statsGrid.innerHTML = `
        <div class="gd-ex-stat"><span>إجمالي النقاط</span><b>${points.length}</b></div>
        <div class="gd-ex-stat"><span>مطابق (MNH)</span><b>${counts.MNH}</b></div>
        <div class="gd-ex-stat"><span>أرض طبيعية (EG)</span><b>${counts.EG}</b></div>
        <div class="gd-ex-stat"><span>أعمدة إنارة (LP)</span><b>${counts.LP}</b></div>
        <div class="gd-ex-stat"><span>أشجار (TREE)</span><b>${counts.TREE}</b></div>
        <div class="gd-ex-stat"><span>المساحة المغطاة</span><b>${areaM2.toFixed(1)} م²</b></div>
        ${asBuiltComparisons.length ? `<div class="gd-ex-stat" style="border-color:${rejectedCount ? "#ef4444" : "#22c55e"};"><span>تجاوز السماحية</span><b style="color:${rejectedCount ? "#f87171" : "#4ade80"};">${rejectedCount} / ${asBuiltComparisons.length}</b></div>` : ""}
      `;
    }
  }

  function renderTableRows() {
    const tbody = document.getElementById("gdExcelTableBody");
    if (!tbody) return;

    const filtered = points
      .map((pt, idx) => ({ pt, idx }))
      .filter(({ pt }) => activeFilter === "ALL" || pt.cat === activeFilter);

    tbody.innerHTML = filtered.map(({ pt, idx }) => {
      const info = classifyCode(pt.code);
      const comp = asBuiltComparisons.find(c => c.idx === idx + 1);
      const compBadge = comp
        ? `<span style="display:inline-block;padding:3px 7px;border-radius:6px;font-weight:900;font-size:11px;background:${comp.isPass ? "rgba(34,197,94,0.18)" : "rgba(239,68,68,0.22)"};color:${comp.isPass ? "#4ade80" : "#f87171"};">${comp.dXY_mm.toFixed(1)} مم</span>`
        : `<span style="color:var(--ex-muted);font-size:11px;">—</span>`;

      return `
        <tr>
          <td style="font-weight:900;color:${info.color};white-space:nowrap;">${info.label}</td>
          <td><input class="gd-ex-cell" data-row="${idx}" data-field="p" value="${pt.p}"></td>
          <td><input class="gd-ex-cell" type="number" step="any" data-row="${idx}" data-field="e" value="${pt.e}"></td>
          <td><input class="gd-ex-cell" type="number" step="any" data-row="${idx}" data-field="n" value="${pt.n}"></td>
          <td><input class="gd-ex-cell" type="number" step="any" data-row="${idx}" data-field="z" value="${pt.z}"></td>
          <td><input class="gd-ex-cell" data-row="${idx}" data-field="code" value="${pt.code}"></td>
          <td>${compBadge}</td>
          <td><button type="button" data-del-row="${idx}" style="background:#7f1d1d;color:#fecaca;border:0;border-radius:6px;padding:4px 8px;cursor:pointer;">✕</button></td>
        </tr>
      `;
    }).join("");
  }

  /* ============================================================================
     رسم معاينة CAD بالرموز وأسهم ترحيل As-Built
     ============================================================================ */
  function drawCadCanvas() {
    const canvas = document.getElementById("gdCadCanvas");
    if (!canvas) return;
    const rect = canvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(300, rect.width) * dpr;
    canvas.height = 260 * dpr;
    const ctx = canvas.getContext("2d");
    ctx.scale(dpr, dpr);

    const W = rect.width || 320;
    const H = 260;

    ctx.fillStyle = "#05080f";
    ctx.fillRect(0, 0, W, H);

    // شبكة إحداثيات خفيفة
    ctx.strokeStyle = "rgba(56, 189, 248, 0.08)";
    ctx.lineWidth = 1;
    for (let x = 0; x < W; x += 28) {
      ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, H); ctx.stroke();
    }
    for (let y = 0; y < H; y += 28) {
      ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(W, y); ctx.stroke();
    }

    const allPts = [...points, ...designPoints];
    if (!allPts.length) return;

    let minE = Infinity, maxE = -Infinity, minN = Infinity, maxN = -Infinity;
    allPts.forEach(pt => {
      if (pt.e < minE) minE = pt.e;
      if (pt.e > maxE) maxE = pt.e;
      if (pt.n < minN) minN = pt.n;
      if (pt.n > maxN) maxN = pt.n;
    });

    const spanE = Math.max(1, maxE - minE);
    const spanN = Math.max(1, maxN - minN);
    const pad = 38;
    const baseScale = Math.min((W - pad * 2) / spanE, (H - pad * 2) / spanN);
    const scale = baseScale * cadZoom;

    const midE = (minE + maxE) / 2;
    const midN = (minN + maxN) / 2;
    const toScreen = (e, n) => ({
      x: W / 2 + (e - midE) * scale + cadPanX,
      y: H / 2 - (n - midN) * scale + cadPanY
    });

    // رسم خط التوصيل حسب الكود إذا كان مفعلاً
    if (connectByCode && points.length > 1) {
      const grouped = {};
      points.forEach(pt => {
        const k = pt.cat || "EG";
        if (!grouped[k]) grouped[k] = [];
        grouped[k].push(pt);
      });
      Object.values(grouped).forEach(arr => {
        if (arr.length < 2) return;
        ctx.strokeStyle = classifyCode(arr[0].code).color;
        ctx.lineWidth = 1.3;
        ctx.beginPath();
        arr.forEach((pt, idx) => {
          const s = toScreen(pt.e, pt.n);
          if (idx === 0) ctx.moveTo(s.x, s.y);
          else ctx.lineTo(s.x, s.y);
        });
        ctx.stroke();
      });
    }

    // رسم نقاط التصميم وأسهم الترحيل (As-Built Shift Vectors)
    asBuiltComparisons.forEach(c => {
      const pDes = toScreen(c.designE, c.designN);
      const pAct = toScreen(c.actualE, c.actualN);

      // نقطة التصميم (مربع سماوي مفرغ)
      ctx.strokeStyle = "#38bdf8";
      ctx.lineWidth = 1.5;
      ctx.strokeRect(pDes.x - 4, pDes.y - 4, 8, 8);

      // خط/سهم الإزاحة
      ctx.strokeStyle = c.isPass ? "#22c55e" : "#ef4444";
      ctx.lineWidth = 1.6;
      ctx.setLineDash([3, 2]);
      ctx.beginPath();
      ctx.moveTo(pDes.x, pDes.y);
      ctx.lineTo(pAct.x, pAct.y);
      ctx.stroke();
      ctx.setLineDash([]);
    });

    // رسم النقاط المرفوعة برموزها
    points.forEach(pt => {
      if (activeFilter !== "ALL" && pt.cat !== activeFilter) return;
      const s = toScreen(pt.e, pt.n);
      const info = classifyCode(pt.code);

      ctx.fillStyle = info.color;
      ctx.strokeStyle = info.color;
      ctx.lineWidth = 1.6;

      if (info.cat === "MNH") {
        ctx.beginPath(); ctx.arc(s.x, s.y, 4.5, 0, Math.PI * 2); ctx.stroke();
      } else if (info.cat === "LP") {
        ctx.beginPath();
        ctx.moveTo(s.x, s.y - 5); ctx.lineTo(s.x + 4, s.y);
        ctx.lineTo(s.x, s.y + 5); ctx.lineTo(s.x - 4, s.y);
        ctx.closePath(); ctx.fill();
      } else if (info.cat === "TREE") {
        ctx.beginPath(); ctx.arc(s.x, s.y, 4, 0, Math.PI * 2); ctx.fill();
      } else if (info.cat === "COL") {
        ctx.fillRect(s.x - 4, s.y - 4, 8, 8);
      } else {
        ctx.beginPath();
        ctx.moveTo(s.x, s.y - 4.5); ctx.lineTo(s.x + 4, s.y + 3.5); ctx.lineTo(s.x - 4, s.y + 3.5);
        ctx.closePath(); ctx.stroke();
      }

      if (showLabels) {
        ctx.fillStyle = "#e2e8f0";
        ctx.font = "bold 10px sans-serif";
        ctx.fillText(`Z=${pt.z} [${pt.code}] ${pt.p}`, s.x + 7, s.y + 3);
      }
    });
  }

  /* ============================================================================
     نافذة "تصدير التقرير الهندسي" ببيانات المشروع والمهندس والسماحية
     ============================================================================ */
  function openReportConfigModal() {
    const old = document.getElementById("gdEngineeringReportModal");
    if (old) old.remove();

    const defaultEngName =
      sessionStorage.getItem("gd_name") && sessionStorage.getItem("gd_name") !== "ضيف"
        ? sessionStorage.getItem("gd_name")
        : "م. أحمد سمير";
    const todayDate = new Date().toISOString().slice(0, 10);

    const overlay = document.createElement("div");
    overlay.id = "gdEngineeringReportModal";
    overlay.style.cssText = `
      position: fixed; inset: 0; z-index: 999999;
      background: rgba(10, 9, 8, 0.88); backdrop-filter: blur(7px);
      display: flex; align-items: center; justify-content: center;
      padding: 14px; direction: rtl; font-family: inherit;
    `;

    overlay.innerHTML = `
      <div style="width:100%;max-width:480px;background:linear-gradient(165deg,#1c1917,#141210);border:2px solid #e09f24;border-radius:18px;padding:18px;color:#f5f5f4;box-shadow:0 24px 60px rgba(0,0,0,0.85);max-height:92vh;overflow-y:auto;">
        <div style="display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #3f3a36;padding-bottom:10px;margin-bottom:12px;">
          <div>
            <span style="display:inline-block;background:rgba(224,159,36,0.16);border:1px solid #e09f24;color:#fbbf24;font-size:11px;font-weight:900;padding:2px 9px;border-radius:999px;margin-bottom:3px;">
              تقرير مساحي وهندسي معتمد
            </span>
            <h3 style="margin:0;font-size:17px;font-weight:900;color:#fbbf24;">بيانات تصدير التقرير الهندسي</h3>
          </div>
          <button type="button" id="gdCloseRepModalBtn" style="width:32px;height:32px;border-radius:8px;border:1px solid #44403c;background:#292524;color:#d6d3d1;font-weight:900;cursor:pointer;">✕</button>
        </div>

        <div style="display:flex;flex-direction:column;gap:10px;font-size:13px;">
          <div>
            <label style="display:block;font-weight:800;color:#fbbf24;margin-bottom:4px;">اسم المشروع *</label>
            <input id="repProjectName" type="text" placeholder="مثال: مشروع إنشاء مبنى إداري / شبكة صرف صحي"
              style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;margin:0;">
          </div>

          <div>
            <label style="display:block;font-weight:800;color:#fbbf24;margin-bottom:4px;">نوع الأعمال *</label>
            <input id="repWorkType" type="text" placeholder="مثال: رفع تنفيذي ومراجعة توقيع الأعمدة (As-Built)"
              style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;margin:0;">
          </div>

          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
            <div>
              <label style="display:block;font-weight:800;color:#fbbf24;margin-bottom:4px;">تاريخ الرفع *</label>
              <input id="repSurveyDate" type="date" value="${todayDate}"
                style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;margin:0;">
            </div>
            <div>
              <label style="display:block;font-weight:800;color:#fbbf24;margin-bottom:4px;">حد السماحية المعتمد (مم) *</label>
              <input id="repToleranceMm" type="number" step="any" value="20" placeholder="مثال: 20 مم"
                style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;margin:0;">
            </div>
          </div>

          <div>
            <label style="display:block;font-weight:800;color:#fbbf24;margin-bottom:4px;">اسم المهندس المسؤول *</label>
            <input id="repEngineerName" type="text" value="${defaultEngName}" placeholder="اكتب اسم المهندس المسؤول"
              style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;margin:0;">
          </div>

          <div>
            <label style="display:block;font-weight:800;color:#fbbf24;margin-bottom:4px;">ملاحظات واعتماد الاستشاري (اختياري)</label>
            <input id="repApprovalNote" type="text" value="مطابق للمواصفات الهندسية وحدود السماحية المعتمدة"
              style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;margin:0;">
          </div>
        </div>

        <div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:16px;">
          <button type="button" id="gdGeneratePrintReportBtn"
            style="flex:1;min-width:160px;padding:12px;border-radius:11px;border:0;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-size:13.5px;font-weight:900;cursor:pointer;font-family:inherit;">
            معاينة وحفظ التقرير PDF
          </button>
          <button type="button" id="gdGenerateExcelReportBtn"
            style="flex:1;min-width:150px;padding:12px;border-radius:11px;border:1.5px solid #4ade80;background:#14532d;color:#ffffff;font-size:13.5px;font-weight:900;cursor:pointer;font-family:inherit;">
            تصدير التقرير Excel
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);

    document.getElementById("gdCloseRepModalBtn").onclick = () => overlay.remove();
    overlay.addEventListener("click", e => { if (e.target === overlay) overlay.remove(); });

    const collectMeta = () => {
      const tol = parseFloat(document.getElementById("repToleranceMm").value) || 20;
      if (designPoints.length) computeAsBuiltComparison(tol);
      return {
        projectName: document.getElementById("repProjectName").value.trim() || "مشروع رفع مساحي وتنفيذي",
        workType: document.getElementById("repWorkType").value.trim() || "أعمال رفع وتوقيع مساحي",
        surveyDate: document.getElementById("repSurveyDate").value || todayDate,
        engineerName: document.getElementById("repEngineerName").value.trim() || defaultEngName,
        toleranceMm: tol,
        approvalNote: document.getElementById("repApprovalNote").value.trim()
      };
    };

    document.getElementById("gdGeneratePrintReportBtn").onclick = () => {
      const meta = collectMeta();
      overlay.remove();
      exportPrintablePdfReport(meta);
    };

    document.getElementById("gdGenerateExcelReportBtn").onclick = () => {
      const meta = collectMeta();
      overlay.remove();
      exportExcelReportCsv(meta);
    };
  }

  /* ============================================================================
     توليد التقرير الطباعي الرسمي (PDF / Print Window)
     ============================================================================ */
  function exportPrintablePdfReport(meta) {
    const hasComp = asBuiltComparisons.length > 0;
    const passed = asBuiltComparisons.filter(c => c.isPass).length;
    const failed = asBuiltComparisons.filter(c => !c.isPass).length;
    const maxShift = hasComp ? Math.max(...asBuiltComparisons.map(c => c.dXY_mm)) : 0;

    const tableHeader = hasComp
      ? `<tr>
          <th>#</th><th>رقم النقطة</th><th>الوصف</th>
          <th>E التصميمي</th><th>N التصميمي</th>
          <th>E المرفوع</th><th>N المرفوع</th>
          <th>ΔE (مم)</th><th>ΔN (مم)</th><th>الإزاحة ΔXY (مم)</th><th>ΔZ (مم)</th><th>الحالة</th>
        </tr>`
      : `<tr>
          <th>#</th><th>رقم النقطة (P)</th><th>الإحداثي الشرقي (E)</th><th>الإحداثي الشمالي (N)</th><th>المنسوب (Z)</th><th>الكود</th><th>التصنيف الهندسي</th>
        </tr>`;

    const tableRows = hasComp
      ? asBuiltComparisons.map(c => `
          <tr>
            <td>${c.idx}</td>
            <td><b>${c.p}</b></td>
            <td>${c.code}</td>
            <td>${c.designE.toFixed(3)}</td>
            <td>${c.designN.toFixed(3)}</td>
            <td>${c.actualE.toFixed(3)}</td>
            <td>${c.actualN.toFixed(3)}</td>
            <td>${c.dE_mm.toFixed(1)}</td>
            <td>${c.dN_mm.toFixed(1)}</td>
            <td><b>${c.dXY_mm.toFixed(1)}</b></td>
            <td>${c.dZ_mm.toFixed(1)}</td>
            <td style="font-weight:bold;color:${c.isPass ? "#15803d" : "#b91c1c"};">${c.status}</td>
          </tr>
        `).join("")
      : points.map((pt, i) => `
          <tr>
            <td>${i + 1}</td>
            <td><b>${pt.p}</b></td>
            <td>${Number(pt.e).toFixed(3)}</td>
            <td>${Number(pt.n).toFixed(3)}</td>
            <td>${Number(pt.z).toFixed(3)}</td>
            <td>${pt.code}</td>
            <td>${classifyCode(pt.code).label}</td>
          </tr>
        `).join("");

    const reportHtml = `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<title>تقرير مساحي وهندسي - ${meta.projectName}</title>
<style>
  body { font-family: Tahoma, Arial, sans-serif; direction: rtl; color: #0f172a; margin: 24px; background: #fff; }
  .header { border: 2px solid #0f172a; border-radius: 10px; padding: 14px 18px; margin-bottom: 16px; background: #f8fafc; }
  .header h1 { margin: 0 0 6px; font-size: 20px; color: #0f172a; }
  .meta-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px 16px; font-size: 13px; margin-top: 10px; }
  .meta-item b { color: #b45309; }
  .summary { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
  .sum-box { flex: 1; min-width: 130px; border: 1px solid #cbd5e1; border-radius: 8px; padding: 10px; text-align: center; background: #f8fafc; }
  .sum-box span { display: block; font-size: 11px; color: #475569; }
  .sum-box b { font-size: 16px; color: #0f172a; }
  table { width: 100%; border-collapse: collapse; font-size: 11.5px; margin-bottom: 20px; }
  th, td { border: 1px solid #94a3b8; padding: 6px; text-align: center; }
  th { background: #0f172a; color: #fff; }
  tr:nth-child(even) { background: #f8fafc; }
  .footer-sign { display: flex; justify-content: space-between; margin-top: 28px; font-size: 13px; font-weight: bold; }
  @media print { .no-print { display: none; } }
</style>
</head>
<body>
  <div class="no-print" style="margin-bottom:14px;text-align:left;">
    <button onclick="window.print()" style="padding:10px 18px;background:#0f172a;color:#fbbf24;border:0;border-radius:8px;font-weight:bold;cursor:pointer;">
      طباعة / حفظ كـ PDF
    </button>
  </div>

  <div class="header">
    <h1>التقرير المساحي والهندسي المعتمد — GeoDev Studio Pro</h1>
    <div class="meta-grid">
      <div class="meta-item"><b>اسم المشروع:</b> ${meta.projectName}</div>
      <div class="meta-item"><b>نوع الأعمال:</b> ${meta.workType}</div>
      <div class="meta-item"><b>تاريخ الرفع:</b> ${meta.surveyDate}</div>
      <div class="meta-item"><b>اسم المهندس المسؤول:</b> ${meta.engineerName}</div>
      <div class="meta-item"><b>حد السماحية المعتمد:</b> ± ${meta.toleranceMm} مم</div>
      <div class="meta-item"><b>ملاحظات الاعتماد:</b> ${meta.approvalNote || "—"}</div>
    </div>
  </div>

  <div class="summary">
    <div class="sum-box"><span>إجمالي النقاط المرفوعة</span><b>${points.length}</b></div>
    ${hasComp ? `
      <div class="sum-box"><span>النقاط المطابقة للسماحية</span><b style="color:#15803d;">${passed}</b></div>
      <div class="sum-box"><span>النقاط المتجاوزة للسماحية</span><b style="color:#b91c1c;">${failed}</b></div>
      <div class="sum-box"><span>أقصى إزاحة أفقية مسجلة</span><b>${maxShift.toFixed(1)} مم</b></div>
    ` : ""}
  </div>

  <table>
    <thead>${tableHeader}</thead>
    <tbody>${tableRows}</tbody>
  </table>

  <div class="footer-sign">
    <div>إعداد المهندس المساحي: ${meta.engineerName}</div>
    <div>اعتماد الاستشاري / مدير المشروع: ............................</div>
  </div>
</body>
</html>`;

    const win = window.open("", "_blank");
    if (win) {
      win.document.write(reportHtml);
      win.document.close();
    } else {
      downloadTextFile(reportHtml, `GeoDev_Report_${meta.surveyDate}.html`, "text/html;charset=utf-8");
    }
  }

  /* ============================================================================
     تصدير التقرير بصيغة Excel (CSV بترميز UTF-8 BOM)
     ============================================================================ */
  function exportExcelReportCsv(meta) {
    const bom = "\uFEFF";
    const lines = [
      `اسم المشروع,${meta.projectName}`,
      `نوع الأعمال,${meta.workType}`,
      `تاريخ الرفع,${meta.surveyDate}`,
      `اسم المهندس المسؤول,${meta.engineerName}`,
      `حد السماحية المعتمد (مم),${meta.toleranceMm}`,
      `ملاحظات الاعتماد,${meta.approvalNote}`,
      ""
    ];

    if (asBuiltComparisons.length) {
      lines.push("م,رقم النقطة,الكود,E التصميمي,N التصميمي,E المرفوع,N المرفوع,فرق E (مم),فرق N (مم),الازاحة الكلية XY (مم),فرق المنسوب Z (مم),الحالة");
      asBuiltComparisons.forEach(c => {
        lines.push(`${c.idx},${c.p},${c.code},${c.designE.toFixed(3)},${c.designN.toFixed(3)},${c.actualE.toFixed(3)},${c.actualN.toFixed(3)},${c.dE_mm.toFixed(1)},${c.dN_mm.toFixed(1)},${c.dXY_mm.toFixed(1)},${c.dZ_mm.toFixed(1)},${c.status}`);
      });
    } else {
      lines.push("م,رقم النقطة (P),الشرق (E),الشمال (N),المنسوب (Z),الكود,التصنيف الهندسي");
      points.forEach((pt, i) => {
        lines.push(`${i + 1},${pt.p},${Number(pt.e).toFixed(3)},${Number(pt.n).toFixed(3)},${Number(pt.z).toFixed(3)},${pt.code},${classifyCode(pt.code).label}`);
      });
    }

    downloadTextFile(bom + lines.join("\n"), `GeoDev_Report_${meta.surveyDate}.csv`, "text/csv;charset=utf-8");
  }

  /* ============================================================================
     التصدير الهندسي للبرامج والأجهزة
     ============================================================================ */
  function downloadTextFile(content, filename, mime = "text/plain;charset=utf-8") {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function handleExportFormat(fmt) {
    if (!points.length) return;
    const order = (document.getElementById("gdExportPointOrder")?.value) || "PENZD";

    const formatRow = pt => {
      if (order === "PNEZD") return `${pt.p},${pt.n},${pt.e},${pt.z},${pt.code}`;
      if (order === "PENZ") return `${pt.p},${pt.e},${pt.n},${pt.z}`;
      return `${pt.p},${pt.e},${pt.n},${pt.z},${pt.code}`;
    };

    if (fmt === "civil3d") {
      // بدون BOM كما يفضل Civil 3D
      downloadTextFile(points.map(formatRow).join("\n"), "GeoDev_Civil3D.csv");
    } else if (fmt === "txt") {
      downloadTextFile(points.map(formatRow).join("\n"), "GeoDev_Points.txt");
    } else if (fmt === "excel") {
      downloadTextFile("\uFEFF" + points.map(formatRow).join("\n"), "GeoDev_Clean_Excel.csv");
    } else if (fmt === "surfer_dat") {
      downloadTextFile(points.map(pt => `${pt.e} ${pt.n} ${pt.z}`).join("\n"), "GeoDev_Surfer.dat");
    } else if (fmt === "sokkia_str") {
      const str = points.map(pt => `${pt.p},${pt.n},${pt.e},${pt.z},${pt.code}`).join("\n");
      downloadTextFile(str, "GeoDev_Sokkia.str");
    } else if (fmt === "leica_gsi") {
      const gsi = points.map((pt, i) => {
        const idx = String(i + 1).padStart(6, "0");
        const e = Math.round(pt.e * 1000).toString().padStart(16, "0");
        const n = Math.round(pt.n * 1000).toString().padStart(16, "0");
        const z = Math.round(pt.z * 1000).toString().padStart(16, "0");
        return `11${idx}+000000000000000${pt.p} 81..10+${e} 82..10+${n} 83..10+${z}`;
      }).join("\n");
      downloadTextFile(gsi, "GeoDev_Leica.gsi");
    } else if (fmt === "dxf") {
      let dxf = "0\nSECTION\n2\nENTITIES\n";
      points.forEach(pt => {
        const layer = pt.cat || "SURVEY";
        dxf += `0\nPOINT\n8\n${layer}\n10\n${pt.e}\n20\n${pt.n}\n30\n${pt.z}\n`;
        dxf += `0\nTEXT\n8\n${layer}_TXT\n10\n${pt.e + 0.4}\n20\n${pt.n + 0.4}\n30\n${pt.z}\n40\n0.5\n1\n${pt.p} (${pt.code})\n`;
      });
      dxf += "0\nENDSEC\n0\nEOF\n";
      downloadTextFile(dxf, "GeoDev_AutoCAD.dxf");
    } else if (fmt === "geojson") {
      const gj = {
        type: "FeatureCollection",
        features: points.map(pt => ({
          type: "Feature",
          properties: { point: pt.p, code: pt.code, category: pt.cat, elevation: pt.z },
          geometry: { type: "Point", coordinates: [Number(pt.e), Number(pt.n), Number(pt.z)] }
        }))
      };
      downloadTextFile(JSON.stringify(gj, null, 2), "GeoDev_GIS.geojson", "application/json");
    }
  }

  /* ============================================================================
     تنفيذ الأوامر الذكية على الجدول
     ============================================================================ */
  function executeSmartCommand(cmdType, rawText = "") {
    pushUndo();
    const txt = String(rawText || cmdType).trim();

    if (cmdType === "classify" || txt.includes("صنف")) {
      points.forEach(pt => {
        pt.cat = classifyCode(pt.code).cat;
      });
    } else if (cmdType === "swapEN" || txt.includes("تبديل")) {
      points.forEach(pt => {
        const tmp = pt.e; pt.e = pt.n; pt.n = tmp;
      });
    } else if (cmdType === "shiftZ" || txt.includes("اجمع") || txt.includes("ازاحة") || txt.includes("إزاحة")) {
      const m = txt.match(/([+-]?\d+(?:\.\d+)?)/);
      const delta = m ? parseFloat(m[1]) : 0.5;
      points.forEach(pt => {
        pt.z = Number((Number(pt.z) + delta).toFixed(3));
      });
    } else if (cmdType === "qaqc" || txt.includes("فحص") || txt.includes("أخطاء")) {
      const seen = new Set();
      let dups = 0;
      points.forEach(pt => {
        const key = `${pt.e}_${pt.n}`;
        if (seen.has(key)) dups++;
        seen.add(key);
      });
      if (typeof toast === "function") {
        toast(dups ? `تم اكتشاف ${dups} نقطة مكررة الإحداثيات` : "فحص الجودة سليم: لا توجد نقاط مكررة", dups ? "warn" : "success");
      }
    } else if (cmdType === "area" || txt.includes("مساحة")) {
      let area = 0;
      for (let i = 0; i < points.length; i++) {
        const j = (i + 1) % points.length;
        area += points[i].e * points[j].n - points[j].e * points[i].n;
      }
      area = Math.abs(area) / 2;
      const feddan = area / 4200.83;
      if (typeof toast === "function") {
        toast(`المساحة = ${area.toFixed(2)} م² (${feddan.toFixed(4)} فدان)`, "success");
      }
    }

    refreshAllViews();
  }

  /* ============================================================================
     ربط الأحداث
     ============================================================================ */
  function bindEvents() {
    document.addEventListener("click", e => {
      // زر تصدير التقرير الهندسي
      if (e.target.closest("#gdOpenReportModalBtn")) {
        openReportConfigModal();
        return;
      }

      // تبديل سمة المحطة (3 سمات)
      if (e.target.closest("#gdCycleThemeBulbBtn")) {
        const nextIdx = (THEMES.indexOf(currentTheme) + 1) % THEMES.length;
        currentTheme = THEMES[nextIdx];
        localStorage.setItem(THEME_STORAGE_KEY, currentTheme);
        const host = document.getElementById("gdAiExcelStudio");
        if (host) host.setAttribute("data-ex-theme", currentTheme);
        const btn = document.getElementById("gdCycleThemeBulbBtn");
        if (btn) btn.textContent = THEME_LABELS[currentTheme];
        return;
      }

      // تنفيذ أمر ذكي
      if (e.target.closest("#gdExcelAiExecBtn")) {
        const val = document.getElementById("gdExcelAiInput")?.value || "";
        executeSmartCommand("custom", val);
        return;
      }
      const chipCmd = e.target.closest("[data-ai-cmd]");
      if (chipCmd) {
        executeSmartCommand(chipCmd.dataset.aiCmd);
        return;
      }

      // استيراد ملف الرفع الجديد
      if (e.target.closest("#gdImportSurveyBtn")) {
        document.getElementById("gdImportSurveyFile")?.click();
        return;
      }

      // استيراد ملف التصميم (القديم) للمقارنة
      if (e.target.closest("#gdImportDesignBtn")) {
        document.getElementById("gdImportDesignFile")?.click();
        return;
      }

      // تشغيل مقارنة As-Built vs Design
      if (e.target.closest("#gdRunAsBuiltBtn")) {
        computeAsBuiltComparison(20);
        refreshAllViews();
        if (typeof toast === "function") {
          toast(`تمت مقارنة ${asBuiltComparisons.length} نقطة وحساب الإزاحات بالمليمتر`, "success");
        }
        return;
      }

      if (e.target.closest("#gdClearDesignBtn")) {
        designPoints = [];
        asBuiltComparisons = [];
        renderStation();
        return;
      }

      // إضافة نقطة / تراجع / تفريغ
      if (e.target.closest("#gdAddPointRowBtn")) {
        pushUndo();
        const nextP = String(points.length + 1);
        const last = points[points.length - 1] || { e: 312450, n: 3110200, z: 24.5 };
        points.push({ p: nextP, e: Number((last.e + 5).toFixed(2)), n: Number((last.n + 5).toFixed(2)), z: last.z, code: "EG", cat: "EG" });
        refreshAllViews();
        return;
      }

      if (e.target.closest("#gdUndoExcelBtn")) {
        if (undoStack.length) {
          points = JSON.parse(undoStack.pop());
          refreshAllViews();
        }
        return;
      }

      if (e.target.closest("#gdClearExcelBtn")) {
        pushUndo();
        points = [];
        asBuiltComparisons = [];
        refreshAllViews();
        return;
      }

      // حذف صف
      const delBtn = e.target.closest("[data-del-row]");
      if (delBtn) {
        pushUndo();
        points.splice(Number(delBtn.dataset.delRow), 1);
        refreshAllViews();
        return;
      }

      // فلتر الفئات
      const catBtn = e.target.closest("[data-cat-filter]");
      if (catBtn) {
        activeFilter = catBtn.dataset.catFilter;
        refreshAllViews();
        return;
      }

      // أزرار معاينة الكاد
      if (e.target.closest("#gdToggleLabelsBtn")) {
        showLabels = !showLabels;
        e.target.closest("#gdToggleLabelsBtn").classList.toggle("active", showLabels);
        drawCadCanvas();
        return;
      }
      if (e.target.closest("#gdToggleConnectBtn")) {
        connectByCode = !connectByCode;
        e.target.closest("#gdToggleConnectBtn").classList.toggle("active", connectByCode);
        drawCadCanvas();
        return;
      }
      if (e.target.closest("#gdCadZoomInBtn")) {
        cadZoom = Math.min(8, cadZoom * 1.3);
        drawCadCanvas();
        return;
      }
      if (e.target.closest("#gdCadZoomOutBtn")) {
        cadZoom = Math.max(0.25, cadZoom / 1.3);
        drawCadCanvas();
        return;
      }
      if (e.target.closest("#gdCadZoomResetBtn")) {
        cadZoom = 1.0; cadPanX = 0; cadPanY = 0;
        drawCadCanvas();
        return;
      }

      // التصدير
      const expBtn = e.target.closest("[data-export-fmt]");
      if (expBtn) {
        handleExportFormat(expBtn.dataset.exportFmt);
      }
    });

    // تعديل القيم داخل الجدول مباشرة
    document.addEventListener("input", e => {
      const cell = e.target.closest(".gd-ex-cell");
      if (!cell) return;
      const idx = Number(cell.dataset.row);
      const field = cell.dataset.field;
      if (!points[idx]) return;

      if (field === "e" || field === "n" || field === "z") {
        points[idx][field] = parseFloat(cell.value) || 0;
      } else {
        points[idx][field] = cell.value;
      }
      if (designPoints.length) computeAsBuiltComparison(20);
      renderFiltersAndStats();
      drawCadCanvas();
    });

    // قراءة الملفات المرفوعة
    document.addEventListener("change", async e => {
      if (e.target.id === "gdImportSurveyFile" && e.target.files[0]) {
        const text = await e.target.files[0].text();
        const parsed = parseCoordinateText(text);
        if (parsed.length) {
          pushUndo();
          points = parsed;
          if (designPoints.length) computeAsBuiltComparison(20);
          refreshAllViews();
          if (typeof toast === "function") toast(`تم استيراد ${parsed.length} نقطة بنجاح`, "success");
        }
        e.target.value = "";
      }

      if (e.target.id === "gdImportDesignFile" && e.target.files[0]) {
        const text = await e.target.files[0].text();
        const parsed = parseCoordinateText(text);
        if (parsed.length) {
          designPoints = parsed;
          computeAsBuiltComparison(20);
          renderStation();
          if (typeof toast === "function") toast(`تم تحميل ${parsed.length} نقطة تصميمية وحساب التفاوتات`, "success");
        }
        e.target.value = "";
      }
    });

    // سحب لوحة الكاد بالماوس واللمس
    document.addEventListener("mousedown", e => {
      if (e.target.id === "gdCadCanvas") {
        isDraggingCad = true;
        dragStartX = e.clientX - cadPanX;
        dragStartY = e.clientY - cadPanY;
      }
    });
    window.addEventListener("mousemove", e => {
      if (!isDraggingCad) return;
      cadPanX = e.clientX - dragStartX;
      cadPanY = e.clientY - dragStartY;
      drawCadCanvas();
    });
    window.addEventListener("mouseup", () => { isDraggingCad = false; });

    document.addEventListener("touchstart", e => {
      if (e.target.id === "gdCadCanvas" && e.touches.length === 1) {
        isDraggingCad = true;
        dragStartX = e.touches[0].clientX - cadPanX;
        dragStartY = e.touches[0].clientY - cadPanY;
      }
    }, { passive: true });
    window.addEventListener("touchmove", e => {
      if (!isDraggingCad || e.touches.length !== 1) return;
      cadPanX = e.touches[0].clientX - dragStartX;
      cadPanY = e.touches[0].clientY - dragStartY;
      drawCadCanvas();
    }, { passive: true });
    window.addEventListener("touchend", () => { isDraggingCad = false; });
  }

  function init() {
    injectStyles();
    renderStation();
    bindEvents();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
