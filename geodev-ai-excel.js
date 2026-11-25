/* ============================================================================
   GeoDev Survey & GIS Excel Workstation (V6.3 ENTERPRISE - Elevation Hypsometric Map + As-Built & Report)
   - إصلاح فصل نقاط الأرض الطبيعية (EG) عن نقاط المطابق (MNH) حتى لو كان عمود Type يحتوي على SAN
   - تصدير Civil 3D بدون BOM (PENZD / PENZ / PNEZD / PEN)
   - مركز تصدير متكامل + 3 سمات + دعم AutoCAD DXF/DWG و Sokkia STR/SDR و ArcGIS Pro
   - خريطة المناسيب الملونة بالتدريج (أزرق غامق للأقل -> بني غامق للأعلى) + مفتاح ألوان بجانب الرسمة
   - مقارنة نقاط التصميم (Design) بالرفع التنفيذي الجديد (As-Built) وحساب الإزاحات بالمليمتر
   - تصدير التقرير المساحي الرسمي (PDF / Excel) ببيانات المشروع ونوع الأعمال وتاريخ الرفع واسم المساح وحد السماحية
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevAIExcelV63Init) return;
  window.GeoDevAIExcelV63Init = true;

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
    { P: "1", E: 312450.120, N: 3110210.450, Z: 24.510, Code: "BM" },
    { P: "2", E: 312475.840, N: 3110225.180, Z: 24.680, Code: "MNH" },
    { P: "3", E: 312510.300, N: 3110205.900, Z: 24.820, Code: "MNH" },
    { P: "4", E: 312495.600, N: 3110175.400, Z: 24.390, Code: "LP" },
    { P: "5", E: 312480.150, N: 3110190.250, Z: 24.420, Code: "TREE" },
    { P: "6", E: 312455.200, N: 3110180.100, Z: 24.450, Code: "EG" }
  ];

  // بيانات مقارنة التصميم بالرفع الجديد (As-Built vs Design)
  let designRows = [];
  let asBuiltResults = [];
  let currentToleranceMm = 20;

  let historyStack = [];
  let activeSubTab = "grid";
  let activeCodeFilter = "ALL"; // ALL | SEWER | LIGHT | TREES | NGL | ROAD | DIR
  let currentTheme = localStorage.getItem("gd_excel_theme_v6") || "dark_gold";
  let crsMode = "UTM36N";
  let showLabelsOnCanvas = true;
  let elevationColorMode = true; // تفعيل تلوين المناسيب بالتدريج (أزرق غامق للأقل -> بني غامق للأعلى)
  let connectMode = "BY_CODE";
  let searchFilter = "";
  let isLoggedInUser = false;

  let exportColOrder = "PENZD";
  let exportTargetScope = "ALL";

  let cadZoom = 1;
  let cadPanX = 0;
  let cadPanY = 0;
  let isDraggingCanvas = false;
  let dragStartX = 0;
  let dragStartY = 0;
  let initialPinchDist = 0;
  let initialPinchZoom = 1;

  function checkAuthState() {
    try {
      if (window.firebase && firebase.apps && firebase.apps.length && firebase.auth().currentUser) {
        return true;
      }
    } catch (_) {}
    return isLoggedInUser;
  }

  /* ============================================================================
     دالة التدرج اللوني للمناسيب (أقل منسوب = أزرق غامق -> أعلى منسوب = بني غامق)
     ============================================================================ */
  const ELEVATION_STOPS = [
    { t: 0.00, r: 15,  g: 43,  b: 107 }, // أزرق غامق جداً (أقل منسوب) #0f2b6b
    { t: 0.22, r: 29,  g: 112, b: 184 }, // أزرق متوسط
    { t: 0.42, r: 34,  g: 168, b: 110 }, // أخضر مائل للتركواز
    { t: 0.62, r: 234, g: 196, b: 53  }, // أصفر رملي
    { t: 0.82, r: 180, g: 83,  b: 9   }, // برتقالي محروق / بني فاتح
    { t: 1.00, r: 74,  g: 29,  b: 6   }  // بني غامق (أعلى منسوب) #4a1d06
  ];

  function getElevationColor(zVal, minZ, maxZ) {
    if (isNaN(zVal) || maxZ <= minZ) return "rgb(29, 112, 184)";
    const t = Math.max(0, Math.min(1, (zVal - minZ) / (maxZ - minZ)));

    for (let i = 0; i < ELEVATION_STOPS.length - 1; i++) {
      const s0 = ELEVATION_STOPS[i];
      const s1 = ELEVATION_STOPS[i + 1];
      if (t >= s0.t && t <= s1.t) {
        const localT = (t - s0.t) / Math.max(1e-6, s1.t - s0.t);
        const r = Math.round(s0.r + (s1.r - s0.r) * localT);
        const g = Math.round(s0.g + (s1.g - s0.g) * localT);
        const b = Math.round(s0.b + (s1.b - s0.b) * localT);
        return `rgb(${r}, ${g}, ${b})`;
      }
    }
    return "rgb(74, 29, 6)";
  }

  /* ============================================================================
     محرك التصنيف الدقيق: يعتمد أولاً على الكود الأساسي للنقطة (CODE) لمنع تداخل EG مع MNH
     ============================================================================ */
  function classifySurveyCode(rawCode, extraText = "") {
    const primary = String(rawCode || "").trim().toUpperCase();
    const extra = String(extraText || "").trim().toUpperCase();

    if (primary && primary !== "----" && primary !== "-") {
      if (/^(EG|NGL|GROUND|SPOT|BM|TBM|CP|GPS)$/.test(primary) || /^(EG|NGL)\b/.test(primary)) {
        const isBm = /^(BM|TBM|CP)$/.test(primary);
        return {
          group: "NGL",
          labelAr: isBm ? "نقطة روبير (BM)" : "أرض طبيعية (EG)",
          layer: isBm ? "GD_BENCHMARKS" : "GD_NGL_EG",
          color: "#34d399",
          dxfColor: 3,
          symbol: isBm ? "TRIANGLE" : "DOT"
        };
      }

      if (/^(MNH|MH|SAN|SEW|SEWER|IC|CB|GR|GULLY|PIT|MANHOLE|PIPE)$/.test(primary) || /\b(MNH|MH|MANHOLE|IC|CB)\b/.test(primary)) {
        return {
          group: "SEWER",
          labelAr: "مطبق / صرف صحي (MNH)",
          layer: "GD_SEWER_MNH",
          color: "#fb923c",
          dxfColor: 30,
          symbol: "MANHOLE"
        };
      }

      if (/^(LP|LIGHT|POLE|LAMP|EL|ELEC|TR|CAB)$/.test(primary) || /\b(LP|LIGHT|POLE)\b/.test(primary)) {
        return {
          group: "LIGHT",
          labelAr: "عمود إنارة (LP)",
          layer: "GD_LIGHT_POLES",
          color: "#facc15",
          dxfColor: 2,
          symbol: "LIGHT"
        };
      }

      if (/^(TREE|TR|PALM|SHRUB|GRN|GREEN|PLANT|VEG)$/.test(primary) || /\b(TREE|PALM)\b/.test(primary)) {
        return {
          group: "TREES",
          labelAr: "أشجار / نخيل (TREE)",
          layer: "GD_TREES",
          color: "#22c55e",
          dxfColor: 3,
          symbol: "TREE"
        };
      }

      if (/^(RD|ROAD|CL|EP|EDGE|CURB|KERB|SW|ASPH|PAV)$/.test(primary) || /\b(RD|ROAD|CURB|ASPH)\b/.test(primary)) {
        return {
          group: "ROAD",
          labelAr: "طريق / رصيف (RD)",
          layer: "GD_ROADS",
          color: "#60a5fa",
          dxfColor: 5,
          symbol: "DOT"
        };
      }

      if (/^(N|S|NORTH|SOUTH|NE|NW|SE|SW)$/.test(primary)) {
        const isSouth = /^(S|SOUTH|SE|SW)$/.test(primary);
        return {
          group: "DIR",
          labelAr: isSouth ? "اتجاه جنوب (South)" : "اتجاه شمال (North)",
          layer: isSouth ? "GD_DIR_SOUTH" : "GD_DIR_NORTH",
          color: isSouth ? "#f43f5e" : "#38bdf8",
          dxfColor: isSouth ? 1 : 4,
          symbol: "DIR"
        };
      }

      return {
        group: "OTHER",
        labelAr: `كود (${rawCode})`,
        layer: `GD_${String(rawCode).replace(/[^a-zA-Z0-9_]/g, "_")}`,
        color: "#e09f24",
        dxfColor: 7,
        symbol: "DOT"
      };
    }

    if (/\b(MNH|MH|MANHOLE|IC|CB)\b/.test(extra)) {
      return { group: "SEWER", labelAr: "مطبق / صرف صحي (MNH)", layer: "GD_SEWER_MNH", color: "#fb923c", dxfColor: 30, symbol: "MANHOLE" };
    }
    if (/\b(LP|LIGHT|POLE)\b/.test(extra)) {
      return { group: "LIGHT", labelAr: "عمود إنارة (LP)", layer: "GD_LIGHT_POLES", color: "#facc15", dxfColor: 2, symbol: "LIGHT" };
    }
    if (/\b(TREE|PALM)\b/.test(extra)) {
      return { group: "TREES", labelAr: "أشجار / نخيل (TREE)", layer: "GD_TREES", color: "#22c55e", dxfColor: 3, symbol: "TREE" };
    }

    return { group: "OTHER", labelAr: "نقطة عامة", layer: "GD_GENERAL", color: "#94a3b8", dxfColor: 7, symbol: "DOT" };
  }

  function getRowClassification(rowObj) {
    const { cKey } = getCoordKeys();
    const codeVal = cKey ? rowObj[cKey] : "";
    const extraVals = Object.entries(rowObj)
      .filter(([k]) => !/^(p|e|n|z|east|north|norht|level|id|no|code)$/i.test(k.trim()))
      .map(([, v]) => v)
      .join(" ");
    return classifySurveyCode(codeVal, extraVals);
  }

  /* ============================================================================
     محرك مقارنة النقاط التصميمية (القديمة) بنقاط الرفع الجديد (As-Built vs Design)
     ============================================================================ */
  function computeAsBuiltComparison(tolMm = currentToleranceMm) {
    currentToleranceMm = Number(tolMm) || 20;
    if (!designRows.length || !rows.length) {
      asBuiltResults = [];
      return [];
    }

    const { pKey, eKey, nKey, zKey, cKey } = getCoordKeys();
    const results = [];

    rows.forEach((actRow, idx) => {
      const actP = String(actRow[pKey] || idx + 1).trim();
      const actE = parseFloat(actRow[eKey]);
      const actN = parseFloat(actRow[nKey]);
      const actZ = parseFloat(actRow[zKey]) || 0;
      const actC = String(actRow[cKey] || "").trim();
      if (isNaN(actE) || isNaN(actN)) return;

      let matchedDesign = designRows.find(d => String(d.P || "").trim().toLowerCase() === actP.toLowerCase());

      if (!matchedDesign) {
        let minDist = Infinity;
        designRows.forEach(d => {
          const dist = Math.hypot(actE - d.E, actN - d.N);
          if (dist < minDist) {
            minDist = dist;
            matchedDesign = d;
          }
        });
      }

      if (matchedDesign) {
        const dE_m = actE - matchedDesign.E;
        const dN_m = actN - matchedDesign.N;
        const dZ_m = actZ - (matchedDesign.Z || 0);
        const dXY_m = Math.hypot(dE_m, dN_m);

        const dE_mm = dE_m * 1000;
        const dN_mm = dN_m * 1000;
        const dZ_mm = dZ_m * 1000;
        const dXY_mm = dXY_m * 1000;

        let az = (Math.atan2(dE_m, dN_m) * 180) / Math.PI;
        if (az < 0) az += 360;

        const isPass = dXY_mm <= currentToleranceMm;
        results.push({
          rowIdx: idx,
          P: actP,
          Code: actC || matchedDesign.Code || "PT",
          desE: matchedDesign.E,
          desN: matchedDesign.N,
          desZ: matchedDesign.Z || 0,
          actE,
          actN,
          actZ,
          dE_mm,
          dN_mm,
          dXY_mm,
          dZ_mm,
          azimuth: az,
          isPass,
          statusText: isPass ? "مقبول (داخل السماحية)" : "تجاوز السماحية"
        });
      }
    });

    asBuiltResults = results;
    return results;
  }

  function resetCadView() {
    cadZoom = 1;
    cadPanX = 0;
    cadPanY = 0;
    drawCadPreview();
  }

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
    if (designRows.length) computeAsBuiltComparison();
    renderExcelTable();
    drawCadPreview();
    showExcelLog("تم التراجع عن آخر تعديل بنجاح.", "success");
  }

  function showExcelLog(msg, type = "info") {
    const box = document.getElementById("gdExcelAiLog");
    if (!box) return;
    const palette = {
      success: { bg: "rgba(22,163,74,0.16)", border: "#22c55e", text: "#16a34a" },
      warn: { bg: "rgba(234,179,8,0.16)", border: "#eab308", text: "#ca8a04" },
      error: { bg: "rgba(220,38,38,0.16)", border: "#ef4444", text: "#dc2626" },
      info: { bg: "rgba(2,132,199,0.14)", border: "#0284c7", text: "#0284c7" }
    };
    const isDark = currentTheme !== "light_eng";
    const darkText = { success: "#4ade80", warn: "#facc15", error: "#f87171", info: "#38bdf8" };
    const c = palette[type] || palette.info;
    box.style.background = c.bg;
    box.style.borderColor = c.border;
    box.style.color = isDark ? darkText[type] : c.text;
    box.textContent = msg;
  }

  function openAuthRequiredModal(actionLabel = "حفظ وتصدير الملفات واستخدام أوامر الذكاء الاصطناعي") {
    const oldModal = document.getElementById("gdExcelAuthGateModal");
    if (oldModal) oldModal.remove();

    const overlay = document.createElement("div");
    overlay.id = "gdExcelAuthGateModal";
    overlay.style.cssText = `
      position: fixed; inset: 0; z-index: 999999;
      background: rgba(10, 9, 8, 0.86); backdrop-filter: blur(6px);
      display: flex; align-items: center; justify-content: center;
      padding: 16px; direction: rtl; font-family: inherit;
    `;

    overlay.innerHTML = `
      <div style="width:100%;max-width:430px;background:#1c1917;border:2px solid #e09f24;border-radius:18px;padding:22px;color:#f5f5f4;text-align:center;box-shadow:0 24px 60px rgba(0,0,0,0.75);">
        <div style="width:52px;height:52px;margin:0 auto 12px;border-radius:14px;background:rgba(224,159,36,0.16);border:1.5px solid #e09f24;display:grid;place-items:center;color:#fbbf24;font-weight:900;font-size:20px;">
          ID
        </div>
        <h3 style="margin:0 0 8px;font-size:17.5px;font-weight:900;color:#fbbf24;">
          تسجيل الدخول أو إنشاء حساب للمتابعة
        </h3>
        <p style="margin:0 0 16px;font-size:13px;color:#d6d3d1;line-height:1.75;">
          تمت معاينة بياناتك بنجاح. لكي تتمكن من <b>${actionLabel}</b> على منصة <b>GeoDev</b>، يرجى إنشاء حساب مجاني أو تسجيل الدخول:
        </p>
        <div style="display:flex;flex-direction:column;gap:9px;">
          <button type="button" id="gdGoToLoginFromExcelBtn"
            style="width:100%;padding:13px;border-radius:11px;border:0;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-size:14px;font-weight:900;cursor:pointer;font-family:inherit;">
            إنشاء حساب مجاني / تسجيل الدخول
          </button>
          <button type="button" id="gdStayGuestPreviewBtn"
            style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#292524;color:#a8a29e;font-size:12.5px;font-weight:800;cursor:pointer;font-family:inherit;">
            متابعة المعاينة كزائر
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);
    const closeModal = () => overlay.remove();
    document.getElementById("gdStayGuestPreviewBtn").onclick = closeModal;
    overlay.addEventListener("click", e => { if (e.target === overlay) closeModal(); });

    document.getElementById("gdGoToLoginFromExcelBtn").onclick = () => {
      closeModal();
      const authContainer =
        document.getElementById("authScreen") ||
        document.getElementById("loginScreen") ||
        document.getElementById("authBox") ||
        document.querySelector(".auth-card, .login-card, #s-account, form");
      if (authContainer) {
        authContainer.scrollIntoView({ behavior: "smooth", block: "center" });
        const emailInput = authContainer.querySelector('input[type="email"], input[type="text"]');
        if (emailInput) setTimeout(() => emailInput.focus(), 400);
      } else {
        window.scrollTo({ top: 0, behavior: "smooth" });
      }
    };
  }

  function injectStyles() {
    const old = document.getElementById("gdAiExcelStyles");
    if (old) old.remove();

    const style = document.createElement("style");
    style.id = "gdAiExcelStyles";
    style.textContent = `
      #gdAiExcelStudio[data-theme="dark_gold"] {
        --ex-bg: linear-gradient(165deg, #141210, #1e1b18);
        --ex-panel: #100e0c;
        --ex-card: #211e1b;
        --ex-border: #e09f24;
        --ex-line: #3f3a36;
        --ex-text: #f5f5f4;
        --ex-muted: #a8a29e;
        --ex-accent: #fbbf24;
        --ex-btn-bg: #292524;
        --ex-th-bg: #292524;
        --ex-td-bg: #141210;
        --ex-td-alt: #191614;
      }
      #gdAiExcelStudio[data-theme="light_eng"] {
        --ex-bg: linear-gradient(165deg, #f8fafc, #f1f5f9);
        --ex-panel: #ffffff;
        --ex-card: #ffffff;
        --ex-border: #0284c7;
        --ex-line: #cbd5e1;
        --ex-text: #0f172a;
        --ex-muted: #475569;
        --ex-accent: #0369a1;
        --ex-btn-bg: #e2e8f0;
        --ex-th-bg: #e2e8f0;
        --ex-td-bg: #ffffff;
        --ex-td-alt: #f8fafc;
      }
      #gdAiExcelStudio[data-theme="cad_blue"] {
        --ex-bg: linear-gradient(165deg, #09111e, #0f1c33);
        --ex-panel: #070d18;
        --ex-card: #13223d;
        --ex-border: #38bdf8;
        --ex-line: #1e3a5f;
        --ex-text: #f1f5f9;
        --ex-muted: #94a3b8;
        --ex-accent: #38bdf8;
        --ex-btn-bg: #162847;
        --ex-th-bg: #162847;
        --ex-td-bg: #0b1424;
        --ex-td-alt: #0f1b30;
      }

      #gdGuestWelcomeBanner {
        direction: rtl;
        text-align: right;
        background: linear-gradient(135deg, #1c1917, #292524);
        border: 1.5px solid #e09f24;
        border-radius: 14px;
        padding: 13px 16px;
        margin: 12px auto;
        max-width: 980px;
        color: #f5f5f4;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 10px;
        box-shadow: 0 10px 28px rgba(0, 0, 0, 0.35);
      }
      #gdGuestWelcomeBanner strong { color: #fbbf24; font-size: 14px; display: block; }
      #gdGuestWelcomeBanner span { font-size: 12px; color: #d6d3d1; }

      #gdAiExcelStudio {
        direction: rtl;
        text-align: right;
        background: var(--ex-bg);
        border: 2px solid var(--ex-border);
        border-radius: 18px;
        padding: 16px;
        margin: 14px auto 18px;
        max-width: 980px;
        color: var(--ex-text);
        box-shadow: 0 16px 40px rgba(0, 0, 0, 0.35);
      }
      .gd-ex-top-bar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 10px;
        padding-bottom: 12px;
        border-bottom: 1px solid var(--ex-line);
        margin-bottom: 14px;
      }
      .gd-ex-top-bar h3 {
        margin: 0;
        font-size: 17.5px;
        font-weight: 900;
        color: var(--ex-accent);
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .gd-ex-top-bar p { margin: 4px 0 0; font-size: 12px; color: var(--ex-muted); }

      .gd-ex-header-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
      .gd-theme-bulb-btn {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 12px;
        border-radius: 999px;
        border: 1.5px solid var(--ex-border);
        background: var(--ex-card);
        color: var(--ex-accent);
        font-size: 11.5px;
        font-weight: 900;
        cursor: pointer;
        font-family: inherit;
      }
      .gd-report-launch-btn {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 7px 13px;
        border-radius: 999px;
        border: 1.5px solid #4ade80;
        background: linear-gradient(135deg, #16a34a, #15803d);
        color: #ffffff;
        font-size: 12px;
        font-weight: 900;
        cursor: pointer;
        font-family: inherit;
      }
      .gd-ex-pro-pill {
        background: linear-gradient(135deg, #e09f24, #b45309);
        color: #141210;
        font-weight: 900;
        font-size: 11px;
        padding: 6px 12px;
        border-radius: 999px;
      }

      .gd-ex-ai-panel {
        background: var(--ex-panel);
        border: 1.5px solid var(--ex-border);
        border-radius: 14px;
        padding: 13px;
        margin-bottom: 14px;
      }
      .gd-ex-ai-row { display: flex; gap: 8px; flex-wrap: wrap; }
      .gd-ex-input {
        flex: 1;
        min-width: 200px;
        padding: 10px 12px;
        border-radius: 10px;
        border: 1.5px solid var(--ex-line);
        background: var(--ex-card);
        color: var(--ex-text);
        font-size: 13px;
        font-weight: 700;
        font-family: inherit;
      }
      .gd-ex-input:focus { outline: none; border-color: var(--ex-accent); }
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
      .gd-ex-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 8px; }
      .gd-ex-chip {
        background: var(--ex-btn-bg);
        border: 1px solid var(--ex-line);
        color: var(--ex-text);
        padding: 5px 10px;
        border-radius: 8px;
        font-size: 11px;
        font-weight: 800;
        cursor: pointer;
        font-family: inherit;
      }
      .gd-ex-chip:hover { border-color: var(--ex-accent); color: var(--ex-accent); }
      #gdExcelAiLog {
        margin-top: 9px;
        padding: 8px 12px;
        border-radius: 9px;
        border: 1px solid #0284c7;
        background: rgba(2, 132, 199, 0.12);
        color: #38bdf8;
        font-size: 12px;
        font-weight: 800;
        line-height: 1.6;
      }

      .gd-ex-code-filters {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin-bottom: 12px;
        padding: 9px;
        background: var(--ex-panel);
        border: 1px solid var(--ex-line);
        border-radius: 12px;
      }
      .gd-ex-filter-pill {
        padding: 6px 11px;
        border-radius: 8px;
        border: 1px solid var(--ex-line);
        background: var(--ex-card);
        color: var(--ex-text);
        font-size: 11.5px;
        font-weight: 800;
        cursor: pointer;
        font-family: inherit;
      }
      .gd-ex-filter-pill.active {
        background: rgba(224, 159, 36, 0.22);
        border-color: var(--ex-border);
        color: var(--ex-accent);
      }

      .gd-ex-mode-tabs {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 7px;
        margin-bottom: 13px;
      }
      @media (max-width: 640px) {
        .gd-ex-mode-tabs { grid-template-columns: repeat(2, 1fr); }
      }
      .gd-ex-mode-tab {
        padding: 10px 8px;
        border-radius: 10px;
        border: 1.5px solid var(--ex-line);
        background: var(--ex-card);
        color: var(--ex-muted);
        font-size: 12px;
        font-weight: 800;
        cursor: pointer;
        font-family: inherit;
        text-align: center;
      }
      .gd-ex-mode-tab.active {
        background: rgba(224, 159, 36, 0.18);
        border-color: var(--ex-border);
        color: var(--ex-accent);
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
        border-radius: 9px;
        border: 1px solid var(--ex-line);
        background: var(--ex-btn-bg);
        color: var(--ex-text);
        font-size: 11.5px;
        font-weight: 800;
        cursor: pointer;
        font-family: inherit;
        display: inline-flex;
        align-items: center;
        gap: 5px;
      }
      .gd-ex-btn:hover { border-color: var(--ex-accent); color: var(--ex-accent); }
      .gd-ex-btn.gold {
        background: linear-gradient(135deg, #e09f24, #b45309);
        color: #141210;
        border: 0;
        font-weight: 900;
      }
      .gd-ex-btn.danger {
        background: rgba(185, 28, 28, 0.25);
        border-color: #ef4444;
        color: #ef4444;
      }

      .gd-ex-stats {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(110px, 1fr));
        gap: 7px;
        margin-bottom: 12px;
      }
      .gd-ex-stat {
        background: var(--ex-panel);
        border: 1px solid var(--ex-line);
        border-radius: 10px;
        padding: 8px;
        text-align: center;
      }
      .gd-ex-stat span { display: block; font-size: 10px; color: var(--ex-muted); }
      .gd-ex-stat strong { font-size: 13px; color: var(--ex-accent); font-weight: 900; }

      .gd-ex-grid-wrap {
        max-height: 290px;
        overflow: auto;
        border: 1.5px solid var(--ex-line);
        border-radius: 12px;
        background: var(--ex-panel);
        margin-bottom: 14px;
      }
      .gd-ex-table { width: 100%; border-collapse: collapse; font-size: 12.5px; text-align: center; }
      .gd-ex-table th {
        position: sticky;
        top: 0;
        z-index: 5;
        background: var(--ex-th-bg);
        color: var(--ex-accent);
        padding: 9px 6px;
        border: 1px solid var(--ex-line);
        font-weight: 900;
        white-space: nowrap;
      }
      .gd-ex-table td { border: 1px solid var(--ex-line); padding: 0; background: var(--ex-td-bg); }
      .gd-ex-table tr:nth-child(even) td { background: var(--ex-td-alt); }
      .gd-ex-cell {
        width: 100%;
        min-width: 76px;
        box-sizing: border-box;
        border: 0;
        background: transparent;
        color: var(--ex-text);
        padding: 8px 6px;
        text-align: center;
        font-size: 12.5px;
        font-weight: 700;
        font-family: inherit;
      }
      .gd-ex-cell:focus { outline: none; background: rgba(224, 159, 36, 0.16); }

      .gd-ex-subpanel {
        background: var(--ex-panel);
        border: 1px solid var(--ex-line);
        border-radius: 13px;
        padding: 13px;
        margin-bottom: 14px;
      }
      .gd-ex-textarea {
        width: 100%;
        min-height: 110px;
        box-sizing: border-box;
        background: var(--ex-card);
        border: 1.5px solid var(--ex-line);
        border-radius: 10px;
        color: var(--ex-text);
        padding: 10px;
        font-family: monospace;
        font-size: 12px;
        margin: 8px 0;
        direction: ltr;
      }

      /* لوحة الرسم + مفتاح تدرج ألوان المناسيب بجانبها */
      .gd-ex-canvas-box {
        background: #06080d;
        border: 1.5px solid var(--ex-border);
        border-radius: 14px;
        padding: 10px;
        margin-bottom: 16px;
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
      .gd-cad-stage-layout {
        display: flex;
        gap: 8px;
        align-items: stretch;
        height: 320px;
      }
      .gd-ex-canvas-wrapper {
        position: relative;
        flex: 1;
        min-width: 0;
        height: 100%;
        border-radius: 10px;
        overflow: hidden;
        touch-action: none;
        cursor: grab;
        border: 1px solid #1e293b;
      }
      .gd-ex-canvas-wrapper:active { cursor: grabbing; }
      #gdExcelCadCanvas { width: 100%; height: 100%; display: block; background: #04060a; }

      /* مفتاح ألوان المناسيب المتدرج بجانب الرسمة */
      .gd-elev-legend-panel {
        width: 96px;
        flex-shrink: 0;
        background: #0b111e;
        border: 1px solid #1e293b;
        border-radius: 10px;
        padding: 8px 6px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        color: #f1f5f9;
        font-size: 10px;
      }
      .gd-elev-legend-title {
        text-align: center;
        font-weight: 900;
        color: #fbbf24;
        font-size: 10.5px;
        border-bottom: 1px solid #1e293b;
        padding-bottom: 4px;
        margin-bottom: 5px;
      }
      .gd-elev-legend-body {
        flex: 1;
        display: flex;
        align-items: stretch;
        gap: 6px;
        min-height: 0;
      }
      .gd-elev-gradient-bar {
        width: 15px;
        border-radius: 6px;
        border: 1px solid rgba(255, 255, 255, 0.25);
        /* من الأعلى (أعلى منسوب = بني غامق) إلى الأسفل (أقل منسوب = أزرق غامق) */
        background: linear-gradient(
          to bottom,
          #4a1d06 0%,
          #b45309 20%,
          #eac435 42%,
          #22a86e 64%,
          #1d70b8 82%,
          #0f2b6b 100%
        );
      }
      .gd-elev-ticks {
        flex: 1;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        font-family: monospace;
        font-size: 10px;
        font-weight: 800;
        direction: ltr;
        text-align: left;
      }
      .gd-elev-tick-max { color: #f59e0b; }
      .gd-elev-tick-mid { color: #cbd5e1; }
      .gd-elev-tick-min { color: #38bdf8; }
      .gd-elev-delta-badge {
        margin-top: 5px;
        padding-top: 4px;
        border-top: 1px solid #1e293b;
        text-align: center;
        font-size: 9.5px;
        color: #34d399;
        font-weight: 900;
        direction: ltr;
      }

      .gd-cad-zoom-controls {
        position: absolute;
        bottom: 10px;
        left: 10px;
        display: flex;
        gap: 5px;
        z-index: 10;
      }
      .gd-cad-zbtn {
        background: rgba(15, 23, 42, 0.9);
        border: 1px solid #38bdf8;
        color: #f8fafc;
        padding: 6px 10px;
        border-radius: 8px;
        font-size: 12px;
        font-weight: 900;
        cursor: pointer;
        font-family: inherit;
      }

      .gd-export-hub {
        background: var(--ex-panel);
        border: 1.5px solid var(--ex-border);
        border-radius: 15px;
        padding: 14px;
      }
      .gd-export-hub-title {
        font-size: 14.5px;
        font-weight: 900;
        color: var(--ex-accent);
        margin: 0 0 10px;
        padding-bottom: 8px;
        border-bottom: 1px solid var(--ex-line);
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 8px;
      }
      .gd-export-selectors {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 10px;
        margin-bottom: 12px;
      }
      .gd-export-select-box label {
        display: block;
        font-size: 11.5px;
        font-weight: 900;
        color: var(--ex-muted);
        margin-bottom: 5px;
      }
      .gd-export-select-box select {
        width: 100%;
        padding: 10px 12px;
        border-radius: 10px;
        border: 1.5px solid var(--ex-border);
        background: var(--ex-card);
        color: var(--ex-text);
        font-size: 13px;
        font-weight: 800;
        font-family: inherit;
        cursor: pointer;
      }
      .gd-ex-export-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
        gap: 8px;
      }
      .gd-ex-exp-btn {
        padding: 10px 9px;
        border-radius: 10px;
        border: 1.5px solid var(--ex-line);
        background: var(--ex-card);
        color: var(--ex-text);
        font-size: 12px;
        font-weight: 800;
        cursor: pointer;
        font-family: inherit;
        text-align: center;
        transition: all 0.2s ease;
      }
      .gd-ex-exp-btn:hover { border-color: var(--ex-accent); transform: translateY(-1px); }
      .gd-ex-exp-btn small { display: block; font-size: 10px; color: var(--ex-muted); margin-top: 2px; }
    `;
    document.head.appendChild(style);
  }

  function getThemeLabel(themeKey) {
    if (themeKey === "light_eng") return "السمة: المكتب الفاتح";
    if (themeKey === "cad_blue") return "السمة: الكاد الأزرق";
    return "السمة: الذهبي الداكن";
  }

  function renderStudioHTML() {
    const logged = checkAuthState();

    return `
      <div class="gd-ex-top-bar">
        <div>
          <h3>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
              <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/>
            </svg>
            محطة إكسيل المساحة ونظم المعلومات الجغرافية (Survey &amp; GIS Excel)
          </h3>
          <p>استيراد وتصدير مباشر متوافق 100% مع Civil 3D (PENZD / PENZ) و AutoCAD و Sokkia SDR/STR و ArcGIS Pro</p>
        </div>

        <div class="gd-ex-header-actions">
          <button type="button" id="gdOpenOfficialReportBtn" class="gd-report-launch-btn" title="تصدير التقرير المساحي والهندسي PDF / Excel">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>
            </svg>
            <span>تصدير التقرير</span>
          </button>

          <button type="button" id="gdCycleThemeBulbBtn" class="gd-theme-bulb-btn" title="تغيير سمة الأداة (3 سمات)">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
              <path d="M9 18h6M10 22h4M12 2a7 7 0 0 0-4 12.7V17h8v-2.3A7 7 0 0 0 12 2z"/>
            </svg>
            <span id="gdThemeBulbLabel">${getThemeLabel(currentTheme)}</span>
          </button>
          <span class="gd-ex-pro-pill">${logged ? "ENTERPRISE V6.3" : "وضع المعاينة"}</span>
        </div>
      </div>

      <div class="gd-ex-ai-panel">
        <div class="gd-ex-ai-row">
          <input type="text" id="gdExcelAiPrompt" class="gd-ex-input"
            placeholder="أدخل أمراً هندسياً: صنف الأكواد تلقائياً | افحص أخطاء الرفع | احسب المساحة | اجمع 0.50 على Z...">
          <button type="button" id="gdExcelAiExecBtn" class="gd-ex-ai-btn">تنفيذ الأمر</button>
        </div>
        <div class="gd-ex-chips">
          <button type="button" class="gd-ex-chip" data-ai-cmd="صنف الأكواد المساحية تلقائياً">◈ تصنيف الأكواد (MNH / إنارة / EG)</button>
          <button type="button" class="gd-ex-chip" data-ai-cmd="افحص أخطاء الرفع المساحي">◈ فحص الجودة (QA/QC)</button>
          <button type="button" class="gd-ex-chip" data-ai-cmd="احسب المسافة والانحراف">◈ المسافة والانحراف</button>
          <button type="button" class="gd-ex-chip" data-ai-cmd="احسب المساحة والمحيط">◈ المساحة بالفدان والمتر</button>
          <button type="button" class="gd-ex-chip" data-ai-cmd="اجمع 0.50 على Z">◈ إزاحة المنسوب (+0.50 Z)</button>
          <button type="button" class="gd-ex-chip" data-ai-cmd="بدل الشرق والشمال">◈ تبديل E ⇄ N</button>
        </div>
        <div id="gdExcelAiLog">
          ${logged
            ? "النظام جاهز: تلوين المناسيب مفعّل (الأزرق الغامق لأقل منسوب والبني الغامق لأعلى منسوب) مع مفتاح التدرج."
            : "وضع المعاينة للزوار: يمكنك رفع الملفات ومعاينة خريطة المناسيب على لوحة الكاد مجاناً. سجّل الدخول لتفعيل الحفظ والتصدير."}
        </div>
      </div>

      <div class="gd-ex-toolbar">
        <label class="gd-ex-btn gold" style="cursor:pointer;">
          ⊞ استيراد ملف (Excel / DXF / SDR / STR / GSI / CSV)
          <input type="file" id="gdExcelFileInput" accept=".xlsx,.xls,.csv,.txt,.dxf,.sdr,.str,.cdr,.gsi" style="display:none;">
        </label>
        <button type="button" class="gd-ex-btn" id="gdExAutoClassifyBtn">◈ تصنيف الأكواد تلقائياً</button>
        <button type="button" class="gd-ex-btn" id="gdExAddRowBtn">+ إضافة نقطة</button>
        <button type="button" class="gd-ex-btn" id="gdExSwapENBtn">⇄ تبديل E / N</button>
        <button type="button" class="gd-ex-btn" id="gdExUndoBtn">↶ تراجع</button>
        <button type="button" class="gd-ex-btn danger" id="gdExClearBtn">✕ تفريغ الجدول</button>
      </div>

      <div class="gd-ex-code-filters" id="gdExCodeFilterBar"></div>

      <div class="gd-ex-mode-tabs">
        <button type="button" class="gd-ex-mode-tab active" data-ex-subtab="grid">▤ جدول الإكسيل والبيانات</button>
        <button type="button" class="gd-ex-mode-tab" data-ex-subtab="asbuilt_compare">⇄ مقارنة التصميم بالرفع (As-Built)</button>
        <button type="button" class="gd-ex-mode-tab" data-ex-subtab="cad_import">⬡ استخراج من AutoCAD LIST</button>
        <button type="button" class="gd-ex-mode-tab" data-ex-subtab="calc_tools">⌖ إزاحة وحفر وردم (Cut/Fill)</button>
      </div>

      <!-- تبويب مقارنة نقاط التصميم (القديمة) بنقاط الرفع الجديد (As-Built vs Design) -->
      <div id="gdExSubpanelAsBuilt" class="gd-ex-subpanel" style="display:none;border-color:#e09f24;">
        <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:8px;">
          <div>
            <strong style="color:var(--ex-accent);font-size:13.5px;display:block;">مقارنة نقاط التصميم / التوقيع الأولية بنقاط الرفع الجديد (As-Built vs Design)</strong>
            <span style="font-size:11.5px;color:var(--ex-muted);">ارفع ملف نقاط الأعمدة أو التصميم الأصلي لمقارنته بنقاط الرفع الموجودة في الجدول وحساب الإزاحة بالمليمتر</span>
          </div>
          <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;">
            <label class="gd-ex-btn gold" style="cursor:pointer;">
              ⊞ رفع نقاط التصميم (القديمة)
              <input type="file" id="gdDesignFileInput" accept=".xlsx,.xls,.csv,.txt,.dxf,.sdr,.str,.cdr,.gsi" style="display:none;">
            </label>
            <input type="number" step="any" id="gdAsBuiltTolInput" class="gd-ex-input" value="${currentToleranceMm}" style="width:95px;min-width:95px;text-align:center;padding:7px;" title="حد السماحية بالمليمتر">
            <button type="button" class="gd-ex-btn" id="gdRunAsBuiltCalcBtn">احسب التفاوتات</button>
            <button type="button" class="gd-ex-btn danger" id="gdClearDesignBtn">مسح التصميم</button>
          </div>
        </div>
        <div id="gdAsBuiltSummaryBox" style="font-size:12px;color:var(--ex-muted);padding:8px 10px;border-radius:8px;background:var(--ex-card);border:1px solid var(--ex-line);">
          ${designRows.length
            ? `تم تحميل ${designRows.length} نقطة تصميمية ومقارنتها بنقاط الرفع الجديد.`
            : "لم يتم رفع ملف نقاط التصميم بعد. ارفع ملف نقاط التصميم (Excel / CSV / DXF / SDR / STR) ثم اضغط «احسب التفاوتات»."}
        </div>
      </div>

      <div id="gdExSubpanelCadImport" class="gd-ex-subpanel" style="display:none;">
        <strong style="color:var(--ex-accent);font-size:13px;">استخراج الإحداثيات من الأوتوكاد عبر أمر LIST:</strong>
        <textarea id="gdAcadListPasteArea" class="gd-ex-textarea" placeholder="at point  X= 312450.120  Y= 3110210.450  Z= 24.510"></textarea>
        <button type="button" class="gd-ex-btn gold" id="gdParseAcadListBtn" style="width:100%;justify-content:center;">
          استخراج الإحداثيات وإضافتها للجدول ولوحة الرسم
        </button>
      </div>

      <div id="gdExSubpanelCalcTools" class="gd-ex-subpanel" style="display:none;">
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;">
          <div style="background:var(--ex-card);padding:10px;border-radius:10px;border:1px solid var(--ex-line);">
            <strong style="color:var(--ex-accent);font-size:12.5px;">1. إزاحة الإحداثيات (Coordinate Shift)</strong>
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin:8px 0;">
              <input type="number" step="any" id="gdShiftE" class="gd-ex-input" style="min-width:0;text-align:center;" placeholder="ΔE (0)">
              <input type="number" step="any" id="gdShiftN" class="gd-ex-input" style="min-width:0;text-align:center;" placeholder="ΔN (0)">
              <input type="number" step="any" id="gdShiftZ" class="gd-ex-input" style="min-width:0;text-align:center;" placeholder="ΔZ (0)">
            </div>
            <button type="button" class="gd-ex-btn" id="gdApplyShiftBtn" style="width:100%;justify-content:center;">تطبيق الإزاحة</button>
          </div>

          <div style="background:var(--ex-card);padding:10px;border-radius:10px;border:1px solid var(--ex-line);">
            <strong style="color:#16a34a;font-size:12.5px;">2. حساب الحفر والردم (Cut &amp; Fill)</strong>
            <div style="margin:8px 0;">
              <input type="number" step="any" id="gdDesignLevelInput" class="gd-ex-input" style="width:100%;text-align:center;" placeholder="منسوب التصميم (مثال: 24.50)">
            </div>
            <button type="button" class="gd-ex-btn gold" id="gdCalcCutFillBtn" style="width:100%;justify-content:center;">حساب عمود Cut / Fill</button>
          </div>
        </div>
      </div>

      <div class="gd-ex-stats" id="gdExStatsBar"></div>

      <div class="gd-ex-grid-wrap">
        <table class="gd-ex-table" id="gdExcelDataTable"></table>
      </div>

      <div class="gd-ex-canvas-box">
        <div class="gd-ex-canvas-head">
          <span>معاينة CAD وتدرج المناسيب (أقل منسوب: أزرق غامق ◄► أعلى منسوب: بني غامق)</span>
          <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;">
            <input type="text" id="gdExSearchInput" class="gd-ex-input" style="max-width:100px;padding:4px 8px;font-size:11px;" placeholder="بحث...">
            <button type="button" class="gd-ex-btn" style="padding:4px 8px;font-size:11px;${elevationColorMode ? "border-color:#fbbf24;color:#fbbf24;" : ""}" id="gdToggleElevColorBtn">
              تدرج المناسيب: ${elevationColorMode ? "مفعّل" : "أكواد"}
            </button>
            <button type="button" class="gd-ex-btn" style="padding:4px 8px;font-size:11px;" id="gdToggleLabelsBtn">النصوص</button>
            <button type="button" class="gd-ex-btn" style="padding:4px 8px;font-size:11px;" id="gdCycleConnectBtn">التوصيل: حسب الكود</button>
            <select id="gdCrsSelect" class="gd-ex-btn" style="padding:4px 8px;font-size:11px;">
              <option value="UTM36N" selected>UTM Zone 36N</option>
              <option value="UTM35N">UTM Zone 35N</option>
              <option value="WGS84">WGS84 (Lat/Lon)</option>
            </select>
          </div>
        </div>

        <div class="gd-cad-stage-layout">
          <div class="gd-ex-canvas-wrapper" id="gdCadCanvasWrapper">
            <canvas id="gdExcelCadCanvas"></canvas>
            <div class="gd-cad-zoom-controls">
              <button type="button" class="gd-cad-zbtn" id="gdCadZoomInBtn">+ تقريب</button>
              <button type="button" class="gd-cad-zbtn" id="gdCadZoomOutBtn">- تبعيد</button>
              <button type="button" class="gd-cad-zbtn" id="gdCadZoomResetBtn">توسيط الكل</button>
            </div>
          </div>

          <!-- مفتاح الألوان المتدرج بجانب الرسمة من الأقل للأكبر -->
          <div class="gd-elev-legend-panel" id="gdElevationLegendPanel">
            <div class="gd-elev-legend-title">مفتاح المناسيب</div>
            <div class="gd-elev-legend-body">
              <div class="gd-elev-gradient-bar" title="تدرج الألوان من أقل منسوب (أزرق غامق) إلى أعلى منسوب (بني غامق)"></div>
              <div class="gd-elev-ticks" id="gdElevationLegendTicks">
                <span class="gd-elev-tick-max">--</span>
                <span class="gd-elev-tick-mid">--</span>
                <span class="gd-elev-tick-mid">--</span>
                <span class="gd-elev-tick-mid">--</span>
                <span class="gd-elev-tick-min">--</span>
              </div>
            </div>
            <div class="gd-elev-delta-badge" id="gdElevationDeltaBadge">ΔZ: --</div>
          </div>
        </div>
      </div>

      <div class="gd-export-hub">
        <div class="gd-export-hub-title">
          <span>▣ مركز التصدير والحفظ الاحترافي (Civil 3D / AutoCAD / Total Station / GIS)</span>
          <button type="button" id="gdExportReportHubBtn" class="gd-report-launch-btn">
            تصدير التقرير المساحي الرسمي (PDF / Excel)
          </button>
        </div>

        <div class="gd-export-selectors">
          <div class="gd-export-select-box">
            <label>1. اختيار صيغة وترتيب الأعمدة (Point Format):</label>
            <select id="gdExportColOrderSelect">
              <option value="PENZD" selected>PENZD — (نقطة، شرق E، شمال N، منسوب Z، وصف D)</option>
              <option value="PENZ">PENZ — (نقطة، شرق E، شمال N، منسوب Z)</option>
              <option value="PNEZD">PNEZD — (نقطة، شمال N، شرق E، منسوب Z، وصف D)</option>
              <option value="PEN">PEN — (نقطة، شرق E، شمال N فقط)</option>
              <option value="PNE">PNE — (نقطة، شمال N، شرق E فقط)</option>
            </select>
          </div>

          <div class="gd-export-select-box">
            <label>2. اختيار النقاط المراد تصديرها (Data Scope):</label>
            <select id="gdExportTargetScopeSelect">
              <option value="ALL" selected>تصدير كافة النقاط في الجدول (All Points)</option>
              <option value="FILTERED">تصدير النقاط المفلترة الظاهرة حالياً فقط</option>
              <option value="SEWER">تصدير نقاط المطابق والصرف الصحي الصافية فقط (MNH / SAN)</option>
              <option value="NGL">تصدير نقاط الأرض الطبيعية فقط (EG / NGL)</option>
              <option value="LIGHT">تصدير نقاط أعمدة الإنارة فقط (LP / Light Poles)</option>
              <option value="TREES">تصدير نقاط الأشجار واللاندسكيب فقط (TREE / PALM)</option>
              <option value="ROAD">تصدير نقاط الطرق والأرصفة فقط (RD)</option>
            </select>
          </div>
        </div>

        <div class="gd-ex-export-grid">
          <button type="button" class="gd-ex-exp-btn" data-export-fmt="civil3d_csv" style="border-color:#0284c7;">
            Civil 3D (.csv)
            <small>${logged ? "بدون BOM — يفتح بالـ Point Style فوراً" : "يتطلب تسجيل الدخول"}</small>
          </button>

          <button type="button" class="gd-ex-exp-btn" data-export-fmt="plain_txt" style="border-color:#38bdf8;">
            ملف نصي قياسي (.txt)
            <small>${logged ? "صيغة نصية نظيفة للأجهزة والبرامج" : "يتطلب تسجيل الدخول"}</small>
          </button>

          <button type="button" class="gd-ex-exp-btn" data-export-fmt="xlsx_clean" style="border-color:#22c55e;">
            ملف Excel نظيف (.xlsx)
            <small>${logged ? "يبدأ من أول نقطة مباشرة بلا عناوين" : "يتطلب تسجيل الدخول"}</small>
          </button>

          <button type="button" class="gd-ex-exp-btn" data-export-fmt="dxf" style="border-color:#ef4444;">
            لوحة AutoCAD (.dxf)
            <small>${logged ? "طبقات ورموز هندسية جاهزة" : "يتطلب تسجيل الدخول"}</small>
          </button>

          <button type="button" class="gd-ex-exp-btn" data-export-fmt="dwg_bundle" style="border-color:#f97316;">
            حزمة AutoCAD DWG (.scr + .dxf)
            <small>${logged ? "رسم وحفظ تلقائي بصيغة DWG" : "يتطلب تسجيل الدخول"}</small>
          </button>

          <button type="button" class="gd-ex-exp-btn" data-export-fmt="str" style="border-color:#eab308;">
            توتال ستيشن Sokkia (.str)
            <small>${logged ? "ملف STR جاهز للرفع على الجهاز" : "يتطلب تسجيل الدخول"}</small>
          </button>

          <button type="button" class="gd-ex-exp-btn" data-export-fmt="sdr" style="border-color:#ca8a04;">
            توتال ستيشن Sokkia (.sdr)
            <small>${logged ? "صيغة SDR33 للتوقيع المساحي" : "يتطلب تسجيل الدخول"}</small>
          </button>

          <button type="button" class="gd-ex-exp-btn" data-export-fmt="gsi" style="border-color:#84cc16;">
            توتال ستيشن Leica (.gsi)
            <small>${logged ? "ملف لايكا جاهز" : "يتطلب تسجيل الدخول"}</small>
          </button>

          <button type="button" class="gd-ex-exp-btn" data-export-fmt="surfer" style="border-color:#f59e0b;">
            برنامج Surfer (.dat)
            <small>${logged ? "XYZ للكنتور وحساب الكميات" : "يتطلب تسجيل الدخول"}</small>
          </button>

          <button type="button" class="gd-ex-exp-btn" data-export-fmt="geojson" style="border-color:#10b981;">
            ArcGIS Pro / QGIS (.geojson)
            <small>${logged ? "طبقة GIS ثلاثية الأبعاد" : "يتطلب تسجيل الدخول"}</small>
          </button>

          <button type="button" class="gd-ex-exp-btn" data-export-fmt="wkt" style="border-color:#06b6d4;">
            ArcGIS Pro WKT (.csv)
            <small>${logged ? "حقل POINT Z جاهز" : "يتطلب تسجيل الدخول"}</small>
          </button>

          <button type="button" class="gd-ex-exp-btn" data-export-fmt="kml" style="border-color:#a855f7;">
            Google Earth (.kml)
            <small>${logged ? "عرض النقاط على القمر الصناعي" : "يتطلب تسجيل الدخول"}</small>
          </button>
        </div>
      </div>
    `;
  }

  /* ============================================================================
     نافذة "تصدير التقرير" (اسم المشروع / نوع الأعمال / تاريخ الرفع / المساح المسؤول / حد السماحية)
     ============================================================================ */
  function openOfficialReportModal() {
    if (!checkAuthState()) {
      openAuthRequiredModal("تصدير التقرير المساحي والهندسي");
      return;
    }

    const old = document.getElementById("gdOfficialReportModal");
    if (old) old.remove();

    const defaultSurveyor =
      sessionStorage.getItem("gd_name") && sessionStorage.getItem("gd_name") !== "ضيف"
        ? sessionStorage.getItem("gd_name")
        : "أحمد سمير";
    const todayDate = new Date().toISOString().slice(0, 10);

    const overlay = document.createElement("div");
    overlay.id = "gdOfficialReportModal";
    overlay.style.cssText = `
      position: fixed; inset: 0; z-index: 999999;
      background: rgba(10, 9, 8, 0.88); backdrop-filter: blur(7px);
      display: flex; align-items: center; justify-content: center;
      padding: 14px; direction: rtl; font-family: inherit;
    `;

    overlay.innerHTML = `
      <div style="width:100%;max-width:470px;background:linear-gradient(165deg,#1c1917,#141210);border:2px solid #e09f24;border-radius:18px;padding:18px;color:#f5f5f4;box-shadow:0 24px 60px rgba(0,0,0,0.85);max-height:92vh;overflow-y:auto;">
        <div style="display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #3f3a36;padding-bottom:10px;margin-bottom:12px;">
          <div>
            <span style="display:inline-block;background:rgba(224,159,36,0.16);border:1px solid #e09f24;color:#fbbf24;font-size:11px;font-weight:900;padding:2px 9px;border-radius:999px;margin-bottom:3px;">
              تقرير مساحي معتمد
            </span>
            <h3 style="margin:0;font-size:17px;font-weight:900;color:#fbbf24;">بيانات تصدير التقرير</h3>
          </div>
          <button type="button" id="gdCloseOfficialRepBtn" style="width:32px;height:32px;border-radius:8px;border:1px solid #44403c;background:#292524;color:#d6d3d1;font-weight:900;cursor:pointer;">✕</button>
        </div>

        <div style="display:flex;flex-direction:column;gap:10px;font-size:13px;">
          <div>
            <label style="display:block;font-weight:800;color:#fbbf24;margin-bottom:4px;">اسم المشروع:</label>
            <input id="gdRepProjectName" type="text" placeholder="اكتب اسم المشروع..."
              style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;margin:0;font-family:inherit;">
          </div>

          <div>
            <label style="display:block;font-weight:800;color:#fbbf24;margin-bottom:4px;">نوع الأعمال / العمالة:</label>
            <input id="gdRepWorkType" type="text" placeholder="مثال: رفع تنفيذي ومراجعة توقيع الأعمدة..."
              style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;margin:0;font-family:inherit;">
          </div>

          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
            <div>
              <label style="display:block;font-weight:800;color:#fbbf24;margin-bottom:4px;">تاريخ الرفع:</label>
              <input id="gdRepSurveyDate" type="date" value="${todayDate}"
                style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;margin:0;font-family:inherit;">
            </div>
            <div>
              <label style="display:block;font-weight:800;color:#fbbf24;margin-bottom:4px;">حد السماحية المعتمد (مم):</label>
              <input id="gdRepToleranceMm" type="number" step="any" value="${currentToleranceMm}" placeholder="مثال: 20"
                style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;margin:0;font-family:inherit;">
            </div>
          </div>

          <div>
            <label style="display:block;font-weight:800;color:#fbbf24;margin-bottom:4px;">اسم المساح المسؤول:</label>
            <input id="gdRepSurveyorName" type="text" value="${defaultSurveyor}" placeholder="اكتب اسم المساح المسؤول..."
              style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;margin:0;font-family:inherit;">
          </div>
        </div>

        <div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:16px;">
          <button type="button" id="gdConfirmExportPdfReportBtn"
            style="flex:1;min-width:160px;padding:12px;border-radius:11px;border:0;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-size:13.5px;font-weight:900;cursor:pointer;font-family:inherit;">
            تصدير التقرير PDF
          </button>
          <button type="button" id="gdConfirmExportExcelReportBtn"
            style="flex:1;min-width:150px;padding:12px;border-radius:11px;border:1.5px solid #4ade80;background:#14532d;color:#ffffff;font-size:13.5px;font-weight:900;cursor:pointer;font-family:inherit;">
            تصدير التقرير Excel
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);
    document.getElementById("gdCloseOfficialRepBtn").onclick = () => overlay.remove();
    overlay.addEventListener("click", e => { if (e.target === overlay) overlay.remove(); });

    const getMeta = () => {
      const tol = parseFloat(document.getElementById("gdRepToleranceMm").value) || 20;
      if (designRows.length) computeAsBuiltComparison(tol);
      return {
        projectName: document.getElementById("gdRepProjectName").value.trim() || "مشروع رفع مساحي",
        workType: document.getElementById("gdRepWorkType").value.trim() || "أعمال رفع وتوقيع مساحي",
        surveyDate: document.getElementById("gdRepSurveyDate").value || todayDate,
        surveyorName: document.getElementById("gdRepSurveyorName").value.trim() || defaultSurveyor,
        toleranceMm: tol
      };
    };

    document.getElementById("gdConfirmExportPdfReportBtn").onclick = () => {
      const meta = getMeta();
      overlay.remove();
      generateOfficialPdfReport(meta);
    };

    document.getElementById("gdConfirmExportExcelReportBtn").onclick = () => {
      const meta = getMeta();
      overlay.remove();
      generateOfficialExcelReport(meta);
    };
  }

  function generateOfficialPdfReport(meta) {
    const st = computeSurveyStats();
    const { pKey, eKey, nKey, zKey, cKey } = getCoordKeys();
    const hasComp = asBuiltResults.length > 0;
    const passed = asBuiltResults.filter(c => c.isPass).length;
    const failed = asBuiltResults.filter(c => !c.isPass).length;
    const maxShift = hasComp ? Math.max(...asBuiltResults.map(c => c.dXY_mm)) : 0;

    const tableHead = hasComp
      ? `<tr>
          <th>م</th><th>رقم النقطة</th><th>الكود</th>
          <th>E التصميمي</th><th>N التصميمي</th>
          <th>E المرفوع</th><th>N المرفوع</th>
          <th>فرق E (مم)</th><th>فرق N (مم)</th><th>الإزاحة XY (مم)</th><th>فرق Z (مم)</th><th>حالة السماحية</th>
        </tr>`
      : `<tr>
          <th>م</th><th>رقم النقطة (P)</th><th>الشرق (E)</th><th>الشمال (N)</th><th>المنسوب (Z)</th><th>الكود</th><th>التصنيف الهندسي</th>
        </tr>`;

    const tableBody = hasComp
      ? asBuiltResults.map((c, i) => `
          <tr>
            <td>${i + 1}</td>
            <td><b>${c.P}</b></td>
            <td>${c.Code}</td>
            <td>${c.desE.toFixed(3)}</td>
            <td>${c.desN.toFixed(3)}</td>
            <td>${c.actE.toFixed(3)}</td>
            <td>${c.actN.toFixed(3)}</td>
            <td>${c.dE_mm.toFixed(1)}</td>
            <td>${c.dN_mm.toFixed(1)}</td>
            <td><b>${c.dXY_mm.toFixed(1)}</b></td>
            <td>${c.dZ_mm.toFixed(1)}</td>
            <td style="font-weight:bold;color:${c.isPass ? "#15803d" : "#b91c1c"};">${c.statusText}</td>
          </tr>
        `).join("")
      : rows.map((r, i) => {
          const cls = getRowClassification(r);
          return `
            <tr>
              <td>${i + 1}</td>
              <td><b>${r[pKey] || i + 1}</b></td>
              <td>${(parseFloat(r[eKey]) || 0).toFixed(3)}</td>
              <td>${(parseFloat(r[nKey]) || 0).toFixed(3)}</td>
              <td>${(parseFloat(r[zKey]) || 0).toFixed(3)}</td>
              <td>${r[cKey] || ""}</td>
              <td>${cls.labelAr}</td>
            </tr>
          `;
        }).join("");

    const htmlDoc = `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<title>تقرير مساحي - ${meta.projectName}</title>
<style>
  body { font-family: Tahoma, Arial, sans-serif; direction: rtl; color: #0f172a; margin: 24px; background: #fff; }
  .box { border: 2px solid #0f172a; border-radius: 10px; padding: 14px 18px; margin-bottom: 16px; background: #f8fafc; }
  .box h1 { margin: 0 0 8px; font-size: 19px; color: #0f172a; }
  .meta { display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px 18px; font-size: 13px; }
  .meta b { color: #b45309; }
  .stats { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; }
  .stat { flex: 1; min-width: 120px; border: 1px solid #cbd5e1; border-radius: 8px; padding: 9px; text-align: center; background: #f8fafc; }
  .stat span { display: block; font-size: 11px; color: #475569; }
  .stat b { font-size: 15px; color: #0f172a; }
  table { width: 100%; border-collapse: collapse; font-size: 11.5px; margin-bottom: 24px; }
  th, td { border: 1px solid #94a3b8; padding: 6px; text-align: center; }
  th { background: #0f172a; color: #fff; }
  tr:nth-child(even) { background: #f8fafc; }
  .sign { display: flex; justify-content: space-between; margin-top: 28px; font-size: 13.5px; font-weight: bold; }
  @media print { .no-print { display: none; } }
</style>
</head>
<body>
  <div class="no-print" style="margin-bottom:12px;text-align:left;">
    <button onclick="window.print()" style="padding:10px 18px;background:#0f172a;color:#fbbf24;border:0;border-radius:8px;font-weight:bold;cursor:pointer;">
      طباعة / حفظ كـ PDF
    </button>
  </div>

  <div class="box">
    <h1>تقرير الرفع والمراجعة المساحية — GeoDev Studio Pro</h1>
    <div class="meta">
      <div><b>اسم المشروع:</b> ${meta.projectName}</div>
      <div><b>نوع الأعمال:</b> ${meta.workType}</div>
      <div><b>تاريخ الرفع:</b> ${meta.surveyDate}</div>
      <div><b>اسم المساح المسؤول:</b> ${meta.surveyorName}</div>
      <div><b>حد السماحية المعتمد:</b> ± ${meta.toleranceMm} مم</div>
      <div><b>المساحة المغلقة:</b> ${st.areaM2.toFixed(2)} م²</div>
    </div>
  </div>

  <div class="stats">
    <div class="stat"><span>إجمالي النقاط</span><b>${st.count}</b></div>
    <div class="stat"><span>أقل منسوب (Z)</span><b>${st.minZ}</b></div>
    <div class="stat"><span>أعلى منسوب (Z)</span><b>${st.maxZ}</b></div>
    <div class="stat"><span>مطابق (MNH)</span><b>${st.sewerCount}</b></div>
    ${hasComp ? `
      <div class="stat"><span>مطابق للسماحية</span><b style="color:#15803d;">${passed}</b></div>
      <div class="stat"><span>تجاوز السماحية</span><b style="color:#b91c1c;">${failed}</b></div>
      <div class="stat"><span>أقصى إزاحة XY</span><b>${maxShift.toFixed(1)} مم</b></div>
    ` : ""}
  </div>

  <table>
    <thead>${tableHead}</thead>
    <tbody>${tableBody}</tbody>
  </table>

  <div class="sign">
    <div>مساح: ${meta.surveyorName}</div>
    <div>اعتماد الاستشاري: ............................</div>
  </div>
</body>
</html>`;

    const win = window.open("", "_blank");
    if (win) {
      win.document.write(htmlDoc);
      win.document.close();
    } else {
      downloadBlob(htmlDoc, `GeoDev_Report_${meta.surveyDate}.html`, "text/html;charset=utf-8;");
    }
    showExcelLog(`تم إصدار التقرير المساحي الرسمي لمشروع «${meta.projectName}» باسم مساح: ${meta.surveyorName}.`, "success");
  }

  function generateOfficialExcelReport(meta) {
    ensureSheetJS(() => {
      const { pKey, eKey, nKey, zKey, cKey } = getCoordKeys();
      const aoa = [
        ["اسم المشروع", meta.projectName],
        ["نوع الأعمال", meta.workType],
        ["تاريخ الرفع", meta.surveyDate],
        ["مساح", meta.surveyorName],
        ["حد السماحية المعتمد (مم)", meta.toleranceMm],
        []
      ];

      if (asBuiltResults.length) {
        aoa.push(["م", "رقم النقطة", "الكود", "E التصميمي", "N التصميمي", "E المرفوع", "N المرفوع", "فرق E (مم)", "فرق N (مم)", "الإزاحة XY (مم)", "فرق Z (مم)", "حالة السماحية"]);
        asBuiltResults.forEach((c, i) => {
          aoa.push([i + 1, c.P, c.Code, Number(c.desE.toFixed(3)), Number(c.desN.toFixed(3)), Number(c.actE.toFixed(3)), Number(c.actN.toFixed(3)), Number(c.dE_mm.toFixed(1)), Number(c.dN_mm.toFixed(1)), Number(c.dXY_mm.toFixed(1)), Number(c.dZ_mm.toFixed(1)), c.statusText]);
        });
      } else {
        aoa.push(["م", "رقم النقطة (P)", "الشرق (E)", "الشمال (N)", "المنسوب (Z)", "الكود", "التصنيف الهندسي"]);
        rows.forEach((r, i) => {
          const cls = getRowClassification(r);
          aoa.push([i + 1, r[pKey] || i + 1, parseFloat(r[eKey]) || 0, parseFloat(r[nKey]) || 0, parseFloat(r[zKey]) || 0, r[cKey] || "", cls.labelAr]);
        });
      }

      const ws = window.XLSX.utils.aoa_to_sheet(aoa);
      const wb = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(wb, ws, "Survey_Report");
      window.XLSX.writeFile(wb, `GeoDev_Report_${meta.surveyDate}.xlsx`);
      showExcelLog(`تم تصدير التقرير المساحي إلى Excel بنجاح.`, "success");
    });
  }

  function getCoordKeys() {
    const pKey = columns.find(c => /^(p|pt|point|id|no|num|رقم|النقطة)$/i.test(c.trim())) || columns[0];
    const eKey = columns.find(c => /^(e|east|easting|x|lon|long|شرق)$/i.test(c.trim())) || columns[1];
    const nKey = columns.find(c => /^(n|north|northing|norht|y|lat|شمال)$/i.test(c.trim())) || columns[2];
    const zKey = columns.find(c => /^(z|elev|elevation|rl|level|منسوب)$/i.test(c.trim())) || columns[3];
    const cKey = columns.find(c => /^(code|desc|description|note|كود|وصف)$/i.test(c.trim())) || columns[4];
    return { pKey, eKey, nKey, zKey, cKey };
  }

  function computeSurveyStats() {
    const { eKey, nKey, zKey } = getCoordKeys();
    let sewerCount = 0, lightCount = 0, treeCount = 0, dirCount = 0, nglCount = 0, roadCount = 0;

    const validPts = [];
    rows.forEach(r => {
      const cls = getRowClassification(r);
      if (cls.group === "SEWER") sewerCount++;
      else if (cls.group === "LIGHT") lightCount++;
      else if (cls.group === "TREES") treeCount++;
      else if (cls.group === "DIR") dirCount++;
      else if (cls.group === "NGL") nglCount++;
      else if (cls.group === "ROAD") roadCount++;

      const e = parseFloat(r[eKey]);
      const n = parseFloat(r[nKey]);
      const z = parseFloat(r[zKey]);
      if (!isNaN(e) && !isNaN(n)) validPts.push({ e, n, z });
    });

    if (!validPts.length) {
      return { count: rows.length, minZ: "-", maxZ: "-", perimeter: 0, areaM2: 0, sewerCount, lightCount, treeCount, dirCount, nglCount, roadCount };
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
    return {
      count: rows.length,
      minZ,
      maxZ,
      perimeter,
      areaM2: Math.abs(area2) / 2,
      sewerCount,
      lightCount,
      treeCount,
      dirCount,
      nglCount,
      roadCount
    };
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

  function handleDesignFileUpload(file) {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();

    const finishDesignImport = parsedList => {
      if (!parsedList.length) {
        showExcelLog("لم يتم العثور على نقاط صالحة داخل ملف التصميم.", "error");
        return;
      }
      designRows = parsedList;
      const tol = parseFloat(document.getElementById("gdAsBuiltTolInput")?.value) || currentToleranceMm;
      const res = computeAsBuiltComparison(tol);
      const passed = res.filter(r => r.isPass).length;
      const failed = res.filter(r => !r.isPass).length;

      const sumBox = document.getElementById("gdAsBuiltSummaryBox");
      if (sumBox) {
        sumBox.innerHTML = `<strong style="color:#fbbf24;">تم تحميل ${designRows.length} نقطة تصميمية:</strong> مطابقة ${res.length} نقطة — <span style="color:#4ade80;font-weight:900;">مقبول: ${passed}</span> | <span style="color:#f87171;font-weight:900;">تجاوز السماحية (±${tol} مم): ${failed}</span>`;
      }
      renderExcelTable();
      drawCadPreview();
      showExcelLog(`تم تحميل ملف التصميم (${file.name}) ومقارنته بالرفع الجديد (${passed} مقبول / ${failed} تجاوز السماحية).`, "success");
    };

    if (ext === "xlsx" || ext === "xls") {
      ensureSheetJS(() => {
        const reader = new FileReader();
        reader.onload = e => {
          const wb = window.XLSX.read(e.target.result, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const aoa = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
          const parsed = [];
          aoa.forEach((r, idx) => {
            if (Array.isArray(r) && r.length >= 3 && !isNaN(parseFloat(r[1])) && !isNaN(parseFloat(r[2]))) {
              parsed.push({
                P: String(r[0] || idx + 1).trim(),
                E: parseFloat(r[1]),
                N: parseFloat(r[2]),
                Z: parseFloat(r[3]) || 0,
                Code: String(r[4] || "DESIGN").trim()
              });
            }
          });
          finishDesignImport(parsed);
        };
        reader.readAsArrayBuffer(file);
      });
      return;
    }

    const reader = new FileReader();
    reader.onload = e => {
      const text = String(e.target.result || "");
      if (ext === "dxf") {
        finishDesignImport(parseDxfContent(text));
        return;
      }
      const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
      const sep = lines[0]?.includes("\t") ? "\t" : lines[0]?.includes(";") ? ";" : ",";
      const parsed = [];
      lines.forEach((ln, idx) => {
        const parts = ln.split(sep);
        if (parts.length >= 3 && !isNaN(parseFloat(parts[1])) && !isNaN(parseFloat(parts[2]))) {
          parsed.push({
            P: String(parts[0] || idx + 1).trim(),
            E: parseFloat(parts[1]),
            N: parseFloat(parts[2]),
            Z: parseFloat(parts[3]) || 0,
            Code: String(parts[4] || "DESIGN").trim()
          });
        }
      });
      finishDesignImport(parsed);
    };
    reader.readAsText(file);
  }

  function handleFileUpload(file) {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    saveSnapshot();
    activeCodeFilter = "ALL";

    if (ext === "xlsx" || ext === "xls") {
      ensureSheetJS(() => {
        const reader = new FileReader();
        reader.onload = e => {
          try {
            const wb = window.XLSX.read(e.target.result, { type: "array" });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const aoa = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
            if (!aoa.length) return showExcelLog("ملف الإكسيل فارغ.", "error");

            const dataRows = aoa.filter(r => {
              if (!Array.isArray(r) || r.length < 3) return false;
              const v1 = parseFloat(r[1]);
              const v2 = parseFloat(r[2]);
              return !isNaN(v1) && !isNaN(v2);
            });

            if (dataRows.length) {
              const maxCols = Math.max(...dataRows.map(r => r.length));
              const defaultHeaders = ["P", "E", "N", "Z", "Code", "Type", "Pole_Name", "Pole_H"];
              columns = Array.from({ length: maxCols }, (_, i) => defaultHeaders[i] || `Col_${i + 1}`);
              rows = dataRows.map((r, idx) => {
                const obj = {};
                columns.forEach((c, cIdx) => {
                  obj[c] = r[cIdx] !== undefined && r[cIdx] !== "" ? r[cIdx] : (cIdx === 0 ? String(idx + 1) : "");
                });
                return obj;
              });
            } else {
              const jsonRows = window.XLSX.utils.sheet_to_json(sheet, { defval: "" });
              columns = Object.keys(jsonRows[0] || {});
              rows = jsonRows;
            }

            if (designRows.length) computeAsBuiltComparison();
            const st = computeSurveyStats();
            resetCadView();
            renderExcelTable();
            showExcelLog(
              `تم استيراد ${rows.length} نقطة وتصنيفها وتلوين المناسيب من ${st.minZ}م إلى ${st.maxZ}م.`,
              "success"
            );
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
        if (!pts.length) return showExcelLog("لم يتم العثور على نقاط داخل ملف DXF.", "error");
        columns = ["P", "E", "N", "Z", "Code"];
        rows = pts;
        if (designRows.length) computeAsBuiltComparison();
        resetCadView();
        renderExcelTable();
        showExcelLog(`تم استيراد ${rows.length} نقطة من لوحة AutoCAD DXF (${file.name}).`, "success");
        return;
      }

      const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
      if (!lines.length) return;

      if (ext === "sdr" || ext === "str" || ext === "cdr") {
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
          if (designRows.length) computeAsBuiltComparison();
          resetCadView();
          renderExcelTable();
          showExcelLog(`تم استيراد ${rows.length} نقطة من ملف التوتال ستيشن (${file.name}).`, "success");
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
          if (designRows.length) computeAsBuiltComparison();
          resetCadView();
          renderExcelTable();
          showExcelLog(`تم استيراد ${rows.length} نقطة من ملف Leica GSI.`, "success");
          return;
        }
      }

      const sep = lines[0].includes("\t") ? "\t" : lines[0].includes(";") ? ";" : ",";
      const numericLines = lines.filter(ln => {
        const parts = ln.split(sep);
        return parts.length >= 3 && !isNaN(parseFloat(parts[1])) && !isNaN(parseFloat(parts[2]));
      });

      const targetLines = numericLines.length ? numericLines : lines;
      const sampleParts = targetLines[0].split(sep);
      columns = ["P", "E", "N", "Z", "Code"].slice(0, Math.max(3, sampleParts.length));
      rows = targetLines.map((ln, rIdx) => {
        const vals = ln.split(sep);
        const obj = {};
        columns.forEach((c, idx) => (obj[c] = (vals[idx] || (idx === 0 ? String(rIdx + 1) : "")).trim()));
        return obj;
      });
      if (designRows.length) computeAsBuiltComparison();
      resetCadView();
      renderExcelTable();
      showExcelLog(`تم استيراد ${rows.length} نقطة من ${file.name}.`, "success");
    };
    reader.readAsText(file);
  }

  function getRowsForExportScope() {
    const scope = document.getElementById("gdExportTargetScopeSelect")?.value || exportTargetScope;
    if (scope === "FILTERED") {
      return getFilteredRowsWithIndex().map(item => item.r);
    }
    if (scope === "ALL") return rows;

    return rows.filter(r => {
      const cls = getRowClassification(r);
      if (scope === "LIGHT") return cls.group === "LIGHT";
      if (scope === "SEWER") return cls.group === "SEWER";
      if (scope === "NGL") return cls.group === "NGL";
      if (scope === "TREES") return cls.group === "TREES";
      if (scope === "ROAD") return cls.group === "ROAD";
      return true;
    });
  }

  function formatPointTuple(rowObj, idx, orderMode) {
    const { pKey, eKey, nKey, zKey, cKey } = getCoordKeys();
    const rawP = rowObj[pKey] !== undefined && String(rowObj[pKey]).trim() !== "" ? String(rowObj[pKey]).trim() : String(idx + 1);
    const p = rawP.replace(/[^\w\-.]/g, "") || String(idx + 1);
    const eNum = parseFloat(rowObj[eKey]);
    const nNum = parseFloat(rowObj[nKey]);
    const zNum = parseFloat(rowObj[zKey]);
    if (isNaN(eNum) || isNaN(nNum)) return null;

    const e = eNum.toFixed(4);
    const n = nNum.toFixed(4);
    const z = (isNaN(zNum) ? 0 : zNum).toFixed(4);
    const rawD = rowObj[cKey] !== undefined ? String(rowObj[cKey]).trim() : "PT";
    const d = rawD.replace(/[,;"'\r\n]+/g, " ").trim() || "PT";

    if (orderMode === "PENZ") return [p, e, n, z];
    if (orderMode === "PNEZD") return [p, n, e, z, d];
    if (orderMode === "PEN") return [p, e, n];
    if (orderMode === "PNE") return [p, n, e];
    return [p, e, n, z, d];
  }

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

  function buildDxfString(targetRows) {
    const { pKey, eKey, nKey, zKey, cKey } = getCoordKeys();
    const dxf = ["0", "SECTION", "2", "ENTITIES"];
    const groupedPolylines = {};

    targetRows.forEach((r, idx) => {
      const e = parseFloat(r[eKey]);
      const n = parseFloat(r[nKey]);
      const z = parseFloat(r[zKey]) || 0;
      const p = String(r[pKey] || idx + 1);
      const c = String(r[cKey] || "");
      if (isNaN(e) || isNaN(n)) return;

      const cls = getRowClassification(r);
      if (!groupedPolylines[cls.layer]) groupedPolylines[cls.layer] = [];
      groupedPolylines[cls.layer].push({ e, n, z, color: cls.dxfColor });

      dxf.push("0", "POINT", "8", cls.layer, "62", String(cls.dxfColor), "10", e.toFixed(4), "20", n.toFixed(4), "30", z.toFixed(4));

      if (cls.symbol === "MANHOLE") {
        dxf.push("0", "CIRCLE", "8", cls.layer, "62", "30", "10", e.toFixed(4), "20", n.toFixed(4), "30", z.toFixed(4), "40", "0.60");
      } else if (cls.symbol === "LIGHT") {
        dxf.push("0", "CIRCLE", "8", cls.layer, "62", "2", "10", e.toFixed(4), "20", n.toFixed(4), "30", z.toFixed(4), "40", "0.40");
      } else if (cls.symbol === "TREE") {
        dxf.push("0", "CIRCLE", "8", cls.layer, "62", "3", "10", e.toFixed(4), "20", n.toFixed(4), "30", z.toFixed(4), "40", "0.75");
      }

      dxf.push("0", "TEXT", "8", `${cls.layer}_ID`, "62", String(cls.dxfColor), "10", (e + 0.45).toFixed(4), "20", (n + 0.35).toFixed(4), "30", z.toFixed(4), "40", "0.50", "1", `${p} ${c}`.trim());
      dxf.push("0", "TEXT", "8", `${cls.layer}_ELEV`, "62", "7", "10", (e + 0.45).toFixed(4), "20", (n - 0.35).toFixed(4), "30", z.toFixed(4), "40", "0.42", "1", z.toFixed(3));
    });

    Object.entries(groupedPolylines).forEach(([layerName, pts]) => {
      if (pts.length >= 2 && layerName !== "GD_LIGHT_POLES" && layerName !== "GD_TREES" && layerName !== "GD_BENCHMARKS") {
        dxf.push("0", "POLYLINE", "8", `${layerName}_LINE`, "62", String(pts[0].color), "66", "1", "70", "8");
        pts.forEach(pt => {
          dxf.push("0", "VERTEX", "8", `${layerName}_LINE`, "10", pt.e.toFixed(4), "20", pt.n.toFixed(4), "30", pt.z.toFixed(4));
        });
        dxf.push("0", "SEQEND");
      }
    });

    dxf.push("0", "ENDSEC", "0", "EOF");
    return dxf.join("\r\n") + "\r\n";
  }

  function exportData(format) {
    if (!checkAuthState()) {
      openAuthRequiredModal("حفظ وتصدير الملفات بصيغ Civil 3D و AutoCAD و Excel");
      return;
    }

    const orderMode = document.getElementById("gdExportColOrderSelect")?.value || exportColOrder;
    const scopeMode = document.getElementById("gdExportTargetScopeSelect")?.value || exportTargetScope;
    const targetRows = getRowsForExportScope();

    if (!targetRows.length) {
      showExcelLog("لا توجد نقاط مطابقة للفئة المختارة للتصدير.", "warn");
      return;
    }

    const { pKey, eKey, nKey, zKey, cKey } = getCoordKeys();

    if (format === "civil3d_csv") {
      const lines = [];
      targetRows.forEach((r, idx) => {
        const tuple = formatPointTuple(r, idx, orderMode);
        if (tuple) lines.push(tuple.join(","));
      });
      downloadBlob(lines.join("\r\n") + "\r\n", `GeoDev_Civil3D_${orderMode}_${scopeMode}_${Date.now()}.csv`, "text/csv;charset=ascii;");
      showExcelLog(`تم تصدير ${lines.length} نقطة صافية بصيغة Civil 3D (${orderMode}) بدون BOM.`, "success");
      return;
    }

    if (format === "plain_txt") {
      const lines = [];
      targetRows.forEach((r, idx) => {
        const tuple = formatPointTuple(r, idx, orderMode);
        if (tuple) lines.push(tuple.join(","));
      });
      downloadBlob(lines.join("\r\n") + "\r\n", `GeoDev_${orderMode}_${scopeMode}_${Date.now()}.txt`, "text/plain;charset=ascii;");
      showExcelLog(`تم تصدير ${lines.length} نقطة إلى ملف نصي (${orderMode} .txt) بنجاح.`, "success");
      return;
    }

    if (format === "xlsx_clean") {
      ensureSheetJS(() => {
        const aoaData = [];
        targetRows.forEach((r, idx) => {
          const tuple = formatPointTuple(r, idx, orderMode);
          if (!tuple) return;
          const numericRow = tuple.map((val, colIdx) => {
            if (colIdx === 0) return val;
            if (orderMode.endsWith("D") && colIdx === tuple.length - 1) return val;
            const n = parseFloat(val);
            return isNaN(n) ? val : n;
          });
          aoaData.push(numericRow);
        });
        const ws = window.XLSX.utils.aoa_to_sheet(aoaData);
        const wb = window.XLSX.utils.book_new();
        window.XLSX.utils.book_append_sheet(wb, ws, "Points");
        window.XLSX.writeFile(wb, `GeoDev_Excel_${orderMode}_${scopeMode}_${Date.now()}.xlsx`);
        showExcelLog(`تم تصدير ${aoaData.length} نقطة إلى Excel نظيف بصيغة (${orderMode}) يبدأ من النقطة 1 مباشرة.`, "success");
      });
      return;
    }

    if (format === "dxf") {
      const dxfStr = buildDxfString(targetRows);
      downloadBlob(dxfStr, `GeoDev_AutoCAD_${scopeMode}_${Date.now()}.dxf`, "application/dxf");
      showExcelLog(`تم تصدير لوحة AutoCAD DXF (${targetRows.length} نقطة) بطبقات ورموز هندسية جاهزة.`, "success");
      return;
    }

    if (format === "dwg_bundle") {
      const dxfStr = buildDxfString(targetRows);
      const stamp = Date.now();
      downloadBlob(dxfStr, `GeoDev_DWG_Source_${stamp}.dxf`, "application/dxf");

      const scrLines = ["_OSMODE 0", "_PDMODE 34", "_PDSIZE 0.5"];
      targetRows.forEach((r, idx) => {
        const e = parseFloat(r[eKey]);
        const n = parseFloat(r[nKey]);
        const z = parseFloat(r[zKey]) || 0;
        const p = String(r[pKey] || idx + 1);
        const c = String(r[cKey] || "");
        if (!isNaN(e) && !isNaN(n)) {
          scrLines.push(`_POINT ${e.toFixed(4)},${n.toFixed(4)},${z.toFixed(4)}`);
          scrLines.push(`_TEXT ${(e + 0.4).toFixed(4)},${(n + 0.35).toFixed(4)} 0.5 0 ${p} ${c}`.trim());
        }
      });
      scrLines.push("_ZOOM _E");
      scrLines.push(`_SAVEAS 2018 GeoDev_Drawing_${stamp}.dwg`);
      setTimeout(() => {
        downloadBlob(scrLines.join("\r\n") + "\r\n", `GeoDev_AutoSaveDWG_${stamp}.scr`, "text/plain;charset=ascii;");
      }, 350);
      showExcelLog("تم تحميل ملف اللوحة وسكريبت الحفظ التلقائي بصيغة AutoCAD DWG.", "success");
      return;
    }

    if (format === "str" || format === "sdr") {
      const extName = format === "str" ? "str" : "sdr";
      const sdrLines = ["00NMSDR33 V04-04.02  GEODEV STAKEOUT"];
      targetRows.forEach((r, idx) => {
        const p = String(r[pKey] || idx + 1).padEnd(16, " ").slice(0, 16);
        const n = (parseFloat(r[nKey]) || 0).toFixed(3).padStart(16, " ");
        const e = (parseFloat(r[eKey]) || 0).toFixed(3).padStart(16, " ");
        const z = (parseFloat(r[zKey]) || 0).toFixed(3).padStart(16, " ");
        const c = String(r[cKey] || "").padEnd(16, " ").slice(0, 16);
        sdrLines.push(`08KI${p}${n}${e}${z}${c}`);
      });
      downloadBlob(sdrLines.join("\r\n") + "\r\n", `GeoDev_TotalStation_${scopeMode}_${Date.now()}.${extName}`, "text/plain;charset=ascii;");
      showExcelLog(`تم تصدير ${targetRows.length} نقطة بصيغة التوتال ستيشن (.${extName}) جاهزة للرفع على الجهاز.`, "success");
      return;
    }

    if (format === "gsi") {
      const gsiLines = targetRows.map((r, idx) => {
        const p = String(r[pKey] || idx + 1).padStart(16, "0");
        const eMm = Math.round((parseFloat(r[eKey]) || 0) * 1000);
        const nMm = Math.round((parseFloat(r[nKey]) || 0) * 1000);
        const zMm = Math.round((parseFloat(r[zKey]) || 0) * 1000);
        const fmt = val => (val >= 0 ? "+" : "-") + String(Math.abs(val)).padStart(16, "0");
        return `11${String(idx + 1).padStart(4, "0")}+${p} 81..10${fmt(eMm)} 82..10${fmt(nMm)} 83..10${fmt(zMm)}`;
      });
      downloadBlob(gsiLines.join("\r\n") + "\r\n", `GeoDev_Leica_${scopeMode}_${Date.now()}.gsi`, "text/plain;charset=ascii;");
      showExcelLog("تم تصدير ملف Leica GSI (.gsi) بنجاح.", "success");
      return;
    }

    if (format === "surfer") {
      const datLines = ["Easting,Northing,Elevation,PointID,Code"];
      targetRows.forEach((r, idx) => {
        const e = parseFloat(r[eKey]);
        const n = parseFloat(r[nKey]);
        const z = parseFloat(r[zKey]) || 0;
        const p = String(r[pKey] || idx + 1);
        const c = String(r[cKey] || "");
        if (!isNaN(e) && !isNaN(n)) datLines.push(`${e.toFixed(4)},${n.toFixed(4)},${z.toFixed(4)},${p},${c}`);
      });
      downloadBlob(datLines.join("\r\n") + "\r\n", `GeoDev_Surfer_${scopeMode}_${Date.now()}.dat`, "text/plain;charset=ascii;");
      showExcelLog("تم تصدير ملف Golden Software Surfer (.dat) بنجاح.", "success");
      return;
    }

    if (format === "wkt") {
      const wktHeader = [...columns, "AI_Category", "CAD_Layer", "GEOMETRY_WKT"].join(",");
      const wktRows = targetRows.map(r => {
        const e = parseFloat(r[eKey]);
        const n = parseFloat(r[nKey]);
        const z = parseFloat(r[zKey]) || 0;
        const cls = getRowClassification(r);
        const geom = !isNaN(e) && !isNaN(n) ? `"POINT Z (${e.toFixed(4)} ${n.toFixed(4)} ${z.toFixed(4)})"` : '""';
        const vals = columns.map(col => `"${String(r[col] ?? "").replace(/"/g, '""')}"`);
        return [...vals, `"${cls.labelAr}"`, `"${cls.layer}"`, geom].join(",");
      });
      downloadBlob("\uFEFF" + [wktHeader, ...wktRows].join("\n"), `GeoDev_ArcGISPro_WKT_${Date.now()}.csv`, "text/csv;charset=utf-8;");
      showExcelLog("تم تصدير جدول ArcGIS Pro بالتصنيف الهندسي وحقل POINT Z.", "success");
      return;
    }

    if (format === "geojson") {
      const features = targetRows
        .map(r => {
          const e = parseFloat(r[eKey]);
          const n = parseFloat(r[nKey]);
          if (isNaN(e) || isNaN(n)) return null;
          const ll = projectToWGS84(e, n);
          const cls = getRowClassification(r);
          return {
            type: "Feature",
            properties: { ...r, AI_Category: cls.labelAr, CAD_Layer: cls.layer, CRS_Source: crsMode },
            geometry: {
              type: "Point",
              coordinates: [Number(ll.lon.toFixed(7)), Number(ll.lat.toFixed(7)), parseFloat(r[zKey]) || 0]
            }
          };
        })
        .filter(Boolean);

      downloadBlob(JSON.stringify({ type: "FeatureCollection", features }, null, 2), `GeoDev_GIS_${scopeMode}_${Date.now()}.geojson`, "application/geo+json;charset=utf-8;");
      showExcelLog("تم تصدير طبقة GeoJSON لـ ArcGIS Pro و QGIS بنجاح.", "success");
      return;
    }

    if (format === "kml") {
      const marks = targetRows
        .map((r, idx) => {
          const e = parseFloat(r[eKey]);
          const n = parseFloat(r[nKey]);
          const z = parseFloat(r[zKey]) || 0;
          if (isNaN(e) || isNaN(n)) return "";
          const ll = projectToWGS84(e, n);
          const cls = getRowClassification(r);
          return `<Placemark><name>${String(r[pKey] || idx + 1)} (${cls.labelAr})</name><description>E=${e} | N=${n} | Z=${z}</description><Point><coordinates>${ll.lon.toFixed(7)},${ll.lat.toFixed(7)},${z}</coordinates></Point></Placemark>`;
        })
        .join("\n");
      const kml = `<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>GeoDev Survey</name>\n${marks}\n</Document></kml>`;
      downloadBlob(kml, `GeoDev_GoogleEarth_${scopeMode}_${Date.now()}.kml`, "application/vnd.google-earth.kml+xml;charset=utf-8;");
      showExcelLog("تم تصدير ملف Google Earth (.kml) بنجاح.", "success");
    }
  }

  function runAutoCodeClassification() {
    saveSnapshot();
    if (!columns.includes("التصنيف_الهندسي")) columns.push("التصنيف_الهندسي");
    if (!columns.includes("طبقة_الأوتوكاد")) columns.push("طبقة_الأوتوكاد");

    rows.forEach(r => {
      const cls = getRowClassification(r);
      r["التصنيف_الهندسي"] = cls.labelAr;
      r["طبقة_الأوتوكاد"] = cls.layer;
    });

    const st = computeSurveyStats();
    renderExcelTable();
    drawCadPreview();
    showExcelLog(
      `تم تصنيف النقاط بدقة: [${st.sewerCount} مطبق MNH صافي] • [${st.nglCount} أرض طبيعية EG] • [${st.lightCount} إنارة] • [${st.treeCount} أشجار].`,
      "success"
    );
  }

  function runSurveyQAQC() {
    const { eKey, nKey, zKey, cKey } = getCoordKeys();
    let zeroCoords = 0, missingCodes = 0, zJumps = 0, duplicates = 0;
    const seen = new Set();

    for (let i = 0; i < rows.length; i++) {
      const e = parseFloat(rows[i][eKey]);
      const n = parseFloat(rows[i][nKey]);
      const z = parseFloat(rows[i][zKey]);
      const code = String(rows[i][cKey] || "").trim();

      if (isNaN(e) || isNaN(n) || e === 0 || n === 0) zeroCoords++;
      if (!code || code === "----" || code === "-") missingCodes++;

      const key = `${e}_${n}`;
      if (seen.has(key)) duplicates++;
      seen.add(key);

      if (i > 0) {
        const prevZ = parseFloat(rows[i - 1][zKey]);
        if (!isNaN(z) && !isNaN(prevZ) && Math.abs(z - prevZ) > 5) zJumps++;
      }
    }

    showExcelLog(
      `تقرير فحص الجودة (QA/QC): [${duplicates} مكررة] • [${missingCodes} بدون كود] • [${zJumps} قفزة شاذة في Z > 5م] • [${zeroCoords} إحداثي صفري].`,
      duplicates || zJumps || zeroCoords ? "warn" : "success"
    );
  }

  function executeAICommand(rawPrompt) {
    const cmd = String(rawPrompt || "").trim();
    if (!cmd) return;

    if (!checkAuthState()) {
      openAuthRequiredModal("تشغيل أوامر الذكاء الاصطناعي وحفظ التعديلات");
      return;
    }

    if (/تصنيف|صنف|اكواد|أكواد|mnh|انارة|إنارة|صرف|شجر/i.test(cmd)) {
      runAutoCodeClassification();
      return;
    }

    if (/فحص|جودة|اخطاء|أخطاء|qa|qc/i.test(cmd)) {
      runSurveyQAQC();
      return;
    }

    saveSnapshot();
    const { pKey, eKey, nKey, zKey } = getCoordKeys();

    if (/ترقيم|رقم\s*النقاط/i.test(cmd)) {
      rows.forEach((r, idx) => { r[pKey] = String(idx + 1); });
      renderExcelTable();
      drawCadPreview();
      showExcelLog(`تم إعادة ترقيم جميع النقاط تسلسلياً من 1 إلى ${rows.length}.`, "success");
      return;
    }

    if (/تبديل|بدل\s*الشرق/i.test(cmd)) {
      rows.forEach(r => {
        const tmp = r[eKey];
        r[eKey] = r[nKey];
        r[nKey] = tmp;
      });
      renderExcelTable();
      resetCadView();
      showExcelLog("تم تبديل إحداثيات الشرق (E) والشمال (N) لجميع النقاط.", "success");
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
      showExcelLog(`تم تعديل عمود «${targetCol}» (${op} ${num}) لجميع النقاط.`, "success");
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
      showExcelLog("تم حساب المسافة والانحراف وإضافتهما كأعمدة جديدة.", "success");
      return;
    }

    if (/مساح|محيط|area|perim/i.test(cmd)) {
      const st = computeSurveyStats();
      const feddan = st.areaM2 / 4200.83;
      const qirat = (feddan % 1) * 24;
      const sahm = (qirat % 1) * 24;
      showExcelLog(`المساحة = ${st.areaM2.toFixed(2)} م² (${Math.floor(feddan)} فدان، ${Math.floor(qirat)} قيراط، ${sahm.toFixed(1)} سهم) | المحيط = ${st.perimeter.toFixed(2)} م`, "success");
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
      showExcelLog(`تم حذف ${before - rows.length} نقطة مكررة.`, "success");
      return;
    }

    if (/حفظ|احفظ|تصدير|صدر/i.test(cmd)) {
      exportData("civil3d_csv");
      return;
    }

    showExcelLog(`أدخل أمراً مثل: "صنف الأكواد تلقائياً" أو "افحص أخطاء الرفع" أو "اجمع 0.50 على Z" أو "احسب المساحة".`, "info");
  }

  function getFilteredRowsWithIndex() {
    const q = searchFilter.trim().toLowerCase();
    return rows
      .map((r, originalIdx) => ({ r, originalIdx, cls: getRowClassification(r) }))
      .filter(({ r, cls }) => {
        if (activeCodeFilter !== "ALL" && cls.group !== activeCodeFilter) return false;
        if (!q) return true;
        return columns.some(c => String(r[c] ?? "").toLowerCase().includes(q)) || cls.labelAr.toLowerCase().includes(q);
      });
  }

  function renderExcelTable() {
    const table = document.getElementById("gdExcelDataTable");
    const statsBar = document.getElementById("gdExStatsBar");
    const filterBar = document.getElementById("gdExCodeFilterBar");
    if (!table) return;

    const st = computeSurveyStats();
    const hasComp = asBuiltResults.length > 0;
    const failedCount = asBuiltResults.filter(c => !c.isPass).length;

    if (statsBar) {
      statsBar.innerHTML = `
        <div class="gd-ex-stat"><span>إجمالي النقاط</span><strong>${st.count}</strong></div>
        <div class="gd-ex-stat"><span>أقل منسوب (Z)</span><strong style="color:#38bdf8;">${st.minZ}</strong></div>
        <div class="gd-ex-stat"><span>أعلى منسوب (Z)</span><strong style="color:#f59e0b;">${st.maxZ}</strong></div>
        <div class="gd-ex-stat"><span>مطابق MNH صافي</span><strong style="color:#fb923c;">${st.sewerCount}</strong></div>
        <div class="gd-ex-stat"><span>أرض طبيعية (EG)</span><strong style="color:#10b981;">${st.nglCount}</strong></div>
        <div class="gd-ex-stat"><span>المساحة المغلقة</span><strong>${st.areaM2.toFixed(1)} م²</strong></div>
        ${hasComp ? `<div class="gd-ex-stat" style="border-color:${failedCount ? "#ef4444" : "#22c55e"};"><span>تفاوت As-Built</span><strong style="color:${failedCount ? "#ef4444" : "#22c55e"};">${failedCount ? failedCount + " تجاوز" : "مطابق"}</strong></div>` : ""}
      `;
    }

    if (filterBar) {
      const filters = [
        { id: "ALL", label: `الكل (${st.count})` },
        { id: "SEWER", label: `المطابق والصرف MNH (${st.sewerCount})` },
        { id: "NGL", label: `الأرض الطبيعية EG (${st.nglCount})` },
        { id: "LIGHT", label: `أعمدة الإنارة (${st.lightCount})` },
        { id: "TREES", label: `الأشجار (${st.treeCount})` },
        { id: "ROAD", label: `طرق وأرصفة (${st.roadCount})` },
        { id: "DIR", label: `شمال / جنوب (${st.dirCount})` }
      ];
      filterBar.innerHTML = filters
        .map(f => `<button type="button" class="gd-ex-filter-pill ${activeCodeFilter === f.id ? "active" : ""}" data-code-filter="${f.id}">${f.label}</button>`)
        .join("");
    }

    const filtered = getFilteredRowsWithIndex();
    const { zKey } = getCoordKeys();
    const allZ = rows.map(r => parseFloat(r[zKey])).filter(v => !isNaN(v));
    const minZ = allZ.length ? Math.min(...allZ) : 0;
    const maxZ = allZ.length ? Math.max(...allZ) : 0;

    table.innerHTML = `
      <thead>
        <tr>
          <th>#</th>
          <th>التصنيف الهندسي</th>
          ${columns.map(col => `<th>${col}</th>`).join("")}
          ${hasComp ? `<th>إزاحة XY (مم)</th>` : ""}
          <th>حذف</th>
        </tr>
      </thead>
      <tbody>
        ${filtered.map(({ r, originalIdx, cls }) => {
          const comp = hasComp ? asBuiltResults.find(c => c.rowIdx === originalIdx) : null;
          const zVal = parseFloat(r[zKey]);
          const zBadgeColor = !isNaN(zVal) ? getElevationColor(zVal, minZ, maxZ) : "transparent";

          return `
            <tr>
              <td style="background:var(--ex-th-bg);color:var(--ex-muted);font-weight:800;padding:6px 8px;">${originalIdx + 1}</td>
              <td style="padding:5px 8px;white-space:nowrap;font-size:11px;font-weight:900;color:${cls.color};">
                ${cls.labelAr}
              </td>
              ${columns.map(col => {
                const isZCol = col === zKey;
                const cellExtraStyle = isZCol && !isNaN(zVal)
                  ? `border-right:4px solid ${zBadgeColor};`
                  : "";
                return `
                  <td>
                    <input type="text" class="gd-ex-cell" style="${cellExtraStyle}"
                      data-row="${originalIdx}" data-col="${col}"
                      value="${String(r[col] ?? "").replace(/"/g, "&quot;")}">
                  </td>
                `;
              }).join("")}
              ${hasComp ? `
                <td style="padding:5px 8px;white-space:nowrap;font-size:11.5px;font-weight:900;color:${comp && comp.isPass ? "#22c55e" : "#ef4444"};">
                  ${comp ? `${comp.dXY_mm.toFixed(1)} مم` : "-"}
                </td>
              ` : ""}
              <td>
                <button type="button" style="background:transparent;border:0;color:#ef4444;cursor:pointer;font-weight:900;padding:6px 9px;"
                  data-del-row="${originalIdx}">✕</button>
              </td>
            </tr>
          `;
        }).join("")}
      </tbody>
    `;
  }

  /* ============================================================================
     تحديث مفتاح ألوان المناسيب بجانب الرسمة من الأقل للأكبر
     ============================================================================ */
  function updateElevationLegend(minZ, maxZ) {
    const ticksBox = document.getElementById("gdElevationLegendTicks");
    const deltaBox = document.getElementById("gdElevationDeltaBadge");
    if (!ticksBox || !deltaBox) return;

    if (!isFinite(minZ) || !isFinite(maxZ)) {
      ticksBox.innerHTML = `<span class="gd-elev-tick-max">--</span><span class="gd-elev-tick-min">--</span>`;
      deltaBox.textContent = "ΔZ: --";
      return;
    }

    const span = maxZ - minZ;
    const v4 = maxZ;
    const v3 = minZ + span * 0.75;
    const v2 = minZ + span * 0.50;
    const v1 = minZ + span * 0.25;
    const v0 = minZ;

    ticksBox.innerHTML = `
      <span class="gd-elev-tick-max" title="أعلى منسوب (بني غامق)">${v4.toFixed(2)}</span>
      <span class="gd-elev-tick-mid">${v3.toFixed(2)}</span>
      <span class="gd-elev-tick-mid">${v2.toFixed(2)}</span>
      <span class="gd-elev-tick-mid">${v1.toFixed(2)}</span>
      <span class="gd-elev-tick-min" title="أقل منسوب (أزرق غامق)">${v0.toFixed(2)}</span>
    `;
    deltaBox.textContent = `ΔZ: ${span.toFixed(2)}m`;
  }

  function drawCadPreview() {
    const canvas = document.getElementById("gdExcelCadCanvas");
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    const rect = canvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    const W = rect.width || 320;
    const H = rect.height || 320;

    canvas.width = Math.max(300, W * dpr);
    canvas.height = Math.max(250, H * dpr);
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    ctx.fillStyle = "#080c12";
    ctx.fillRect(0, 0, W, H);
    ctx.strokeStyle = "rgba(148, 163, 184, 0.06)";
    ctx.lineWidth = 1;
    for (let x = 0; x < W; x += 28) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, H); ctx.stroke(); }
    for (let y = 0; y < H; y += 28) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(W, y); ctx.stroke(); }

    const { pKey, eKey, nKey, zKey, cKey } = getCoordKeys();
    const filteredList = getFilteredRowsWithIndex();
    const pts = filteredList
      .map(({ r, cls }) => ({
        p: String(r[pKey] || ""),
        c: String(r[cKey] || ""),
        e: parseFloat(r[eKey]),
        n: parseFloat(r[nKey]),
        z: parseFloat(r[zKey]),
        cls
      }))
      .filter(pt => !isNaN(pt.e) && !isNaN(pt.n));

    if (!pts.length) {
      updateElevationLegend(NaN, NaN);
      ctx.fillStyle = "#64748b";
      ctx.font = "bold 13px sans-serif";
      ctx.textAlign = "center";
      ctx.fillText("لا توجد نقاط مطابقة للعرض حالياً", W / 2, H / 2);
      return;
    }

    const zValues = pts.map(p => p.z).filter(v => !isNaN(v));
    const minZ = zValues.length ? Math.min(...zValues) : 0;
    const maxZ = zValues.length ? Math.max(...zValues) : 0;
    updateElevationLegend(minZ, maxZ);

    const allE = [...pts.map(p => p.e), ...designRows.map(d => d.E)];
    const allN = [...pts.map(p => p.n), ...designRows.map(d => d.N)];
    const minE = Math.min(...allE);
    const maxE = Math.max(...allE);
    const minN = Math.min(...allN);
    const maxN = Math.max(...allN);
    const spanE = Math.max(maxE - minE, 1);
    const spanN = Math.max(maxN - minN, 1);
    const pad = 36;
    const baseScale = Math.min((W - pad * 2) / spanE, (H - pad * 2) / spanN);
    const scale = baseScale * cadZoom;

    const centerX = W / 2 + cadPanX;
    const centerY = H / 2 + cadPanY;
    const midE = (minE + maxE) / 2;
    const midN = (minN + maxN) / 2;

    const toCanvas = (e, n) => ({
      x: centerX + (e - midE) * scale,
      y: centerY - (n - midN) * scale
    });

    // 1. تظليل كنتوري خافت جداً للمناسيب (بدون توهج فاقع)
    if (elevationColorMode && zValues.length > 0) {
      pts.forEach(pt => {
        if (isNaN(pt.z)) return;
        const c = toCanvas(pt.e, pt.n);
        const col = getElevationColor(pt.z, minZ, maxZ);
        const radius = Math.max(12, Math.min(26, 16 * Math.sqrt(cadZoom)));
        const grad = ctx.createRadialGradient(c.x, c.y, 1, c.x, c.y, radius);
        grad.addColorStop(0, col.replace("rgb(", "rgba(").replace(")", ", 0.18)"));
        grad.addColorStop(1, col.replace("rgb(", "rgba(").replace(")", ", 0.0)"));
        ctx.fillStyle = grad;
        ctx.beginPath();
        ctx.arc(c.x, c.y, radius, 0, Math.PI * 2);
        ctx.fill();
      });

      // رسم خطوط كنتورية خافتة بين النقاط المتقاربة في المنسوب
      if (pts.length >= 3 && maxZ > minZ) {
        const levels = 6;
        for (let lv = 1; lv < levels; lv++) {
          const targetZ = minZ + ((maxZ - minZ) * lv) / levels;
          const bandPts = pts.filter(p => !isNaN(p.z) && Math.abs(p.z - targetZ) <= (maxZ - minZ) / (levels * 1.4));
          if (bandPts.length >= 2) {
            const col = getElevationColor(targetZ, minZ, maxZ).replace("rgb(", "rgba(").replace(")", ", 0.55)");
            ctx.strokeStyle = col;
            ctx.lineWidth = 1.2;
            ctx.beginPath();
            bandPts.forEach((bp, i) => {
              const sc = toCanvas(bp.e, bp.n);
              if (i === 0) ctx.moveTo(sc.x, sc.y);
              else ctx.lineTo(sc.x, sc.y);
            });
            ctx.stroke();
          }
        }
      }
    }

    // 2. رسم نقاط التصميم وأسهم الترحيل (As-Built vs Design)
    if (asBuiltResults.length) {
      asBuiltResults.forEach(c => {
        const pDes = toCanvas(c.desE, c.desN);
        const pAct = toCanvas(c.actE, c.actN);

        ctx.strokeStyle = "rgba(56, 189, 248, 0.75)";
        ctx.lineWidth = 1.3;
        ctx.strokeRect(pDes.x - 4, pDes.y - 4, 8, 8);

        ctx.strokeStyle = c.isPass ? "rgba(34, 197, 94, 0.85)" : "rgba(239, 68, 68, 0.85)";
        ctx.lineWidth = 1.4;
        ctx.setLineDash([4, 3]);
        ctx.beginPath();
        ctx.moveTo(pDes.x, pDes.y);
        ctx.lineTo(pAct.x, pAct.y);
        ctx.stroke();
        ctx.setLineDash([]);
      });
    }

    // 3. خطوط الربط الهندسية
    if (connectMode === "SEQUENTIAL" && pts.length >= 2) {
      for (let i = 0; i < pts.length - 1; i++) {
        const c1 = toCanvas(pts[i].e, pts[i].n);
        const c2 = toCanvas(pts[i + 1].e, pts[i + 1].n);
        ctx.beginPath();
        ctx.strokeStyle = elevationColorMode
          ? getElevationColor(pts[i].z, minZ, maxZ).replace("rgb(", "rgba(").replace(")", ", 0.65)")
          : "rgba(56, 189, 248, 0.65)";
        ctx.lineWidth = 1.2;
        ctx.moveTo(c1.x, c1.y);
        ctx.lineTo(c2.x, c2.y);
        ctx.stroke();
      }
    } else if (connectMode === "BY_CODE") {
      const groups = {};
      pts.forEach(pt => {
        if (pt.cls.group === "LIGHT" || pt.cls.group === "TREES" || pt.cls.group === "NGL") return;
        if (!groups[pt.cls.layer]) groups[pt.cls.layer] = [];
        groups[pt.cls.layer].push(pt);
      });
      Object.values(groups).forEach(gPts => {
        if (gPts.length < 2) return;
        for (let i = 0; i < gPts.length - 1; i++) {
          const c1 = toCanvas(gPts[i].e, gPts[i].n);
          const c2 = toCanvas(gPts[i + 1].e, gPts[i + 1].n);
          ctx.beginPath();
          ctx.strokeStyle = elevationColorMode
            ? getElevationColor(gPts[i].z, minZ, maxZ).replace("rgb(", "rgba(").replace(")", ", 0.65)")
            : gPts[0].cls.color;
          ctx.lineWidth = 1.3;
          ctx.moveTo(c1.x, c1.y);
          ctx.lineTo(c2.x, c2.y);
          ctx.stroke();
        }
      });
    }

    const visiblePts = [];
    pts.forEach(pt => {
      const c = toCanvas(pt.e, pt.n);
      if (c.x >= -20 && c.x <= W + 20 && c.y >= -20 && c.y <= H + 20) {
        visiblePts.push({ ...pt, cx: c.x, cy: c.y });
      }
    });

    const allowText = showLabelsOnCanvas && (visiblePts.length <= 80 || cadZoom >= 2.2);

    // 4. رسم رموز النقاط بألوان هادئة واضحة
    visiblePts.forEach(pt => {
      const baseCol = elevationColorMode && !isNaN(pt.z)
        ? getElevationColor(pt.z, minZ, maxZ)
        : pt.cls.color;

      ctx.strokeStyle = baseCol;
      ctx.fillStyle = baseCol;
      ctx.lineWidth = 1.5;

      if (pt.cls.symbol === "MANHOLE") {
        ctx.beginPath();
        ctx.arc(pt.cx, pt.cy, 5, 0, Math.PI * 2);
        ctx.stroke();
        ctx.beginPath();
        ctx.arc(pt.cx, pt.cy, 2, 0, Math.PI * 2);
        ctx.fill();
      } else if (pt.cls.symbol === "LIGHT") {
        ctx.beginPath();
        ctx.arc(pt.cx, pt.cy, 3.5, 0, Math.PI * 2);
        ctx.fill();
        ctx.beginPath();
        ctx.moveTo(pt.cx - 6, pt.cy); ctx.lineTo(pt.cx + 6, pt.cy);
        ctx.moveTo(pt.cx, pt.cy - 6); ctx.lineTo(pt.cx, pt.cy + 6);
        ctx.stroke();
      } else if (pt.cls.symbol === "TREE") {
        ctx.beginPath();
        ctx.arc(pt.cx, pt.cy, 4.5, 0, Math.PI * 2);
        ctx.stroke();
      } else if (pt.cls.symbol === "TRIANGLE") {
        ctx.beginPath();
        ctx.moveTo(pt.cx, pt.cy - 5.5);
        ctx.lineTo(pt.cx - 5, pt.cy + 4);
        ctx.lineTo(pt.cx + 5, pt.cy + 4);
        ctx.closePath();
        ctx.stroke();
      } else {
        // علامة (×) المساحية الهادئة لنقاط الأرض الطبيعية
        ctx.beginPath();
        ctx.moveTo(pt.cx - 3.5, pt.cy - 3.5); ctx.lineTo(pt.cx + 3.5, pt.cy + 3.5);
        ctx.moveTo(pt.cx + 3.5, pt.cy - 3.5); ctx.lineTo(pt.cx - 3.5, pt.cy + 3.5);
        ctx.stroke();
      }

      if (allowText) {
        ctx.fillStyle = "rgba(226, 232, 240, 0.85)";
        ctx.font = "600 10px sans-serif";
        ctx.textAlign = "left";
        const zTxt = !isNaN(pt.z) ? ` ${pt.z.toFixed(2)}` : "";
        const cTxt = pt.c ? `[${pt.c}]` : "";
        ctx.fillText(`${pt.p} ${cTxt}${zTxt}`, pt.cx + 6, pt.cy - 3);
      }
    });

    ctx.fillStyle = "rgba(15, 23, 42, 0.78)";
    ctx.fillRect(W - 115, 8, 107, 22);
    ctx.fillStyle = "#94a3b8";
    ctx.font = "bold 11px sans-serif";
    ctx.textAlign = "center";
    ctx.fillText(`Zoom: ${cadZoom.toFixed(1)}x`, W - 61, 23);
  }

  function bindCanvasInteractions() {
    const wrapper = document.getElementById("gdCadCanvasWrapper");
    if (!wrapper || wrapper.dataset.bound) return;
    wrapper.dataset.bound = "1";

    wrapper.addEventListener("wheel", e => {
      e.preventDefault();
      const factor = e.deltaY < 0 ? 1.25 : 0.8;
      cadZoom = Math.min(80, Math.max(0.3, cadZoom * factor));
      drawCadPreview();
    }, { passive: false });

    wrapper.addEventListener("mousedown", e => {
      if (e.target.closest(".gd-cad-zbtn")) return;
      isDraggingCanvas = true;
      dragStartX = e.clientX - cadPanX;
      dragStartY = e.clientY - cadPanY;
    });

    window.addEventListener("mousemove", e => {
      if (!isDraggingCanvas) return;
      cadPanX = e.clientX - dragStartX;
      cadPanY = e.clientY - dragStartY;
      drawCadPreview();
    });

    window.addEventListener("mouseup", () => { isDraggingCanvas = false; });

    wrapper.addEventListener("touchstart", e => {
      if (e.target.closest(".gd-cad-zbtn")) return;
      if (e.touches.length === 1) {
        isDraggingCanvas = true;
        dragStartX = e.touches[0].clientX - cadPanX;
        dragStartY = e.touches[0].clientY - cadPanY;
      } else if (e.touches.length === 2) {
        isDraggingCanvas = false;
        initialPinchDist = Math.hypot(
          e.touches[0].clientX - e.touches[1].clientX,
          e.touches[0].clientY - e.touches[1].clientY
        );
        initialPinchZoom = cadZoom;
      }
    }, { passive: true });

    wrapper.addEventListener("touchmove", e => {
      if (e.target.closest(".gd-cad-zbtn")) return;
      if (e.touches.length === 1 && isDraggingCanvas) {
        e.preventDefault();
        cadPanX = e.touches[0].clientX - dragStartX;
        cadPanY = e.touches[0].clientY - dragStartY;
        drawCadPreview();
      } else if (e.touches.length === 2 && initialPinchDist > 0) {
        e.preventDefault();
        const curDist = Math.hypot(
          e.touches[0].clientX - e.touches[1].clientX,
          e.touches[0].clientY - e.touches[1].clientY
        );
        cadZoom = Math.min(80, Math.max(0.3, initialPinchZoom * (curDist / initialPinchDist)));
        drawCadPreview();
      }
    }, { passive: false });

    wrapper.addEventListener("touchend", () => {
      isDraggingCanvas = false;
      initialPinchDist = 0;
    });
  }

  function setupMainAssistantRouter() {
    document.addEventListener("click", e => {
      const sendBtn = e.target.closest("#sendBtn, .send-btn, button[type='submit']");
      if (!sendBtn) return;

      const chatInput =
        document.getElementById("userInput") ||
        document.getElementById("chatInput") ||
        document.querySelector("#s-chat textarea, #s-chat input[type='text'], #s-assistant textarea");

      if (!chatInput || e.target.closest("#gdAiExcelStudio")) return;
      const q = String(chatInput.value || "").trim();
      if (!q) return;

      if (/(اكسيل|إكسيل|excel|penz|penzd|civil\s*3d|sdr|str|توتال|اعمدة\s*انارة|أعمدة\s*إنارة|صرف\s*صحي|mnh|تحويل\s*نقاط|مناسيب)/i.test(q)) {
        setTimeout(() => {
          const chatBox =
            document.getElementById("chatBox") ||
            document.getElementById("messages") ||
            document.querySelector("#s-chat .msgs, #s-chat .chat-body");
          if (!chatBox || document.getElementById("gdAssistantExcelRedirectCard")) return;

          const card = document.createElement("div");
          card.id = "gdAssistantExcelRedirectCard";
          card.style.cssText = `
            direction: rtl; text-align: right; margin: 10px 0; padding: 13px 15px;
            border-radius: 14px; border: 1.5px solid #e09f24;
            background: linear-gradient(135deg, #1c1917, #292524); color: #f5f5f4;
          `;
          card.innerHTML = `
            <strong style="color:#fbbf24;font-size:13.5px;display:block;margin-bottom:4px;">
              توجيه من مساعد GeoDev الهندسي
            </strong>
            <p style="margin:0 0 10px;font-size:12.5px;color:#e7e5e4;line-height:1.65;">
              لإنشاء أو تعديل ملفات الإكسيل المساحية (<b>PENZD / PENZ</b>) ومعاينة تدرج المناسيب الملون ومقارنة التصميم بالرفع والتصدير لـ <b>Civil 3D</b> و <b>AutoCAD</b>، يرجى التوجه إلى <b>قسم المساحة — محطة إكسيل المساحة</b>:
            </p>
            <button type="button" id="gdJumpToSurvExcelBtn"
              style="padding:9px 15px;border-radius:9px;border:0;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-weight:900;font-size:12.5px;cursor:pointer;font-family:inherit;">
              الانتقال إلى قسم المساحة ومحطة الإكسيل الآن
            </button>
          `;
          chatBox.appendChild(card);
          chatBox.scrollTop = chatBox.scrollHeight;
        }, 450);
      }
    }, true);

    document.addEventListener("click", e => {
      if (e.target.closest("#gdJumpToSurvExcelBtn")) {
        const navButtons = Array.from(document.querySelectorAll("nav button, [data-t]"));
        const survBtn = navButtons.find(btn => (btn.textContent || "").trim() === "المساحة");
        if (survBtn) survBtn.click();
        setTimeout(() => {
          const studio = document.getElementById("gdAiExcelStudio");
          if (studio) studio.scrollIntoView({ behavior: "smooth", block: "start" });
        }, 250);
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

  function findGuestLandingHost() {
    return (
      document.getElementById("authScreen") ||
      document.getElementById("loginScreen") ||
      document.getElementById("authOverlay") ||
      document.querySelector(".auth-wrap, .login-wrap") ||
      document.body
    );
  }

  function mountAIExcelStudio() {
    const logged = checkAuthState();

    const oldGuestBanner = document.getElementById("gdGuestWelcomeBanner");
    if (logged && oldGuestBanner) oldGuestBanner.remove();

    if (logged) {
      const targetSec = findSurveyingSection();
      if (!targetSec) return;

      let container = document.getElementById("gdAiExcelStudio");
      if (!container) {
        container = document.createElement("div");
        container.id = "gdAiExcelStudio";
        container.dataset.theme = currentTheme;
      }
      if (container.parentElement !== targetSec) {
        container.innerHTML = renderStudioHTML();
        targetSec.prepend(container);
        renderExcelTable();
        bindCanvasInteractions();
        setTimeout(drawCadPreview, 80);
      }
      return;
    }

    const guestHost = findGuestLandingHost();
    if (!guestHost) return;

    if (!document.getElementById("gdGuestWelcomeBanner")) {
      const banner = document.createElement("div");
      banner.id = "gdGuestWelcomeBanner";
      banner.innerHTML = `
        <div>
          <strong>أهلاً بك في منصة GeoDev — جرّب محطة إكسيل المساحة مجاناً قبل تسجيل الدخول</strong>
          <span>ارفع ملفات (Excel / DXF / CSV / SDR / STR) وعاين النقاط على لوحة الرسم الهندسية فوراً. سجّل الدخول لتفعيل الحفظ والتصدير.</span>
        </div>
        <button type="button" id="gdBannerScrollLoginBtn" class="gd-ex-btn gold">
          تسجيل الدخول / إنشاء حساب
        </button>
      `;
      if (guestHost === document.body) document.body.prepend(banner);
      else guestHost.appendChild(banner);
    }

    let container = document.getElementById("gdAiExcelStudio");
    if (!container) {
      container = document.createElement("div");
      container.id = "gdAiExcelStudio";
      container.dataset.theme = currentTheme;
      container.innerHTML = renderStudioHTML();
      const bannerEl = document.getElementById("gdGuestWelcomeBanner");
      if (bannerEl && bannerEl.parentElement) bannerEl.insertAdjacentElement("afterend", container);
      else guestHost.appendChild(container);
      renderExcelTable();
      bindCanvasInteractions();
      setTimeout(drawCadPreview, 80);
    }
  }

  function bindEvents() {
    document.addEventListener("click", e => {
      // زر تبديل تلوين المناسيب المتدرج
      const elevBtn = e.target.closest("#gdToggleElevColorBtn");
      if (elevBtn) {
        elevationColorMode = !elevationColorMode;
        elevBtn.textContent = `تدرج المناسيب: ${elevationColorMode ? "مفعّل" : "أكواد"}`;
        elevBtn.style.borderColor = elevationColorMode ? "#fbbf24" : "";
        elevBtn.style.color = elevationColorMode ? "#fbbf24" : "";
        const legend = document.getElementById("gdElevationLegendPanel");
        if (legend) legend.style.opacity = elevationColorMode ? "1" : "0.45";
        drawCadPreview();
        return;
      }

      // فتح نافذة تصدير التقرير الرسمي
      if (e.target.closest("#gdOpenOfficialReportBtn") || e.target.closest("#gdExportReportHubBtn")) {
        openOfficialReportModal();
        return;
      }

      // تشغيل حساب مقارنة التصميم بالرفع الجديد
      if (e.target.closest("#gdRunAsBuiltCalcBtn")) {
        if (!designRows.length) {
          showExcelLog("ارفع ملف نقاط التصميم (القديمة) أولاً للمقارنة.", "warn");
          return;
        }
        const tol = parseFloat(document.getElementById("gdAsBuiltTolInput")?.value) || 20;
        const res = computeAsBuiltComparison(tol);
        const passed = res.filter(r => r.isPass).length;
        const failed = res.filter(r => !r.isPass).length;
        const sumBox = document.getElementById("gdAsBuiltSummaryBox");
        if (sumBox) {
          sumBox.innerHTML = `<strong style="color:#fbbf24;">نتيجة مقارنة ${res.length} نقطة:</strong> <span style="color:#4ade80;font-weight:900;">مقبول: ${passed}</span> | <span style="color:#f87171;font-weight:900;">تجاوز السماحية (±${tol} مم): ${failed}</span>`;
        }
        renderExcelTable();
        drawCadPreview();
        showExcelLog(`تم حساب التفاوتات لـ ${res.length} نقطة (${passed} مقبول / ${failed} تجاوز السماحية).`, "success");
        return;
      }

      if (e.target.closest("#gdClearDesignBtn")) {
        designRows = [];
        asBuiltResults = [];
        const sumBox = document.getElementById("gdAsBuiltSummaryBox");
        if (sumBox) sumBox.textContent = "تم مسح نقاط التصميم.";
        renderExcelTable();
        drawCadPreview();
        return;
      }

      if (e.target.closest("#gdCycleThemeBulbBtn")) {
        const order = ["dark_gold", "light_eng", "cad_blue"];
        const nextIdx = (order.indexOf(currentTheme) + 1) % order.length;
        currentTheme = order[nextIdx];
        localStorage.setItem("gd_excel_theme_v6", currentTheme);
        const studio = document.getElementById("gdAiExcelStudio");
        if (studio) studio.dataset.theme = currentTheme;
        const lbl = document.getElementById("gdThemeBulbLabel");
        if (lbl) lbl.textContent = getThemeLabel(currentTheme);
        return;
      }

      if (e.target.closest("#gdBannerScrollLoginBtn")) {
        const authForm = document.querySelector('#authScreen input, #loginScreen input, form input[type="email"], form input[type="password"]');
        if (authForm) {
          authForm.scrollIntoView({ behavior: "smooth", block: "center" });
          authForm.focus();
        }
        return;
      }

      if (e.target.closest("#gdExcelAiExecBtn")) {
        const inp = document.getElementById("gdExcelAiPrompt");
        if (inp) executeAICommand(inp.value);
        return;
      }

      if (e.target.closest("#gdExAutoClassifyBtn")) {
        if (!checkAuthState()) return openAuthRequiredModal("تصنيف الأكواد وإضافة أعمدة الطبقات");
        runAutoCodeClassification();
        return;
      }

      const filterBtn = e.target.closest("[data-code-filter]");
      if (filterBtn) {
        activeCodeFilter = filterBtn.dataset.codeFilter;
        const scopeSel = document.getElementById("gdExportTargetScopeSelect");
        if (scopeSel) {
          if (activeCodeFilter === "LIGHT") scopeSel.value = "LIGHT";
          else if (activeCodeFilter === "SEWER") scopeSel.value = "SEWER";
          else if (activeCodeFilter === "NGL") scopeSel.value = "NGL";
          else if (activeCodeFilter === "TREES") scopeSel.value = "TREES";
          else if (activeCodeFilter === "ROAD") scopeSel.value = "ROAD";
          else if (activeCodeFilter === "ALL") scopeSel.value = "ALL";
          else scopeSel.value = "FILTERED";
        }
        renderExcelTable();
        drawCadPreview();
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

      if (e.target.closest("#gdCadZoomInBtn")) { cadZoom = Math.min(80, cadZoom * 1.45); drawCadPreview(); return; }
      if (e.target.closest("#gdCadZoomOutBtn")) { cadZoom = Math.max(0.3, cadZoom / 1.45); drawCadPreview(); return; }
      if (e.target.closest("#gdCadZoomResetBtn")) { resetCadView(); return; }

      const subTab = e.target.closest("[data-ex-subtab]");
      if (subTab) {
        activeSubTab = subTab.dataset.exSubtab;
        document.querySelectorAll("[data-ex-subtab]").forEach(b => b.classList.toggle("active", b.dataset.exSubtab === activeSubTab));
        const pAsBuilt = document.getElementById("gdExSubpanelAsBuilt");
        const pCad = document.getElementById("gdExSubpanelCadImport");
        const pCalc = document.getElementById("gdExSubpanelCalcTools");
        if (pAsBuilt) pAsBuilt.style.display = activeSubTab === "asbuilt_compare" ? "block" : "none";
        if (pCad) pCad.style.display = activeSubTab === "cad_import" ? "block" : "none";
        if (pCalc) pCalc.style.display = activeSubTab === "calc_tools" ? "block" : "none";
        return;
      }

      if (e.target.closest("#gdParseAcadListBtn")) {
        const txt = document.getElementById("gdAcadListPasteArea")?.value || "";
        const extracted = parseAutoCADListText(txt);
        if (!extracted.length) return showExcelLog("لم يتم العثور على إحداثيات X= ... Y= ... في النص الملصق.", "error");
        saveSnapshot();
        rows.push(...extracted);
        renderExcelTable();
        resetCadView();
        showExcelLog(`تم استخراج ${extracted.length} نقطة وعرضها في لوحة الكاد.`, "success");
        return;
      }

      if (e.target.closest("#gdApplyShiftBtn")) {
        if (!checkAuthState()) return openAuthRequiredModal("تطبيق إزاحة الإحداثيات");
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
        if (designRows.length) computeAsBuiltComparison();
        renderExcelTable();
        drawCadPreview();
        showExcelLog(`تم تطبيق إزاحة الإحداثيات (ΔE=${dE}, ΔN=${dN}, ΔZ=${dZ}).`, "success");
        return;
      }

      if (e.target.closest("#gdCalcCutFillBtn")) {
        if (!checkAuthState()) return openAuthRequiredModal("حساب كميات الحفر والردم");
        const designLvl = parseFloat(document.getElementById("gdDesignLevelInput")?.value);
        if (isNaN(designLvl)) return showExcelLog("أدخل منسوب التصميم أولاً.", "error");
        saveSnapshot();
        const { zKey } = getCoordKeys();
        if (!columns.includes("Cut_Fill")) columns.push("Cut_Fill");
        rows.forEach(r => {
          const z = parseFloat(r[zKey]);
          if (!isNaN(z)) {
            const diff = z - designLvl;
            r.Cut_Fill = diff > 0 ? `Cut +${diff.toFixed(3)}` : diff < 0 ? `Fill ${diff.toFixed(3)}` : "0.000";
          }
        });
        renderExcelTable();
        showExcelLog(`تم حساب الحفر والردم مقارنة بمنسوب ${designLvl} م.`, "success");
        return;
      }

      if (e.target.closest("#gdExSwapENBtn")) {
        saveSnapshot();
        const { eKey, nKey } = getCoordKeys();
        rows.forEach(r => {
          const tmp = r[eKey];
          r[eKey] = r[nKey];
          r[nKey] = tmp;
        });
        if (designRows.length) computeAsBuiltComparison();
        renderExcelTable();
        resetCadView();
        showExcelLog("تم تبديل إحداثيات الشرق (E) والشمال (N).", "success");
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

      if (e.target.closest("#gdExUndoBtn")) { undoLastAction(); return; }

      if (e.target.closest("#gdExClearBtn")) {
        if (!confirm("هل تريد تفريغ جميع النقاط من الجدول؟")) return;
        saveSnapshot();
        rows = [];
        asBuiltResults = [];
        renderExcelTable();
        resetCadView();
        return;
      }

      const delRowBtn = e.target.closest("[data-del-row]");
      if (delRowBtn) {
        saveSnapshot();
        rows.splice(parseInt(delRowBtn.dataset.delRow, 10), 1);
        if (designRows.length) computeAsBuiltComparison();
        renderExcelTable();
        drawCadPreview();
        return;
      }

      if (e.target.closest("#gdToggleLabelsBtn")) { showLabelsOnCanvas = !showLabelsOnCanvas; drawCadPreview(); return; }

      const cycleBtn = e.target.closest("#gdCycleConnectBtn");
      if (cycleBtn) {
        if (connectMode === "BY_CODE") {
          connectMode = "SEQUENTIAL";
          cycleBtn.textContent = "التوصيل: متسلسل";
        } else if (connectMode === "SEQUENTIAL") {
          connectMode = "NONE";
          cycleBtn.textContent = "التوصيل: بدون خطوط";
        } else {
          connectMode = "BY_CODE";
          cycleBtn.textContent = "التوصيل: حسب الكود";
        }
        drawCadPreview();
        return;
      }

      const expBtn = e.target.closest("[data-export-fmt]");
      if (expBtn) exportData(expBtn.dataset.exportFmt);
    });

    document.addEventListener("input", e => {
      if (e.target.classList.contains("gd-ex-cell")) {
        const rIdx = parseInt(e.target.dataset.row, 10);
        const col = e.target.dataset.col;
        if (rows[rIdx]) {
          rows[rIdx][col] = e.target.value;
          if (designRows.length) computeAsBuiltComparison();
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
      if (e.target.id === "gdDesignFileInput" && e.target.files?.[0]) {
        handleDesignFileUpload(e.target.files[0]);
        e.target.value = "";
        return;
      }
      if (e.target.id === "gdCrsSelect") {
        crsMode = e.target.value;
        showExcelLog(`تم ضبط نظام الإحداثيات الجغرافي إلى: ${crsMode}`, "info");
        return;
      }
      if (e.target.id === "gdExportColOrderSelect") {
        exportColOrder = e.target.value;
        return;
      }
      if (e.target.id === "gdExportTargetScopeSelect") {
        exportTargetScope = e.target.value;
      }
    });

    document.addEventListener("keydown", e => {
      if (e.target.id === "gdExcelAiPrompt" && e.key === "Enter") {
        e.preventDefault();
        executeAICommand(e.target.value);
      }
    });
  }

  function startAuthWatcher() {
    const attach = () => {
      if (window.firebase && firebase.apps && firebase.apps.length) {
        firebase.auth().onAuthStateChanged(user => {
          isLoggedInUser = Boolean(user);
          const oldStudio = document.getElementById("gdAiExcelStudio");
          if (oldStudio) oldStudio.remove();
          mountAIExcelStudio();
        });
        return true;
      }
      return false;
    };

    if (!attach()) {
      const iv = setInterval(() => {
        if (attach()) clearInterval(iv);
      }, 400);
      setTimeout(() => clearInterval(iv), 10000);
    }
  }

  function init() {
    ensureSheetJS();
    injectStyles();
    bindEvents();
    setupMainAssistantRouter();
    mountAIExcelStudio();
    startAuthWatcher();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
