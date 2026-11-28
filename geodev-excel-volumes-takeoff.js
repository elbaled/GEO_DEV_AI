/* ============================================================================
   GeoDev Survey & GIS Excel — Surface-to-Surface Cut/Fill & Batch Area/Length Take-off (V1.0)
   - حساب كميات الحفر والردم (m³) بمجرد وضع السطحين (سطح الأرض الطبيعية NGL + سطح التصميم Design)
   - حصر المساحات المغلقة وتجميع الأطوال دفعة واحدة (م.ط / م² / فدان-قيراط-سهم) لكل كود/مضلع
   - تصدير جدول الحصر والكميات مباشرة إلى Excel (.xlsx) متوافق مع المتصفح وتطبيق الأندرويد
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevVolumesTakeoffV1Init) return;
  window.GeoDevVolumesTakeoffV1Init = true;

  let surface1NGL = [];    // السطح الأول: الأرض الطبيعية
  let surface2Design = []; // السطح الثاني: سطح التصميم
  let lastVolumeReport = null;
  let lastTakeoffReport = null;

  function ensureXLSX(cb) {
    if (window.XLSX) return cb();
    const s = document.createElement("script");
    s.src = "https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js";
    s.onload = () => cb();
    document.head.appendChild(s);
  }

  /* ============================================================================
     1. استخراج النقاط من الجدول الحالي في محطة الإكسيل أو من ملف مرفوع
     ============================================================================ */
  function readCurrentExcelTablePoints() {
    const table = document.getElementById("gdExcelDataTable");
    if (!table) return [];

    const headers = Array.from(table.querySelectorAll("thead th")).map(th => th.textContent.trim());
    const trs = Array.from(table.querySelectorAll("tbody tr"));
    const pts = [];

    // تحديد أعمدة P, E, N, Z, Code ديناميكياً
    let idxP = headers.findIndex(h => /^(p|pt|point|id|no|رقم|النقطة)$/i.test(h));
    let idxE = headers.findIndex(h => /^(e|east|easting|x|شرق)$/i.test(h));
    let idxN = headers.findIndex(h => /^(n|north|northing|norht|y|شمال)$/i.test(h));
    let idxZ = headers.findIndex(h => /^(z|elev|elevation|rl|level|منسوب)$/i.test(h));
    let idxC = headers.findIndex(h => /^(code|desc|description|note|كود|وصف)$/i.test(h));

    // في الجدول أول عمودين هما (#) و (التصنيف الهندسي)
    if (idxP < 0) idxP = 2;
    if (idxE < 0) idxE = 3;
    if (idxN < 0) idxN = 4;
    if (idxZ < 0) idxZ = 5;
    if (idxC < 0) idxC = 6;

    trs.forEach((tr, rIdx) => {
      const cells = Array.from(tr.children);
      const getVal = colIndex => {
        const td = cells[colIndex];
        if (!td) return "";
        const inp = td.querySelector("input");
        return inp ? inp.value.trim() : td.textContent.trim();
      };

      const e = parseFloat(getVal(idxE));
      const n = parseFloat(getVal(idxN));
      const z = parseFloat(getVal(idxZ));
      const p = getVal(idxP) || String(rIdx + 1);
      const code = getVal(idxC) || "DEFAULT";

      if (!isNaN(e) && !isNaN(n)) {
        pts.push({ P: p, E: e, N: n, Z: isNaN(z) ? 0 : z, Code: code });
      }
    });

    return pts;
  }

  function parsePointsFromFile(file) {
    return new Promise((resolve, reject) => {
      if (!file) return resolve([]);
      const ext = file.name.split(".").pop().toLowerCase();

      if (ext === "xlsx" || ext === "xls") {
        ensureXLSX(() => {
          const reader = new FileReader();
          reader.onload = e => {
            try {
              const wb = window.XLSX.read(e.target.result, { type: "array" });
              const sheet = wb.Sheets[wb.SheetNames[0]];
              const aoa = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
              const pts = [];
              aoa.forEach((r, idx) => {
                if (!Array.isArray(r) || r.length < 3) return;
                const eVal = parseFloat(r[1]);
                const nVal = parseFloat(r[2]);
                const zVal = parseFloat(r[3]);
                if (!isNaN(eVal) && !isNaN(nVal)) {
                  pts.push({
                    P: String(r[0] || idx + 1).trim(),
                    E: eVal,
                    N: nVal,
                    Z: isNaN(zVal) ? 0 : zVal,
                    Code: String(r[4] || "PT").trim()
                  });
                }
              });
              resolve(pts);
            } catch (err) {
              reject(err);
            }
          };
          reader.readAsArrayBuffer(file);
        });
        return;
      }

      const reader = new FileReader();
      reader.onload = e => {
        const text = String(e.target.result || "");
        const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        const pts = [];

        if (ext === "dxf") {
          let counter = 1;
          for (let i = 0; i < lines.length - 1; i++) {
            if (lines[i] === "0" && (lines[i + 1] === "POINT" || lines[i + 1] === "VERTEX")) {
              let x = null, y = null, z = 0, layer = "DXF";
              for (let j = i + 2; j < Math.min(i + 40, lines.length - 1); j += 2) {
                if (lines[j] === "0") break;
                if (lines[j] === "8") layer = lines[j + 1];
                if (lines[j] === "10") x = parseFloat(lines[j + 1]);
                if (lines[j] === "20") y = parseFloat(lines[j + 1]);
                if (lines[j] === "30") z = parseFloat(lines[j + 1]) || 0;
              }
              if (x !== null && y !== null && !isNaN(x) && !isNaN(y)) {
                pts.push({ P: String(counter++), E: x, N: y, Z: z, Code: layer });
              }
            }
          }
          return resolve(pts);
        }

        if (ext === "sdr" || ext === "str") {
          lines.forEach(line => {
            if (line.startsWith("08KI") || line.startsWith("08TP")) {
              const p = line.slice(4, 20).trim();
              const n = parseFloat(line.slice(20, 36).trim());
              const eVal = parseFloat(line.slice(36, 52).trim());
              const z = parseFloat(line.slice(52, 68).trim());
              const code = line.slice(68).trim() || "SDR";
              if (!isNaN(eVal) && !isNaN(n)) {
                pts.push({ P: p || String(pts.length + 1), E: eVal, N: n, Z: isNaN(z) ? 0 : z, Code: code });
              }
            }
          });
          return resolve(pts);
        }

        const sep = lines[0]?.includes("\t") ? "\t" : lines[0]?.includes(";") ? ";" : ",";
        lines.forEach((ln, idx) => {
          const parts = ln.split(sep);
          if (parts.length >= 3) {
            const eVal = parseFloat(parts[1]);
            const nVal = parseFloat(parts[2]);
            const zVal = parseFloat(parts[3]);
            if (!isNaN(eVal) && !isNaN(nVal)) {
              pts.push({
                P: String(parts[0] || idx + 1).trim(),
                E: eVal,
                N: nVal,
                Z: isNaN(zVal) ? 0 : zVal,
                Code: String(parts[4] || "PT").trim()
              });
            }
          }
        });
        resolve(pts);
      };
      reader.readAsText(file);
    });
  }

  /* ============================================================================
     2. محرك حصر المساحات المغلقة وتجميع الأطوال دفعة واحدة
     ============================================================================ */
  function computeBatchAreasAndLengths(points, forceCloseAll = true) {
    if (!points || points.length < 2) return null;

    // تجميع النقاط حسب الكود (Code) لحصر كل مضلع/خط مغلق دفعة واحدة
    const groupsMap = new Map();
    points.forEach(pt => {
      const key = String(pt.Code || "عام").trim() || "عام";
      if (!groupsMap.has(key)) groupsMap.set(key, []);
      groupsMap.get(key).push(pt);
    });

    const items = [];
    let grandTotalOpenLength = 0;
    let grandTotalClosedPerim = 0;
    let grandTotalAreaM2 = 0;

    groupsMap.forEach((pts, codeName) => {
      if (pts.length < 2) return;

      let openLen = 0;
      for (let i = 0; i < pts.length - 1; i++) {
        openLen += Math.hypot(pts[i + 1].E - pts[i].E, pts[i + 1].N - pts[i].N);
      }

      const closingDist = pts.length >= 3
        ? Math.hypot(pts[0].E - pts[pts.length - 1].E, pts[0].N - pts[pts.length - 1].N)
        : 0;

      const isClosed = pts.length >= 3 && (forceCloseAll || closingDist <= 1.0);
      const closedPerim = isClosed ? openLen + closingDist : openLen;

      let areaM2 = 0;
      if (isClosed && pts.length >= 3) {
        let sum2 = 0;
        for (let i = 0; i < pts.length; i++) {
          const j = (i + 1) % pts.length;
          sum2 += pts[i].E * pts[j].N - pts[j].E * pts[i].N;
        }
        areaM2 = Math.abs(sum2) / 2;
      }

      const feddanTotal = areaM2 / 4200.83;
      const fed = Math.floor(feddanTotal);
      const qirFloat = (feddanTotal - fed) * 24;
      const qir = Math.floor(qirFloat);
      const sahm = (qirFloat - qir) * 24;

      grandTotalOpenLength += openLen;
      grandTotalClosedPerim += closedPerim;
      grandTotalAreaM2 += areaM2;

      items.push({
        code: codeName,
        pointCount: pts.length,
        isClosed,
        openLen,
        closedPerim,
        areaM2,
        feddanStr: areaM2 > 0 ? `${fed} ف • ${qir} ط • ${sahm.toFixed(1)} س` : "-"
      });
    });

    const totalFed = grandTotalAreaM2 / 4200.83;
    const tFed = Math.floor(totalFed);
    const tQirF = (totalFed - tFed) * 24;
    const tQir = Math.floor(tQirF);
    const tSahm = (tQirF - tQir) * 24;

    return {
      items,
      grandTotalOpenLength,
      grandTotalClosedPerim,
      grandTotalAreaM2,
      grandFeddanStr: `${tFed} فدان • ${tQir} قيراط • ${tSahm.toFixed(1)} سهم`
    };
  }

  /* ============================================================================
     3. محرك حساب كميات الحفر والردم بين سطحين (Surface-to-Surface Cut/Fill)
     ============================================================================ */
  function pointInPolygon(x, y, poly) {
    let inside = false;
    for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
      const xi = poly[i].E, yi = poly[i].N;
      const xj = poly[j].E, yj = poly[j].N;
      const intersect = ((yi > y) !== (yj > y)) && (x < ((xj - xi) * (y - yi)) / (yj - yi + 1e-12) + xi);
      if (intersect) inside = !inside;
    }
    return inside;
  }

  function computeConvexHull(pts) {
    if (pts.length < 3) return pts;
    const sorted = [...pts].sort((a, b) => (a.E === b.E ? a.N - b.N : a.E - b.E));
    const cross = (o, a, b) => (a.E - o.E) * (b.N - o.N) - (a.N - o.N) * (b.E - o.E);

    const lower = [];
    for (const p of sorted) {
      while (lower.length >= 2 && cross(lower[lower.length - 2], lower[lower.length - 1], p) <= 0) {
        lower.pop();
      }
      lower.push(p);
    }

    const upper = [];
    for (let i = sorted.length - 1; i >= 0; i--) {
      const p = sorted[i];
      while (upper.length >= 2 && cross(upper[upper.length - 2], upper[upper.length - 1], p) <= 0) {
        upper.pop();
      }
      upper.push(p);
    }

    upper.pop();
    lower.pop();
    return lower.concat(upper);
  }

  // استيفاء منسوب السطح عند أي إحداثي (E, N) بطريقة IDW الدقيقة لأقرب 6 نقاط
  function interpolateSurfaceZ(e, n, surfacePts) {
    if (!surfacePts.length) return NaN;
    if (surfacePts.length === 1) return surfacePts[0].Z;

    const nearest = surfacePts
      .map(p => ({ z: p.Z, d2: (p.E - e) ** 2 + (p.N - n) ** 2 }))
      .sort((a, b) => a.d2 - b.d2)
      .slice(0, 6);

    if (nearest[0].d2 < 1e-6) return nearest[0].z;

    let num = 0;
    let den = 0;
    for (const item of nearest) {
      const w = 1 / item.d2;
      num += w * item.z;
      den += w;
    }
    return num / den;
  }

  function calculateTwoSurfacesVolume(nglPts, designPts, gridStepMeters = 1.0) {
    if (nglPts.length < 3 || designPts.length < 1) return null;

    const hull = computeConvexHull(nglPts);
    const minE = Math.min(...hull.map(p => p.E));
    const maxE = Math.max(...hull.map(p => p.E));
    const minN = Math.min(...hull.map(p => p.N));
    const maxN = Math.max(...hull.map(p => p.N));

    const spanE = Math.max(maxE - minE, 1);
    const spanN = Math.max(maxN - minN, 1);

    // ضبط دقة الشبكة تلقائياً لضمان سرعة فائقة ودقة هندسية عالية (حتى 120×120 خلية)
    const autoStep = Math.max(gridStepMeters, Math.max(spanE, spanN) / 110);
    const cellArea = autoStep * autoStep;

    let cutVol = 0;
    let fillVol = 0;
    let cutArea = 0;
    let fillArea = 0;
    let zeroArea = 0;
    let sampleCells = 0;

    for (let x = minE + autoStep / 2; x <= maxE; x += autoStep) {
      for (let y = minN + autoStep / 2; y <= maxN; y += autoStep) {
        if (!pointInPolygon(x, y, hull)) continue;

        const zNGL = interpolateSurfaceZ(x, y, nglPts);
        const zDES = interpolateSurfaceZ(x, y, designPts);
        if (isNaN(zNGL) || isNaN(zDES)) continue;

        sampleCells++;
        const diff = zNGL - zDES; // موجب = حفر (Cut) | سالب = ردم (Fill)

        if (diff > 0.001) {
          cutVol += diff * cellArea;
          cutArea += cellArea;
        } else if (diff < -0.001) {
          fillVol += Math.abs(diff) * cellArea;
          fillArea += cellArea;
        } else {
          zeroArea += cellArea;
        }
      }
    }

    const totalArea = cutArea + fillArea + zeroArea;
    const netVol = cutVol - fillVol;

    return {
      nglCount: nglPts.length,
      designCount: designPts.length,
      gridStep: autoStep,
      sampleCells,
      totalArea,
      cutArea,
      fillArea,
      cutVol,
      fillVol,
      netVol
    };
  }

  /* ============================================================================
     4. واجهة المستخدم الهندسية المدمجة داخل محطة إكسيل المساحة
     ============================================================================ */
  function injectStyles() {
    if (document.getElementById("gdVolTakeoffStyles")) return;
    const st = document.createElement("style");
    st.id = "gdVolTakeoffStyles";
    st.textContent = `
      #gdVolTakeoffSuite {
        direction: rtl;
        text-align: right;
        background: var(--ex-panel, #100e0c);
        border: 1.5px solid #e09f24;
        border-radius: 15px;
        padding: 14px;
        margin-top: 14px;
        color: var(--ex-text, #f5f5f4);
      }
      .gd-vt-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 8px;
        padding-bottom: 10px;
        border-bottom: 1px solid var(--ex-line, #3f3a36);
        margin-bottom: 12px;
      }
      .gd-vt-header h4 {
        margin: 0;
        font-size: 15px;
        font-weight: 900;
        color: #fbbf24;
      }
      .gd-vt-tabs {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 8px;
        margin-bottom: 12px;
      }
      .gd-vt-tab {
        padding: 10px;
        border-radius: 10px;
        border: 1.5px solid var(--ex-line, #3f3a36);
        background: var(--ex-card, #211e1b);
        color: var(--ex-muted, #a8a29e);
        font-size: 12.5px;
        font-weight: 900;
        cursor: pointer;
        font-family: inherit;
        text-align: center;
      }
      .gd-vt-tab.active {
        border-color: #e09f24;
        background: rgba(224, 159, 36, 0.18);
        color: #fbbf24;
      }
      .gd-vt-surf-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
        gap: 10px;
        margin-bottom: 12px;
      }
      .gd-vt-surf-card {
        background: var(--ex-card, #211e1b);
        border: 1.5px solid var(--ex-line, #3f3a36);
        border-radius: 12px;
        padding: 11px;
      }
      .gd-vt-surf-card strong {
        display: block;
        font-size: 13px;
        margin-bottom: 4px;
      }
      .gd-vt-kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(135px, 1fr));
        gap: 8px;
        margin: 12px 0;
      }
      .gd-vt-kpi {
        background: var(--ex-card, #211e1b);
        border: 1px solid var(--ex-line, #3f3a36);
        border-radius: 11px;
        padding: 10px;
        text-align: center;
      }
      .gd-vt-kpi span { display: block; font-size: 11px; color: var(--ex-muted, #a8a29e); }
      .gd-vt-kpi strong { font-size: 15px; font-weight: 900; display: block; margin-top: 2px; }
      .gd-vt-table-wrap {
        max-height: 250px;
        overflow: auto;
        border: 1px solid var(--ex-line, #3f3a36);
        border-radius: 10px;
        margin-top: 10px;
      }
      .gd-vt-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 12px;
        text-align: center;
      }
      .gd-vt-table th {
        background: var(--ex-th-bg, #292524);
        color: #fbbf24;
        padding: 8px 6px;
        border: 1px solid var(--ex-line, #3f3a36);
        position: sticky;
        top: 0;
      }
      .gd-vt-table td {
        padding: 7px 6px;
        border: 1px solid var(--ex-line, #3f3a36);
      }
    `;
    document.head.appendChild(st);
  }

  function renderSuiteUI() {
    const studio = document.getElementById("gdAiExcelStudio");
    if (!studio || document.getElementById("gdVolTakeoffSuite")) return;

    const box = document.createElement("div");
    box.id = "gdVolTakeoffSuite";
    box.innerHTML = `
      <div class="gd-vt-header">
        <h4>حصر المساحات المغلقة وتجميع الأطوال + كميات الحفر والردم بين سطحين (TIN / Grid)</h4>
        <button type="button" id="gdExportTakeoffExcelBtn" class="gd-ex-btn gold">
          تصدير الحصر والكميات Excel (.xlsx)
        </button>
      </div>

      <div class="gd-vt-tabs">
        <button type="button" class="gd-vt-tab active" data-vt-tab="surfaces">
          1. حساب كميات الحفر والردم (سطحين: NGL + Design)
        </button>
        <button type="button" class="gd-vt-tab" data-vt-tab="takeoff">
          2. حصر المساحات المغلقة وتجميع الأطوال دفعة واحدة
        </button>
      </div>

      <!-- القسم الأول: كميات الحفر والردم بين سطحين -->
      <div id="gdVtSurfacesPane">
        <div class="gd-vt-surf-grid">
          <div class="gd-vt-surf-card" style="border-color:#10b981;">
            <strong style="color:#34d399;">السطح الأول: الأرض الطبيعية (NGL / EG)</strong>
            <div id="gdSurf1Status" style="font-size:11.5px;color:#a8a29e;margin-bottom:8px;">
              لم يتم تحديد السطح الأول بعد (يمكنك سحبه من الجدول الحالي أو رفع ملف)
            </div>
            <div style="display:flex;gap:6px;flex-wrap:wrap;">
              <button type="button" id="gdLoadSurf1FromTableBtn" class="gd-ex-btn" style="flex:1;justify-content:center;">
                استخدام الجدول الحالي
              </button>
              <label class="gd-ex-btn gold" style="flex:1;justify-content:center;cursor:pointer;">
                رفع ملف NGL
                <input type="file" id="gdSurf1FileInp" accept=".xlsx,.xls,.csv,.txt,.dxf,.sdr,.str" style="display:none;">
              </label>
            </div>
          </div>

          <div class="gd-vt-surf-card" style="border-color:#38bdf8;">
            <strong style="color:#38bdf8;">السطح الثاني: سطح التصميم (Design Surface)</strong>
            <div id="gdSurf2Status" style="font-size:11.5px;color:#a8a29e;margin-bottom:8px;">
              ارفع ملف نقاط سطح التصميم أو أدخل منسوب تصميم موحد
            </div>
            <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;">
              <label class="gd-ex-btn gold" style="flex:1;justify-content:center;cursor:pointer;">
                رفع ملف التصميم
                <input type="file" id="gdSurf2FileInp" accept=".xlsx,.xls,.csv,.txt,.dxf,.sdr,.str" style="display:none;">
              </label>
              <input type="number" step="any" id="gdFlatDesignZInp" class="gd-ex-input"
                style="width:110px;min-width:100px;padding:7px;text-align:center;margin:0;"
                placeholder="أو منسوب ثابت">
              <button type="button" id="gdApplyFlatSurf2Btn" class="gd-ex-btn">تطبيق</button>
            </div>
          </div>
        </div>

        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
          <button type="button" id="gdComputeTwoSurfacesBtn" class="gd-ex-btn gold" style="flex:1;justify-content:center;padding:11px;font-size:13px;">
            احسب كميات الحفر والردم بين السطحين فوراً
          </button>
        </div>

        <div id="gdTwoSurfacesResultArea"></div>
      </div>

      <!-- القسم الثاني: حصر المساحات المغلقة وتجميع الأطوال دفعة واحدة -->
      <div id="gdVtTakeoffPane" style="display:none;">
        <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:10px;">
          <button type="button" id="gdRunBatchTakeoffBtn" class="gd-ex-btn gold" style="flex:1;justify-content:center;padding:11px;font-size:13px;">
            حصر جميع المساحات المغلقة وتجميع الأطوال من الجدول دفعة واحدة
          </button>
          <label class="gd-ex-btn" style="cursor:pointer;padding:10px 12px;">
            أو رفع ملف للحصر الفوري
            <input type="file" id="gdTakeoffFileInp" accept=".xlsx,.xls,.csv,.txt,.dxf,.sdr,.str" style="display:none;">
          </label>
        </div>

        <div id="gdBatchTakeoffResultArea"></div>
      </div>
    `;

    studio.appendChild(box);
    bindSuiteEvents();
  }

  function renderVolumeResults(res) {
    const area = document.getElementById("gdTwoSurfacesResultArea");
    if (!area) return;
    if (!res) {
      area.innerHTML = `<div style="margin-top:10px;color:#f87171;font-size:12.5px;font-weight:800;">يرجى تحميل السطح الأول (3 نقاط على الأقل) والسطح الثاني أولاً.</div>`;
      return;
    }

    const netLabel = res.netVol >= 0
      ? `فائض حفر (+${res.netVol.toFixed(2)} م³)`
      : `عجز / احتياج ردم (${res.netVol.toFixed(2)} م³)`;

    area.innerHTML = `
      <div class="gd-vt-kpi-grid">
        <div class="gd-vt-kpi" style="border-color:#ef4444;">
          <span>إجمالي كمية الحفر (Cut)</span>
          <strong style="color:#f87171;">${res.cutVol.toFixed(2)} م³</strong>
          <small style="font-size:10px;color:#a8a29e;">مساحة الحفر: ${res.cutArea.toFixed(1)} م²</small>
        </div>
        <div class="gd-vt-kpi" style="border-color:#38bdf8;">
          <span>إجمالي كمية الردم (Fill)</span>
          <strong style="color:#38bdf8;">${res.fillVol.toFixed(2)} م³</strong>
          <small style="font-size:10px;color:#a8a29e;">مساحة الردم: ${res.fillArea.toFixed(1)} م²</small>
        </div>
        <div class="gd-vt-kpi" style="border-color:#fbbf24;">
          <span>صافي الكميات (Net)</span>
          <strong style="color:#fbbf24;">${netLabel}</strong>
          <small style="font-size:10px;color:#a8a29e;">إجمالي المسطح: ${res.totalArea.toFixed(1)} م²</small>
        </div>
        <div class="gd-vt-kpi">
          <span>دقة الشبكة الحسابية</span>
          <strong style="color:#34d399;">${res.gridStep.toFixed(2)} م</strong>
          <small style="font-size:10px;color:#a8a29e;">${res.sampleCells} خلية حسابية</small>
        </div>
      </div>
    `;
  }

  function renderTakeoffResults(rep) {
    const area = document.getElementById("gdBatchTakeoffResultArea");
    if (!area) return;
    if (!rep || !rep.items.length) {
      area.innerHTML = `<div style="margin-top:10px;color:#f87171;font-size:12.5px;font-weight:800;">لا توجد نقاط كافية في الجدول لحصر الأطوال والمساحات.</div>`;
      return;
    }

    area.innerHTML = `
      <div class="gd-vt-kpi-grid">
        <div class="gd-vt-kpi" style="border-color:#38bdf8;">
          <span>إجمالي الأطوال المفتوحة</span>
          <strong style="color:#38bdf8;">${rep.grandTotalOpenLength.toFixed(2)} م.ط</strong>
        </div>
        <div class="gd-vt-kpi" style="border-color:#fbbf24;">
          <span>إجمالي المحيطات المغلقة</span>
          <strong style="color:#fbbf24;">${rep.grandTotalClosedPerim.toFixed(2)} م.ط</strong>
        </div>
        <div class="gd-vt-kpi" style="border-color:#22c55e;">
          <span>إجمالي المساحات المغلقة</span>
          <strong style="color:#4ade80;">${rep.grandTotalAreaM2.toFixed(2)} م²</strong>
          <small style="font-size:10px;color:#cbd5e1;">${rep.grandFeddanStr}</small>
        </div>
      </div>

      <div class="gd-vt-table-wrap">
        <table class="gd-vt-table">
          <thead>
            <tr>
              <th>الكود / الشكل</th>
              <th>عدد النقاط</th>
              <th>الطول المفتوح (م.ط)</th>
              <th>المحيط المغلق (م.ط)</th>
              <th>المساحة (م²)</th>
              <th>المساحة بالفدان</th>
            </tr>
          </thead>
          <tbody>
            ${rep.items.map(it => `
              <tr>
                <td><b>${it.code}</b></td>
                <td>${it.pointCount}</td>
                <td>${it.openLen.toFixed(2)}</td>
                <td>${it.closedPerim.toFixed(2)}</td>
                <td style="color:#4ade80;font-weight:900;">${it.areaM2.toFixed(2)}</td>
                <td>${it.feddanStr}</td>
              </tr>
            `).join("")}
            <tr style="background:rgba(224,159,36,0.18);font-weight:900;">
              <td>الإجمالي الكلي</td>
              <td>${rep.items.reduce((s, x) => s + x.pointCount, 0)}</td>
              <td>${rep.grandTotalOpenLength.toFixed(2)} م.ط</td>
              <td>${rep.grandTotalClosedPerim.toFixed(2)} م.ط</td>
              <td style="color:#fbbf24;">${rep.grandTotalAreaM2.toFixed(2)} م²</td>
              <td>${rep.grandFeddanStr}</td>
            </tr>
          </tbody>
        </table>
      </div>
    `;
  }

  function bindSuiteEvents() {
    // التبديل بين تبويب السطحين وتبويب حصر المساحات والأطوال
    document.querySelectorAll("[data-vt-tab]").forEach(btn => {
      btn.onclick = () => {
        const tab = btn.dataset.vtTab;
        document.querySelectorAll("[data-vt-tab]").forEach(b => b.classList.toggle("active", b === btn));
        document.getElementById("gdVtSurfacesPane").style.display = tab === "surfaces" ? "block" : "none";
        document.getElementById("gdVtTakeoffPane").style.display = tab === "takeoff" ? "block" : "none";
      };
    });

    // تحميل السطح الأول (NGL) من الجدول الحالي
    document.getElementById("gdLoadSurf1FromTableBtn").onclick = () => {
      const pts = readCurrentExcelTablePoints();
      if (pts.length < 3) {
        if (typeof toast === "function") toast("الجدول الحالي يحتوي على أقل من 3 نقاط", "warn");
        return;
      }
      surface1NGL = pts;
      document.getElementById("gdSurf1Status").innerHTML = `<span style="color:#4ade80;font-weight:900;">تم تحميل السطح الأول (${surface1NGL.length} نقطة من الجدول)</span>`;
      if (surface2Design.length) {
        lastVolumeReport = calculateTwoSurfacesVolume(surface1NGL, surface2Design);
        renderVolumeResults(lastVolumeReport);
      }
    };

    // رفع ملف السطح الأول (NGL)
    document.getElementById("gdSurf1FileInp").onchange = async e => {
      const file = e.target.files?.[0];
      if (!file) return;
      const pts = await parsePointsFromFile(file);
      e.target.value = "";
      if (pts.length < 3) {
        if (typeof toast === "function") toast("ملف السطح الأول يجب أن يحتوي على 3 نقاط على الأقل", "error");
        return;
      }
      surface1NGL = pts;
      document.getElementById("gdSurf1Status").innerHTML = `<span style="color:#4ade80;font-weight:900;">تم تحميل السطح الأول: ${file.name} (${pts.length} نقطة)</span>`;
      if (surface2Design.length) {
        lastVolumeReport = calculateTwoSurfacesVolume(surface1NGL, surface2Design);
        renderVolumeResults(lastVolumeReport);
      }
    };

    // رفع ملف السطح الثاني (Design Surface)
    document.getElementById("gdSurf2FileInp").onchange = async e => {
      const file = e.target.files?.[0];
      if (!file) return;
      const pts = await parsePointsFromFile(file);
      e.target.value = "";
      if (!pts.length) {
        if (typeof toast === "function") toast("لم يتم العثور على نقاط صالحة في ملف التصميم", "error");
        return;
      }
      surface2Design = pts;
      document.getElementById("gdSurf2Status").innerHTML = `<span style="color:#38bdf8;font-weight:900;">تم تحميل سطح التصميم: ${file.name} (${pts.length} نقطة)</span>`;

      // إذا لم يُحدد السطح الأول بعد، اسحبه تلقائياً من الجدول الحالي واحسب فوراً
      if (!surface1NGL.length) {
        const tablePts = readCurrentExcelTablePoints();
        if (tablePts.length >= 3) {
          surface1NGL = tablePts;
          document.getElementById("gdSurf1Status").innerHTML = `<span style="color:#4ade80;font-weight:900;">تم سحب السطح الأول تلقائياً من الجدول (${surface1NGL.length} نقطة)</span>`;
        }
      }

      if (surface1NGL.length >= 3) {
        lastVolumeReport = calculateTwoSurfacesVolume(surface1NGL, surface2Design);
        renderVolumeResults(lastVolumeReport);
      }
    };

    // تطبيق منسوب تصميم موحد كسطح ثانٍ
    document.getElementById("gdApplyFlatSurf2Btn").onclick = () => {
      const zVal = parseFloat(document.getElementById("gdFlatDesignZInp").value);
      if (isNaN(zVal)) {
        if (typeof toast === "function") toast("أدخل منسوب التصميم الثابت أولاً", "warn");
        return;
      }
      if (!surface1NGL.length) {
        surface1NGL = readCurrentExcelTablePoints();
        if (surface1NGL.length >= 3) {
          document.getElementById("gdSurf1Status").innerHTML = `<span style="color:#4ade80;font-weight:900;">تم سحب السطح الأول تلقائياً من الجدول (${surface1NGL.length} نقطة)</span>`;
        }
      }
      surface2Design = (surface1NGL.length ? surface1NGL : [{ E: 0, N: 0 }]).map(p => ({
        P: p.P,
        E: p.E,
        N: p.N,
        Z: zVal,
        Code: "FLAT_DESIGN"
      }));
      document.getElementById("gdSurf2Status").innerHTML = `<span style="color:#38bdf8;font-weight:900;">سطح تصميم أفقي موحد عند منسوب: ${zVal.toFixed(3)} م</span>`;

      if (surface1NGL.length >= 3) {
        lastVolumeReport = calculateTwoSurfacesVolume(surface1NGL, surface2Design);
        renderVolumeResults(lastVolumeReport);
      }
    };

    // زر الحساب المباشر لكميات السطحين
    document.getElementById("gdComputeTwoSurfacesBtn").onclick = () => {
      if (!surface1NGL.length) {
        surface1NGL = readCurrentExcelTablePoints();
        if (surface1NGL.length >= 3) {
          document.getElementById("gdSurf1Status").innerHTML = `<span style="color:#4ade80;font-weight:900;">تم سحب السطح الأول من الجدول (${surface1NGL.length} نقطة)</span>`;
        }
      }
      lastVolumeReport = calculateTwoSurfacesVolume(surface1NGL, surface2Design);
      renderVolumeResults(lastVolumeReport);
    };

    // زر حصر المساحات المغلقة وتجميع الأطوال دفعة واحدة من الجدول
    document.getElementById("gdRunBatchTakeoffBtn").onclick = () => {
      const pts = readCurrentExcelTablePoints();
      lastTakeoffReport = computeBatchAreasAndLengths(pts, true);
      renderTakeoffResults(lastTakeoffReport);
    };

    // رفع ملف خارجي للحصر الفوري للمساحات والأطوال
    document.getElementById("gdTakeoffFileInp").onchange = async e => {
      const file = e.target.files?.[0];
      if (!file) return;
      const pts = await parsePointsFromFile(file);
      e.target.value = "";
      lastTakeoffReport = computeBatchAreasAndLengths(pts, true);
      renderTakeoffResults(lastTakeoffReport);
    };

    // تصدير نتائج الحصر والكميات إلى Excel (.xlsx)
    document.getElementById("gdExportTakeoffExcelBtn").onclick = () => {
      if (!lastVolumeReport && !lastTakeoffReport) {
        const pts = readCurrentExcelTablePoints();
        if (pts.length >= 2) {
          lastTakeoffReport = computeBatchAreasAndLengths(pts, true);
          renderTakeoffResults(lastTakeoffReport);
        } else {
          if (typeof toast === "function") toast("قم بحساب الكميات أو حصر المساحات أولاً", "warn");
          return;
        }
      }

      ensureXLSX(() => {
        const wb = window.XLSX.utils.book_new();

        if (lastVolumeReport) {
          const volSheet = window.XLSX.utils.aoa_to_sheet([
            ["تقرير كميات الحفر والردم بين سطحين - GeoDev Studio Pro"],
            ["نقاط الأرض الطبيعية (NGL)", lastVolumeReport.nglCount],
            ["نقاط سطح التصميم (Design)", lastVolumeReport.designCount],
            ["دقة الشبكة الحسابية (متر)", Number(lastVolumeReport.gridStep.toFixed(2))],
            [],
            ["البند", "القيمة", "الوحدة"],
            ["إجمالي كمية الحفر (Cut Volume)", Number(lastVolumeReport.cutVol.toFixed(3)), "م³"],
            ["مساحة مناطق الحفر (Cut Area)", Number(lastVolumeReport.cutArea.toFixed(2)), "م²"],
            ["إجمالي كمية الردم (Fill Volume)", Number(lastVolumeReport.fillVol.toFixed(3)), "م³"],
            ["مساحة مناطق الردم (Fill Area)", Number(lastVolumeReport.fillArea.toFixed(2)), "م²"],
            ["صافي الكميات (Net Cut - Fill)", Number(lastVolumeReport.netVol.toFixed(3)), "م³"],
            ["إجمالي مساحة المشروع", Number(lastVolumeReport.totalArea.toFixed(2)), "م²"]
          ]);
          window.XLSX.utils.book_append_sheet(wb, volSheet, "Cut_Fill_Surfaces");
        }

        if (lastTakeoffReport) {
          const aoa = [
            ["حصر المساحات المغلقة وتجميع الأطوال - GeoDev Studio Pro"],
            [],
            ["الكود / المضلع", "عدد النقاط", "الطول المفتوح (م.ط)", "المحيط المغلق (م.ط)", "المساحة (م²)", "المساحة بالفدان والقيراط والسهم"]
          ];
          lastTakeoffReport.items.forEach(it => {
            aoa.push([
              it.code,
              it.pointCount,
              Number(it.openLen.toFixed(3)),
              Number(it.closedPerim.toFixed(3)),
              Number(it.areaM2.toFixed(3)),
              it.feddanStr
            ]);
          });
          aoa.push([
            "الإجمالي الكلي",
            lastTakeoffReport.items.reduce((s, x) => s + x.pointCount, 0),
            Number(lastTakeoffReport.grandTotalOpenLength.toFixed(3)),
            Number(lastTakeoffReport.grandTotalClosedPerim.toFixed(3)),
            Number(lastTakeoffReport.grandTotalAreaM2.toFixed(3)),
            lastTakeoffReport.grandFeddanStr
          ]);
          const tkSheet = window.XLSX.utils.aoa_to_sheet(aoa);
          window.XLSX.utils.book_append_sheet(wb, tkSheet, "Areas_Lengths_Takeoff");
        }

        window.XLSX.writeFile(wb, `GeoDev_Volumes_Takeoff_${Date.now()}.xlsx`);
      });
    };
  }

  function init() {
    injectStyles();
    const obs = new MutationObserver(() => renderSuiteUI());
    if (document.body) obs.observe(document.body, { childList: true, subtree: true });
    setTimeout(renderSuiteUI, 800);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
