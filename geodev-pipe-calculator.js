/* ============================================================================
   GeoDev Pipe & Trench Leveling Calculator (V2)
   - مربوط بدقة بقسم "المساحة" فقط (عبر زر المساحة في شريط التنقل)
   - حاسبة منسوب قاع الحفر + حاسبة جري ظهر الماسورة + تصدير CSV
   ============================================================================ */
(function () {
  "use strict";
  if (window.GEOPipeCalcV2Init) return;
  window.GEOPipeCalcV2Init = true;

  let activePipeTab = "trench"; // "trench" = قاع الحفر | "crown" = ظهر الماسورة
  let lastTableRows = [];
  let lastModeTitle = "";

  function injectStyles() {
    if (document.getElementById("gdPipeCalcStyles")) return;
    const style = document.createElement("style");
    style.id = "gdPipeCalcStyles";
    style.textContent = `
      #gdPipeCalcCard {
        direction: rtl;
        text-align: right;
        background: #0f172a;
        border: 1.5px solid #1d4ed8;
        border-radius: 18px;
        padding: 16px;
        margin: 16px 0;
        color: #f8fafc;
        box-shadow: 0 12px 30px rgba(0, 0, 0, 0.25);
      }
      .gd-pc-banner {
        background: linear-gradient(135deg, #0284c7, #1d4ed8);
        color: #ffffff;
        text-align: center;
        padding: 14px;
        border-radius: 13px;
        font-size: 18px;
        font-weight: 900;
        margin-bottom: 12px;
      }
      .gd-pc-tabs {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 9px;
        margin-bottom: 14px;
      }
      .gd-pc-tab-btn {
        padding: 12px 10px;
        border-radius: 11px;
        border: 1.5px solid #334155;
        background: #1e293b;
        color: #cbd5e1;
        font-size: 13.5px;
        font-weight: 800;
        cursor: pointer;
        font-family: inherit;
        transition: all 0.2s ease;
      }
      .gd-pc-tab-btn.active {
        background: #0284c7;
        border-color: #38bdf8;
        color: #ffffff;
      }
      .gd-pc-panel {
        background: #111827;
        border: 1px solid #1f2937;
        border-radius: 14px;
        padding: 14px;
      }
      .gd-pc-panel-title {
        margin: 0 0 12px;
        font-size: 15.5px;
        font-weight: 900;
        color: #38bdf8;
      }
      .gd-pc-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 10px;
      }
      .gd-pc-field label {
        display: block;
        font-size: 12px;
        font-weight: 800;
        color: #e2e8f0;
        margin-bottom: 5px;
      }
      .gd-pc-field input,
      .gd-pc-field select {
        width: 100%;
        box-sizing: border-box;
        padding: 10px 12px;
        border-radius: 10px;
        border: 1.5px solid #334155;
        background: #090d16;
        color: #ffffff;
        font-size: 14px;
        font-weight: 800;
        text-align: center;
        font-family: inherit;
      }
      .gd-pc-submit {
        width: 100%;
        margin-top: 14px;
        padding: 13px;
        border-radius: 12px;
        border: 0;
        background: linear-gradient(135deg, #0284c7, #1d4ed8);
        color: #ffffff;
        font-size: 15px;
        font-weight: 900;
        cursor: pointer;
        font-family: inherit;
      }
      .gd-pc-hint {
        margin: 9px 0 0;
        font-size: 11.5px;
        color: #94a3b8;
        text-align: center;
        line-height: 1.6;
      }
      .gd-pc-summary-box {
        background: #111827;
        border: 1px solid #334155;
        border-radius: 13px;
        padding: 13px;
        margin-top: 14px;
        text-align: center;
        line-height: 1.9;
        font-size: 14px;
        font-weight: 800;
      }
      .gd-pc-table-wrap {
        margin-top: 12px;
        overflow-x: auto;
        border-radius: 12px;
        border: 1px solid #334155;
      }
      .gd-pc-table {
        width: 100%;
        border-collapse: collapse;
        text-align: center;
        font-size: 13px;
      }
      .gd-pc-table th {
        background: #0284c7;
        color: #ffffff;
        padding: 10px 8px;
        font-weight: 900;
        border: 1px solid #0369a1;
      }
      .gd-pc-table td {
        padding: 9px 8px;
        border: 1px solid #1e293b;
        background: #090d16;
        color: #f8fafc;
        font-weight: 700;
      }
      .gd-pc-table tr:nth-child(even) td {
        background: #0f172a;
      }
      .gd-pc-export-btn {
        width: 100%;
        margin-top: 10px;
        padding: 10px;
        border-radius: 10px;
        border: 1px solid #e09f24;
        background: rgba(224, 159, 36, 0.15);
        color: #fbbf24;
        font-weight: 800;
        font-size: 13px;
        cursor: pointer;
        font-family: inherit;
      }
    `;
    document.head.appendChild(style);
  }

  function renderCalculatorHTML() {
    const isTrench = activePipeTab === "trench";

    return `
      <div class="gd-pc-banner">حاسبة أعمال الصرف والحفر</div>

      <div class="gd-pc-tabs">
        <button type="button" class="gd-pc-tab-btn ${isTrench ? "active" : ""}" data-pc-tab="trench">
          حاسبة منسوب قاع الحفر
        </button>
        <button type="button" class="gd-pc-tab-btn ${!isTrench ? "active" : ""}" data-pc-tab="crown">
          حاسبة جري ظهر الماسورة
        </button>
      </div>

      ${isTrench ? `
        <div class="gd-pc-panel">
          <h4 class="gd-pc-panel-title">بيانات قاع الحفر</h4>
          <div class="gd-pc-grid">
            <div class="gd-pc-field">
              <label>منسوب الروبير BM (م)</label>
              <input type="number" step="any" id="pcTrBm" value="20.000">
            </div>
            <div class="gd-pc-field">
              <label>قراءة الخلفية BS (م)</label>
              <input type="number" step="any" id="pcTrBs" value="1.250">
            </div>
            <div class="gd-pc-field">
              <label>منسوب الإنفيرت ليفل (م)</label>
              <input type="number" step="any" id="pcTrInvert" value="18.500">
            </div>
            <div class="gd-pc-field">
              <label>هنزل قد إيه من الإنفيرت؟ (م)</label>
              <input type="number" step="any" id="pcTrDrop" value="0.300">
            </div>
            <div class="gd-pc-field">
              <label>طول المسار (م)</label>
              <input type="number" step="any" id="pcTrLen" value="100">
            </div>
            <div class="gd-pc-field">
              <label>المسافة بين النقاط (م)</label>
              <input type="number" step="any" id="pcTrStep" value="10">
            </div>
            <div class="gd-pc-field">
              <label>الميل الطولي (%)</label>
              <input type="number" step="any" id="pcTrSlope" value="0.500">
            </div>
            <div class="gd-pc-field">
              <label>اتجاه الميل</label>
              <select id="pcTrDir">
                <option value="down" selected>المنسوب يقل مع المسافة (هبوط)</option>
                <option value="up">المنسوب يزيد مع المسافة (صعود)</option>
              </select>
            </div>
          </div>

          <button type="button" class="gd-pc-submit" id="pcCalcTrenchBtn">احسب قاع الحفر</button>
          <p class="gd-pc-hint">
            منسوب قاع الحفر المستهدف = Invert Level − مقدار النزول. لو مقدار النزول صفر، قاع الحفر يساوي الإنفيرت ليفل.
          </p>
        </div>
      ` : `
        <div class="gd-pc-panel">
          <h4 class="gd-pc-panel-title">بيانات ظهر الماسورة</h4>
          <div class="gd-pc-grid">
            <div class="gd-pc-field">
              <label>منسوب الـ BM (م)</label>
              <input type="number" step="any" id="pcCrBm" value="85.000">
            </div>
            <div class="gd-pc-field">
              <label>قراءة الـ BS (م)</label>
              <input type="number" step="any" id="pcCrBs" value="1.500">
            </div>
            <div class="gd-pc-field">
              <label>منسوب الـ Invert عند البداية (م)</label>
              <input type="number" step="any" id="pcCrInvert" value="83.000">
            </div>
            <div class="gd-pc-field">
              <label>القطر الخارجي للماسورة (م)</label>
              <input type="number" step="any" id="pcCrDia" value="0.250">
            </div>
            <div class="gd-pc-field">
              <label>ميل الخط (%)</label>
              <input type="number" step="any" id="pcCrSlope" value="0.325">
            </div>
            <div class="gd-pc-field">
              <label>طول الخط (م)</label>
              <input type="number" step="any" id="pcCrLen" value="30">
            </div>
            <div class="gd-pc-field">
              <label>المسافة بين النقاط (م)</label>
              <input type="number" step="any" id="pcCrStep" value="1">
            </div>
            <div class="gd-pc-field">
              <label>اتجاه الميل</label>
              <select id="pcCrDir">
                <option value="down" selected>هبوط</option>
                <option value="up">صعود</option>
              </select>
            </div>
          </div>

          <button type="button" class="gd-pc-submit" id="pcCalcCrownBtn">احسب ظهر الماسورة</button>
          <p class="gd-pc-hint">
            الميل يُكتب كنسبة مئوية؛ مثال 0.325 يعني 0.325%.<br>
            الحساب: HI = BM + BS، منسوب ظهر الماسورة = منسوب Invert عند البداية ± الميل + القطر الخارجي.
          </p>
        </div>
      `}

      <div id="gdPipeCalcResultsArea">
        <div class="gd-pc-summary-box" style="color:#94a3b8;">أدخل البيانات واضغط «احسب»</div>
      </div>
    `;
  }

  function calculateTrench() {
    const bm = parseFloat(document.getElementById("pcTrBm").value) || 0;
    const bs = parseFloat(document.getElementById("pcTrBs").value) || 0;
    const invert = parseFloat(document.getElementById("pcTrInvert").value) || 0;
    const drop = parseFloat(document.getElementById("pcTrDrop").value) || 0;
    const totalLen = parseFloat(document.getElementById("pcTrLen").value) || 0;
    const step = parseFloat(document.getElementById("pcTrStep").value) || 1;
    const slopePct = parseFloat(document.getElementById("pcTrSlope").value) || 0;
    const dir = document.getElementById("pcTrDir").value === "up" ? 1 : -1;

    if (step <= 0 || totalLen < 0) return;

    const hi = bm + bs;
    const startTrench = invert - drop;
    const startStaff = hi - startTrench;
    const slopeDec = slopePct / 100;

    const rows = [];
    let idx = 1;
    for (let d = 0; d <= totalLen + 1e-9; d += step) {
      const dist = Math.min(d, totalLen);
      const level = startTrench + dir * (dist * slopeDec);
      const staff = hi - level;
      rows.push({
        num: idx++,
        dist: dist.toFixed(3),
        level: level.toFixed(3),
        staff: staff.toFixed(3)
      });
      if (dist >= totalLen) break;
    }

    lastTableRows = rows;
    lastModeTitle = "منسوب قاع الحفر (م)";

    const out = document.getElementById("gdPipeCalcResultsArea");
    out.innerHTML = `
      <div class="gd-pc-summary-box">
        <div>منسوب خط النظر HI = <span style="color:#38bdf8;">${hi.toFixed(3)}</span> م</div>
        <div>منسوب الإنفيرت ليفل = ${invert.toFixed(3)} م</div>
        <div>مقدار النزول = ${drop.toFixed(3)} م</div>
        <div>منسوب قاع الحفر عند البداية = <span style="color:#fbbf24;">${startTrench.toFixed(3)}</span> م</div>
        <div>قراءة القامة المطلوبة عند البداية = <span style="color:#4ade80;">${startStaff.toFixed(3)}</span> م</div>
      </div>

      <div class="gd-pc-table-wrap">
        <table class="gd-pc-table">
          <thead>
            <tr>
              <th>رقم النقطة</th>
              <th>المسافة (م)</th>
              <th>منسوب قاع الحفر (م)</th>
              <th>قراءة القامة المطلوبة (م)</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(r => `
              <tr>
                <td>${r.num}</td>
                <td>${r.dist}</td>
                <td>${r.level}</td>
                <td style="color:#4ade80;">${r.staff}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
      <button type="button" class="gd-pc-export-btn" id="gdExportPipeCsvBtn">تصدير الجدول إلى ملف Excel / CSV</button>
    `;
  }

  function calculateCrown() {
    const bm = parseFloat(document.getElementById("pcCrBm").value) || 0;
    const bs = parseFloat(document.getElementById("pcCrBs").value) || 0;
    const invert = parseFloat(document.getElementById("pcCrInvert").value) || 0;
    const dia = parseFloat(document.getElementById("pcCrDia").value) || 0;
    const slopePct = parseFloat(document.getElementById("pcCrSlope").value) || 0;
    const totalLen = parseFloat(document.getElementById("pcCrLen").value) || 0;
    const step = parseFloat(document.getElementById("pcCrStep").value) || 1;
    const dir = document.getElementById("pcCrDir").value === "up" ? 1 : -1;

    if (step <= 0 || totalLen < 0) return;

    const hi = bm + bs;
    const invertStaff = hi - invert;
    const startCrown = invert + dia;
    const startStaff = hi - startCrown;
    const slopeDec = slopePct / 100;

    const rows = [];
    let idx = 1;
    for (let d = 0; d <= totalLen + 1e-9; d += step) {
      const dist = Math.min(d, totalLen);
      const level = startCrown + dir * (dist * slopeDec);
      const staff = hi - level;
      rows.push({
        num: idx++,
        dist: dist.toFixed(3),
        level: level.toFixed(3),
        staff: staff.toFixed(3)
      });
      if (dist >= totalLen) break;
    }

    lastTableRows = rows;
    lastModeTitle = "منسوب ظهر الماسورة (م)";

    const out = document.getElementById("gdPipeCalcResultsArea");
    out.innerHTML = `
      <div class="gd-pc-summary-box">
        <div>HI = <span style="color:#38bdf8;">${hi.toFixed(3)}</span></div>
        <div>Invert Reading = ${invertStaff.toFixed(3)}</div>
        <div>منسوب ظهر الماسورة عند البداية = <span style="color:#fbbf24;">${startCrown.toFixed(3)}</span> م</div>
        <div>قراءة القامة عند البداية = <span style="color:#4ade80;">${startStaff.toFixed(3)}</span> م</div>
      </div>

      <div class="gd-pc-table-wrap">
        <table class="gd-pc-table">
          <thead>
            <tr>
              <th>رقم النقطة</th>
              <th>المسافة (م)</th>
              <th>منسوب ظهر الماسورة (م)</th>
              <th>قراءة القامة المطلوبة (م)</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(r => `
              <tr>
                <td>${r.num}</td>
                <td>${r.dist}</td>
                <td>${r.level}</td>
                <td style="color:#4ade80;">${r.staff}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
      <button type="button" class="gd-pc-export-btn" id="gdExportPipeCsvBtn">تصدير الجدول إلى ملف Excel / CSV</button>
    `;
  }

  function exportCsvTable() {
    if (!lastTableRows.length) return;
    const header = `رقم النقطة,المسافة (م),${lastModeTitle},قراءة القامة المطلوبة (م)\n`;
    const body = lastTableRows.map(r => `${r.num},${r.dist},${r.level},${r.staff}`).join("\n");
    const blob = new Blob(["\uFEFF" + header + body], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `GeoDev_PipeCalculator_${Date.now()}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  // تحديد قسم "المساحة" بدقة من زر التبويب في شريط التنقل السفلي
  function findSurveyingSection() {
    const navButtons = Array.from(document.querySelectorAll("nav button, [data-t]"));
    const survBtn = navButtons.find(btn => (btn.textContent || "").trim() === "المساحة");
    if (survBtn && survBtn.dataset && survBtn.dataset.t) {
      const secByTab = document.getElementById("s-" + survBtn.dataset.t);
      if (secByTab) return secByTab;
    }

    return (
      document.getElementById("s-surv") ||
      document.getElementById("s-surveying") ||
      document.getElementById("s-sv") ||
      null
    );
  }

  function mountPipeCalculator() {
    const targetSec = findSurveyingSection();
    if (!targetSec) return;

    const existingCard = document.getElementById("gdPipeCalcCard");
    if (existingCard) {
      if (existingCard.parentElement !== targetSec) {
        targetSec.appendChild(existingCard);
      }
      return;
    }

    const card = document.createElement("div");
    card.id = "gdPipeCalcCard";
    card.innerHTML = renderCalculatorHTML();
    targetSec.appendChild(card);
  }

  function bindEvents() {
    document.addEventListener("click", e => {
      const tabBtn = e.target.closest("[data-pc-tab]");
      if (tabBtn) {
        activePipeTab = tabBtn.dataset.pcTab === "crown" ? "crown" : "trench";
        const card = document.getElementById("gdPipeCalcCard");
        if (card) card.innerHTML = renderCalculatorHTML();
        return;
      }

      if (e.target.closest("#pcCalcTrenchBtn")) {
        calculateTrench();
        return;
      }

      if (e.target.closest("#pcCalcCrownBtn")) {
        calculateCrown();
        return;
      }

      if (e.target.closest("#gdExportPipeCsvBtn")) {
        exportCsvTable();
      }
    });
  }

  function init() {
    injectStyles();
    bindEvents();
    mountPipeCalculator();

    const obs = new MutationObserver(() => {
      mountPipeCalculator();
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
