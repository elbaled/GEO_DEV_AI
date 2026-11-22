/* ============================================================================
   GeoDev Admin Tools & Tiers Controller (V1.1 - Zero-Lag Fixed Edition)
   - تم حل مشكلة التهنيج وحلقات التكرار نهائياً
   - التحكم في الأدوات: مجاني (مفتوحة) | PLUS (بلس) | PRO (برو) | مقفول
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevAdminToolsCtrlV11Init) return;
  window.GeoDevAdminToolsCtrlV11Init = true;

  const ADMIN_EMAIL = "eng.ahmedsamir.gis@gmail.com";
  const STORAGE_KEY = "gd_custom_tools_permissions_v1";

  const managedTools = [
    {
      id: "ai_survey_excel",
      name: "محطة إكسيل المساحة الذكية (AI Survey & GIS Excel)",
      category: "أدوات المساحة والـ GIS المتقدمة",
      selector: "#gdAiExcelStudio",
      defaultTier: "free"
    },
    {
      id: "pipe_trench_calc",
      name: "حاسبة أعمال الصرف والحفر (قاع الحفر وظهر الماسورة)",
      category: "أدوات المساحة والـ GIS المتقدمة",
      selector: "#gdPipeCalcCard",
      defaultTier: "free"
    },
    {
      id: "code_lisp_studio",
      name: "استوديو ومكتبة الليسبات والأكواد (.LSP / .PY)",
      category: "أدوات البرمجة والأتمتة",
      selector: "#gdCodeStudioCard, #gdLispLibraryBox",
      defaultTier: "plus"
    },
    {
      id: "arch_autocad_ai",
      name: "الأوتوكاد المعماري الذكي",
      category: "أدوات المساحة والـ GIS المتقدمة",
      selector: "[data-tool='arch-cad'], #archCadCard",
      defaultTier: "pro"
    },
    {
      id: "cert_qr_system",
      name: "استخراج الشهادات المعتمدة والتحقق بالـ QR Code",
      category: "الأكاديمية والشهادات",
      selector: "[data-generate-cert]",
      defaultTier: "free"
    }
  ];

  let permissionsMap = {};
  let userTier = "free";
  let isRendering = false;
  let debounceTimer = null;

  function getDb() {
    try {
      if (window.firebase && typeof firebase.firestore === "function") {
        return firebase.firestore();
      }
    } catch (_) {}
    return null;
  }

  function loadLocalPermissions() {
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
      managedTools.forEach(t => {
        permissionsMap[t.id] = saved[t.id] || t.defaultTier;
      });
    } catch (_) {
      managedTools.forEach(t => {
        permissionsMap[t.id] = t.defaultTier;
      });
    }
  }

  async function syncPermissionsFromCloud() {
    loadLocalPermissions();
    const db = getDb();
    if (!db) {
      enforceToolPermissions();
      return;
    }

    try {
      const snap = await db.collection("geodev_config").doc("tools_permissions").get();
      if (snap.exists && snap.data()) {
        const cloudData = snap.data();
        managedTools.forEach(t => {
          if (cloudData[t.id]) permissionsMap[t.id] = cloudData[t.id];
        });
        localStorage.setItem(STORAGE_KEY, JSON.stringify(permissionsMap));
      }
    } catch (_) {}

    renderAdminCustomToolsPanel(true);
    enforceToolPermissions();
  }

  async function saveToolPermission(toolId, newTier) {
    permissionsMap[toolId] = newTier;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(permissionsMap));
    enforceToolPermissions();

    const statusLabel = document.getElementById("gdCustomToolsSaveStatus");
    if (statusLabel) {
      statusLabel.textContent = "⏳ جاري الحفظ...";
      statusLabel.style.color = "#fbbf24";
    }

    const db = getDb();
    if (db) {
      try {
        await db.collection("geodev_config").doc("tools_permissions").set(
          {
            ...permissionsMap,
            updatedAt: Date.now()
          },
          { merge: true }
        );
        if (statusLabel) {
          statusLabel.textContent = "✅ تم الحفظ وتطبيق الصلاحية فوراً.";
          statusLabel.style.color = "#4ade80";
        }
      } catch (_) {
        if (statusLabel) {
          statusLabel.textContent = "⚠️ تم الحفظ محلياً.";
          statusLabel.style.color = "#facc15";
        }
      }
    }
  }

  function upgradeExistingAdminDropdowns() {
    const adminSec = document.getElementById("s-admin");
    if (!adminSec) return;
    const selects = adminSec.querySelectorAll("select");
    selects.forEach(sel => {
      if (sel.dataset.gdPlusUpgraded || sel.dataset.gdToolTierSelect) return;
      const optsText = Array.from(sel.options).map(o => o.textContent.trim()).join("|");
      if ((optsText.includes("مجاني") || optsText.includes("PRO")) && !optsText.includes("PLUS")) {
        sel.dataset.gdPlusUpgraded = "1";
        const plusOpt = document.createElement("option");
        plusOpt.value = "plus";
        plusOpt.textContent = "PLUS";
        const proIdx = Array.from(sel.options).findIndex(o => o.textContent.includes("PRO"));
        if (proIdx >= 0) sel.insertBefore(plusOpt, sel.options[proIdx]);
        else sel.appendChild(plusOpt);
      }
    });
  }

  function findAdminToolsContainer() {
    const adminSec = document.getElementById("s-admin");
    if (!adminSec) return null;
    const headings = Array.from(adminSec.querySelectorAll("h2, h3, h4, strong, div"));
    const targetTitle = headings.find(el =>
      el.children.length === 0 && (el.textContent || "").includes("التحكم في الأدوات والأقسام")
    );
    if (targetTitle) {
      return targetTitle.closest(".card, .sec, section, div[style]") || targetTitle.parentElement;
    }
    return adminSec;
  }

  function renderAdminCustomToolsPanel(forceRebuild = false) {
    if (isRendering) return;
    const host = findAdminToolsContainer();
    if (!host) return;

    const existingPanel = document.getElementById("gdAdminAddedToolsBox");
    if (existingPanel && !forceRebuild) {
      return; // لا تقم بإعادة البناء إذا كانت اللوحة موجودة بالفعل
    }

    isRendering = true;
    try {
      upgradeExistingAdminDropdowns();

      let panel = existingPanel;
      if (!panel) {
        panel = document.createElement("div");
        panel.id = "gdAdminAddedToolsBox";
        panel.style.cssText = `
          direction: rtl;
          text-align: right;
          margin-top: 16px;
          padding: 14px;
          border-radius: 14px;
          border: 1.5px solid #e09f24;
          background: rgba(224, 159, 36, 0.06);
        `;
        host.appendChild(panel);
      }

      const tierColors = {
        free: "#22c55e",
        plus: "#38bdf8",
        pro: "#fbbf24",
        locked: "#ef4444"
      };

      panel.innerHTML = `
        <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:6px;margin-bottom:8px;">
          <strong style="font-size:14.5px;color:#d97706;">🛠️ التحكم في الأدوات والمحطات الحديثة (مجاني / PLUS / PRO / مقفول)</strong>
          <span id="gdCustomToolsSaveStatus" style="font-size:11.5px;font-weight:800;color:#64748b;">
            اختر مستوى الوصول لكل أداة
          </span>
        </div>

        <div style="display:flex;flex-direction:column;gap:8px;">
          ${managedTools.map(tool => {
            const cur = permissionsMap[tool.id] || tool.defaultTier;
            return `
              <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;padding:9px 12px;border-radius:11px;background:var(--cd,#1e293b);border:1px solid var(--ln,#334155);">
                <div>
                  <div style="font-size:13px;font-weight:900;color:var(--tx,#f8fafc);">${tool.name}</div>
                  <div style="font-size:11px;color:var(--mt,#94a3b8);">${tool.category}</div>
                </div>
                <select data-gd-tool-tier-select="${tool.id}"
                  style="min-width:125px;padding:7px 10px;border-radius:9px;border:1.5px solid ${tierColors[cur] || "#64748b"};background:var(--bg,#0f172a);color:${tierColors[cur] || "#fff"};font-weight:900;font-size:12.5px;cursor:pointer;font-family:inherit;">
                  <option value="free" ${cur === "free" ? "selected" : ""}>مجاني (مفتوحة)</option>
                  <option value="plus" ${cur === "plus" ? "selected" : ""}>PLUS (بلس)</option>
                  <option value="pro" ${cur === "pro" ? "selected" : ""}>PRO (برو)</option>
                  <option value="locked" ${cur === "locked" ? "selected" : ""}>مقفول (مخفي)</option>
                </select>
              </div>
            `;
          }).join("")}
        </div>
      `;
    } finally {
      isRendering = false;
    }
  }

  function hasAccessToTier(requiredTier) {
    if (userTier === "admin") return true;
    if (requiredTier === "free") return true;
    if (requiredTier === "plus") return userTier === "plus" || userTier === "pro";
    if (requiredTier === "pro") return userTier === "pro";
    return false;
  }

  function enforceToolPermissions() {
    managedTools.forEach(tool => {
      const reqTier = permissionsMap[tool.id] || tool.defaultTier;
      const allowed = hasAccessToTier(reqTier);
      const targets = document.querySelectorAll(tool.selector);

      targets.forEach(el => {
        const currentApplied = el.dataset.gdAppliedTier;
        const stateSig = `${reqTier}_${allowed ? "1" : "0"}`;
        if (currentApplied === stateSig) return; // تجنب التعديل المتكرر لنفس العنصر
        el.dataset.gdAppliedTier = stateSig;

        const lockId = `gdToolLockOverlay_${tool.id}`;
        const existingOverlay = el.querySelector(`#${lockId}`);

        if (allowed) {
          el.style.display = "";
          if (existingOverlay) existingOverlay.remove();
          return;
        }

        if (reqTier === "locked") {
          el.style.display = "none";
          return;
        }

        el.style.display = "";
        el.style.position = "relative";
        if (!existingOverlay) {
          const badgeName = reqTier === "plus" ? "باقة PLUS أو PRO" : "باقة PRO";
          const overlay = document.createElement("div");
          overlay.id = lockId;
          overlay.style.cssText = `
            position: absolute; inset: 0; z-index: 50;
            background: rgba(15, 23, 42, 0.92); backdrop-filter: blur(5px);
            border-radius: inherit; display: flex; flex-direction: column;
            align-items: center; justify-content: center; text-align: center;
            padding: 18px; color: #f8fafc;
          `;
          overlay.innerHTML = `
            <div style="font-size:30px;margin-bottom:8px;">🔒</div>
            <h4 style="margin:0 0 6px;font-size:16px;font-weight:900;color:#fbbf24;">
              أداة «${tool.name}» متاحة لمشتركي ${badgeName}
            </h4>
            <p style="margin:0 0 12px;font-size:12.5px;color:#cbd5e1;max-width:360px;line-height:1.7;">
              قم بترقية حسابك إلى <b>${badgeName}</b> للاستفادة الكاملة من هذه الأداة وجميع أدوات المكتب الفني.
            </p>
            <a href="https://wa.me/201015912371?text=${encodeURIComponent("مرحباً م. أحمد، أريد تفعيل " + badgeName + " على منصة GeoDev")}"
               target="_blank" rel="noopener noreferrer"
               style="padding:10px 18px;border-radius:10px;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-weight:900;font-size:13px;text-decoration:none;">
              ⚡ تفعيل ${badgeName} الآن
            </a>
          `;
          el.appendChild(overlay);
        }
      });
    });
  }

  async function detectCurrentUserTier(user) {
    if (!user) {
      userTier = "free";
      enforceToolPermissions();
      return;
    }
    if (user.email && user.email.toLowerCase() === ADMIN_EMAIL) {
      userTier = "admin";
      enforceToolPermissions();
      return;
    }

    const db = getDb();
    if (!db) {
      userTier = "free";
      enforceToolPermissions();
      return;
    }

    try {
      const snap = await db.collection("users").doc(user.uid).get();
      const data = snap.exists ? snap.data() : {};
      const plan = String(data.plan || data.tier || data.subscription || "").toLowerCase();
      const proUntil = Number(data.pro_until || data.proUntil || 0);
      const now = Date.now();

      if ((plan.includes("pro") || data.isPro === true) && (!proUntil || proUntil > now)) {
        userTier = "pro";
      } else if (plan.includes("plus") && (!proUntil || proUntil > now)) {
        userTier = "plus";
      } else {
        userTier = "free";
      }
    } catch (_) {
      userTier = "free";
    }
    enforceToolPermissions();
  }

  function bindEvents() {
    document.addEventListener("change", e => {
      const sel = e.target.closest("[data-gd-tool-tier-select]");
      if (sel) {
        const toolId = sel.dataset.gdToolTierSelect;
        const val = sel.value;
        saveToolPermission(toolId, val);
        renderAdminCustomToolsPanel(true);
      }
    });

    // مراقب خفيف يعمل بتأخير (Debounce) عند التنقل بين الأقسام فقط بدون إرهاق المتصفح
    document.addEventListener("click", e => {
      if (e.target.closest("nav button, [data-t]")) {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
          renderAdminCustomToolsPanel(false);
          enforceToolPermissions();
        }, 250);
      }
    });
  }

  function init() {
    loadLocalPermissions();
    bindEvents();
    setTimeout(() => {
      renderAdminCustomToolsPanel(false);
      enforceToolPermissions();
    }, 400);

    const checkFb = setInterval(() => {
      if (window.firebase && firebase.apps && firebase.apps.length) {
        clearInterval(checkFb);
        syncPermissionsFromCloud();
        firebase.auth().onAuthStateChanged(user => {
          detectCurrentUserTier(user);
          renderAdminCustomToolsPanel(true);
        });
      }
    }, 500);
    setTimeout(() => clearInterval(checkFb), 12000);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
