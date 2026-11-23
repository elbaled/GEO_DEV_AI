/* ============================================================================
   GeoDev Admin Tools & Tiers Controller (V1.2 - 3-Day Free Trial Edition)
   - يمنح المستخدم فترة تجربة مجانية 3 أيام (72 ساعة) للأدوات المدفوعة (PLUS / PRO)
   - يعرض رسالة: "معاك 3 أيام تجربة مجانية..." طوال فترة التجربة
   - بعد انتهاء الـ 3 أيام: يغلق الأداة ويعرض رسالة الترقية والاشتراك
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevAdminToolsCtrlV12Init) return;
  window.GeoDevAdminToolsCtrlV12Init = true;

  const ADMIN_EMAIL = "eng.ahmedsamir.gis@gmail.com";
  const STORAGE_KEY = "gd_custom_tools_permissions_v1";
  const TRIAL_LOCAL_KEY = "gd_tools_trials_v1";
  const THREE_DAYS_MS = 3 * 24 * 60 * 60 * 1000;

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
  let userTrialsMap = {}; // { [toolId]: startTimestampMs }
  let userTier = "free";
  let currentUid = null;
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

  function loadLocalData() {
    try {
      const savedPerm = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
      managedTools.forEach(t => {
        permissionsMap[t.id] = savedPerm[t.id] || t.defaultTier;
      });
    } catch (_) {
      managedTools.forEach(t => {
        permissionsMap[t.id] = t.defaultTier;
      });
    }

    try {
      userTrialsMap = JSON.parse(localStorage.getItem(TRIAL_LOCAL_KEY) || "{}");
    } catch (_) {
      userTrialsMap = {};
    }
  }

  async function syncTrialsFromCloud(uid) {
    if (!uid) return;
    const db = getDb();
    if (!db) return;

    try {
      const snap = await db.collection("users").doc(uid).collection("tool_trials").get();
      if (!snap.empty) {
        snap.docs.forEach(doc => {
          const d = doc.data();
          if (d && d.startedAt) {
            // نأخذ التاريخ الأقدم لمنع إعادة تعيين التجربة
            const localStart = Number(userTrialsMap[doc.id] || 0);
            userTrialsMap[doc.id] = localStart ? Math.min(localStart, Number(d.startedAt)) : Number(d.startedAt);
          }
        });
        localStorage.setItem(TRIAL_LOCAL_KEY, JSON.stringify(userTrialsMap));
      }
    } catch (_) {}
  }

  function getOrStartToolTrial(toolId) {
    const now = Date.now();
    let startedAt = Number(userTrialsMap[toolId] || 0);

    if (!startedAt) {
      startedAt = now;
      userTrialsMap[toolId] = startedAt;
      localStorage.setItem(TRIAL_LOCAL_KEY, JSON.stringify(userTrialsMap));

      // حفظ بداية التجربة في حساب المستخدم على Firestore
      if (currentUid) {
        const db = getDb();
        if (db) {
          db.collection("users").doc(currentUid).collection("tool_trials").doc(toolId).set(
            { startedAt, expiresAt: startedAt + THREE_DAYS_MS },
            { merge: true }
          ).catch(() => {});
        }
      }
    }

    const remainingMs = (startedAt + THREE_DAYS_MS) - now;
    if (remainingMs <= 0) {
      return { active: false, remainingText: "" };
    }

    const remDays = Math.floor(remainingMs / (24 * 60 * 60 * 1000));
    const remHours = Math.floor((remainingMs % (24 * 60 * 60 * 1000)) / (60 * 60 * 1000));
    const remainingText = remDays > 0
      ? `متبقي ${remDays} يوم و ${remHours} ساعة`
      : `متبقي ${Math.max(1, remHours)} ساعة`;

    return { active: true, remainingText };
  }

  async function syncPermissionsFromCloud() {
    loadLocalData();
    const db = getDb();
    if (!db) {
      enforceToolPermissions(true);
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
    enforceToolPermissions(true);
  }

  async function saveToolPermission(toolId, newTier) {
    permissionsMap[toolId] = newTier;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(permissionsMap));
    enforceToolPermissions(true);

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
    if (existingPanel && !forceRebuild) return;

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
            الأدوات المدفوعة تمنح 3 أيام تجربة مجانية تلقائياً
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
                  <option value="plus" ${cur === "plus" ? "selected" : ""}>PLUS (بلس + 3 أيام تجربة)</option>
                  <option value="pro" ${cur === "pro" ? "selected" : ""}>PRO (برو + 3 أيام تجربة)</option>
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

  function hasDirectSubscriptionAccess(requiredTier) {
    if (userTier === "admin") return true;
    if (requiredTier === "free") return true;
    if (requiredTier === "plus") return userTier === "plus" || userTier === "pro";
    if (requiredTier === "pro") return userTier === "pro";
    return false;
  }

  function enforceToolPermissions(forceRefresh = false) {
    managedTools.forEach(tool => {
      const reqTier = permissionsMap[tool.id] || tool.defaultTier;
      const targets = document.querySelectorAll(tool.selector);

      targets.forEach(el => {
        const lockId = `gdToolLockOverlay_${tool.id}`;
        const trialBannerId = `gdToolTrialBanner_${tool.id}`;
        const existingOverlay = el.querySelector(`#${lockId}`);
        const existingBanner = el.querySelector(`#${trialBannerId}`);

        // 1. إذا كان مقفولاً من الإدارة
        if (reqTier === "locked" && userTier !== "admin") {
          el.style.display = "none";
          return;
        }

        // 2. إذا كان المستخدم يملك صلاحية مباشرة (أدمن / مجاني / مشترك بالفعل)
        if (hasDirectSubscriptionAccess(reqTier)) {
          el.style.display = "";
          if (existingOverlay) existingOverlay.remove();
          if (existingBanner) existingBanner.remove();
          return;
        }

        // 3. الأداة تتطلب PLUS أو PRO والمستخدم غير مشترك -> فحص فترة التجربة المجانية 3 أيام
        const trial = getOrStartToolTrial(tool.id);
        const stateSig = `${reqTier}_${trial.active ? "trial" : "expired"}_${userTier}`;
        if (!forceRefresh && el.dataset.gdAppliedTier === stateSig) return;
        el.dataset.gdAppliedTier = stateSig;

        el.style.display = "";
        const badgeName = reqTier === "plus" ? "باقة PLUS أو PRO" : "باقة PRO";

        if (trial.active) {
          // الأداة مفتوحة خلال الـ 3 أيام التجريبية مع رسالة تنبيه في الأعلى
          if (existingOverlay) existingOverlay.remove();

          let banner = existingBanner;
          if (!banner) {
            banner = document.createElement("div");
            banner.id = trialBannerId;
            banner.style.cssText = `
              direction: rtl; text-align: right;
              background: linear-gradient(135deg, rgba(224, 159, 36, 0.18), rgba(180, 83, 9, 0.25));
              border: 1.5px solid #e09f24; border-radius: 12px;
              padding: 10px 14px; margin-bottom: 12px;
              display: flex; align-items: center; justify-content: space-between;
              flex-wrap: wrap; gap: 8px; color: #fef3c7; font-size: 13px;
            `;
            el.prepend(banner);
          }

          banner.innerHTML = `
            <div>
              <strong style="color:#fbbf24;">⏳ معاك 3 أيام تجربة مجانية لاستخدام هذه الأداة (${trial.remainingText})</strong>
              <div style="font-size:11.5px;color:#e7e5e4;margin-top:2px;">
                استمتع بكافة المميزات مجاناً الآن — بعد انتهاء الفترة التجريبية تتطلب الأداة الاشتراك في <b>${badgeName}</b>.
              </div>
            </div>
            <a href="https://wa.me/201015912371?text=${encodeURIComponent("مرحباً م. أحمد، أريد الاشتراك في " + badgeName + " على منصة GeoDev")}"
               target="_blank" rel="noopener noreferrer"
               style="padding:7px 13px;border-radius:9px;background:#e09f24;color:#141210;font-weight:900;font-size:12px;text-decoration:none;white-space:nowrap;">
              ترقية الحساب الآن
            </a>
          `;
        } else {
          // انتهت الـ 3 أيام التجريبية -> إظهار شاشة القفل والترقية
          if (existingBanner) existingBanner.remove();
          el.style.position = "relative";

          if (!existingOverlay) {
            const overlay = document.createElement("div");
            overlay.id = lockId;
            overlay.style.cssText = `
              position: absolute; inset: 0; z-index: 50;
              background: rgba(15, 23, 42, 0.94); backdrop-filter: blur(6px);
              border-radius: inherit; display: flex; flex-direction: column;
              align-items: center; justify-content: center; text-align: center;
              padding: 20px; color: #f8fafc;
            `;
            overlay.innerHTML = `
              <div style="font-size:32px;margin-bottom:8px;">🔒</div>
              <div style="font-size:12px;font-weight:900;color:#f87171;margin-bottom:4px;">
                انتهت فترة الـ 3 أيام التجريبية المجانية
              </div>
              <h4 style="margin:0 0 8px;font-size:16.5px;font-weight:900;color:#fbbf24;">
                أداة «${tool.name}» متاحة لمشتركي ${badgeName}
              </h4>
              <p style="margin:0 0 14px;font-size:12.5px;color:#cbd5e1;max-width:370px;line-height:1.7;">
                لقد استنفدت فترة التجربة المجانية (3 أيام). قم بترقية حسابك إلى <b>${badgeName}</b> للاستمرار في استخدام هذه الأداة وجميع أدوات المكتب الفني.
              </p>
              <a href="https://wa.me/201015912371?text=${encodeURIComponent("مرحباً م. أحمد، انتهت فترة التجربة وأريد تفعيل " + badgeName + " على منصة GeoDev")}"
                 target="_blank" rel="noopener noreferrer"
                 style="padding:11px 20px;border-radius:11px;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-weight:900;font-size:13.5px;text-decoration:none;">
                ⚡ تفعيل ${badgeName} الآن
              </a>
            `;
            el.appendChild(overlay);
          }
        }
      });
    });
  }

  async function detectCurrentUserTier(user) {
    currentUid = user ? user.uid : null;

    if (!user) {
      userTier = "free";
      enforceToolPermissions(true);
      return;
    }

    await syncTrialsFromCloud(user.uid);

    if (user.email && user.email.toLowerCase() === ADMIN_EMAIL) {
      userTier = "admin";
      enforceToolPermissions(true);
      return;
    }

    const db = getDb();
    if (!db) {
      userTier = "free";
      enforceToolPermissions(true);
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
    enforceToolPermissions(true);
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

    document.addEventListener("click", e => {
      if (e.target.closest("nav button, [data-t]")) {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
          renderAdminCustomToolsPanel(false);
          enforceToolPermissions(false);
        }, 250);
      }
    });
  }

  function init() {
    loadLocalData();
    bindEvents();
    setTimeout(() => {
      renderAdminCustomToolsPanel(false);
      enforceToolPermissions(true);
    }, 450);

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
