/* ============================================================================
   GeoDev Admin Tools, Cutoff Trial & Tutorial Video Gate (V1.5)
   - يقفل الأداة فوراً على الحسابات المسجلة من يوم 22 نوفمبر 2026 فما أقدم
   - يمنح الحسابات المسجلة من 23 نوفمبر 2026 فترة تجربة يومين (48 ساعة) فقط
   - يعرض شاشة الاشتراك باسم المستخدم + نافذة فيديو شرح استخدام الأداة
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevAdminToolsCtrlV15Init) return;
  window.GeoDevAdminToolsCtrlV15Init = true;

  const ADMIN_EMAIL = "eng.ahmedsamir.gis@gmail.com";
  const STORAGE_KEY = "gd_custom_tools_permissions_v2";
  const TRIAL_LOCAL_KEY = "gd_tools_trials_v2";
  const TUTORIAL_SEEN_PREFIX = "gd_excel_tut_v1_seen_";
  const TUTORIAL_VIDEO_ID = "JrYV7iNl66I";
  const TUTORIAL_VIDEO_URL = "https://youtu.be/JrYV7iNl66I?si=w-bQyCGmFcvzyjpf";

  // الحد الفاصل: 23 نوفمبر 2026 الساعة 00:00 بتوقيت مصر (أي حساب من يوم 22 لورا مقفول فوراً)
  const CUTOFF_NOV_23_MS = new Date("2026-11-23T00:00:00+02:00").getTime();
  const TWO_DAYS_MS = 2 * 24 * 60 * 60 * 1000; // 48 ساعة

  const managedTools = [
    {
      id: "ai_survey_excel",
      name: "محطة إكسيل المساحة الذكية (AI Survey & GIS Excel)",
      category: "أدوات المساحة والـ GIS المتقدمة",
      selector: "#gdAiExcelStudio",
      defaultTier: "plus"
    },
    {
      id: "pipe_trench_calc",
      name: "حاسبة أعمال الصرف والحفر (قاع الحفر وظهر الماسورة)",
      category: "أدوات المساحة والـ GIS المتقدمة",
      selector: "#gdPipeCalcCard",
      defaultTier: "plus"
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
  let userTrialsMap = {};
  let userTier = "free";
  let currentUid = null;
  let currentUserName = "";
  let currentUserEmail = "";
  let userCreatedAtMs = 0;
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

  function resolveUserDisplayName(user, docData = {}) {
    const fromDoc = docData.name || docData.fullName || docData.displayName || "";
    const fromAuth = (user && user.displayName) || "";
    const fromSession = sessionStorage.getItem("gd_name") || "";
    const fromEmail = (user && user.email) ? user.email.split("@")[0] : "";
    const chosen = String(fromDoc || fromAuth || fromSession || fromEmail || "مهندسنا العزيز").trim();
    return chosen === "ضيف" && fromEmail ? fromEmail : chosen;
  }

  function parseTimestampMs(val) {
    if (!val) return 0;
    if (typeof val === "number") return val;
    if (typeof val.toMillis === "function") return val.toMillis();
    if (val.seconds) return val.seconds * 1000;
    const parsed = Date.parse(String(val));
    return isNaN(parsed) ? 0 : parsed;
  }

  function isLegacyBeforeNov23() {
    return userCreatedAtMs > 0 && userCreatedAtMs < CUTOFF_NOV_23_MS;
  }

  /* ============================================================================
     نافذة فيديو شرح استخدام الأداة
     ============================================================================ */
  function getTutorialStorageKey() {
    return TUTORIAL_SEEN_PREFIX + (currentUid || "guest");
  }

  function hasSeenTutorial() {
    try {
      return localStorage.getItem(getTutorialStorageKey()) === "1";
    } catch (_) {
      return false;
    }
  }

  function markTutorialSeen() {
    try {
      localStorage.setItem(getTutorialStorageKey(), "1");
    } catch (_) {}

    if (currentUid) {
      const db = getDb();
      if (db) {
        db.collection("users").doc(currentUid).set(
          { excelTutorialSeenV1: true, excelTutorialSeenAt: Date.now() },
          { merge: true }
        ).catch(() => {});
      }
    }
  }

  function openToolTutorialModal(forceShow = false) {
    if (!forceShow && hasSeenTutorial()) return;
    if (document.getElementById("gdToolTutorialGateModal")) return;

    const studioEl = document.getElementById("gdAiExcelStudio");
    if (!studioEl && !forceShow) return;

    const overlay = document.createElement("div");
    overlay.id = "gdToolTutorialGateModal";
    overlay.style.cssText = `
      position: fixed; inset: 0; z-index: 999999;
      background: rgba(10, 9, 8, 0.90); backdrop-filter: blur(7px);
      display: flex; align-items: center; justify-content: center;
      padding: 14px; direction: rtl; font-family: inherit;
    `;

    overlay.innerHTML = `
      <div style="width:100%;max-width:560px;background:linear-gradient(165deg,#1c1917,#141210);border:2px solid #e09f24;border-radius:18px;padding:18px;color:#f5f5f4;box-shadow:0 25px 65px rgba(0,0,0,0.85);">
        <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;padding-bottom:10px;border-bottom:1px solid #3f3a36;margin-bottom:12px;">
          <div>
            <span style="display:inline-block;background:rgba(224,159,36,0.18);border:1px solid #e09f24;color:#fbbf24;font-size:11px;font-weight:900;padding:2px 9px;border-radius:999px;margin-bottom:4px;">
              دليل التشغيل السريع
            </span>
            <h3 style="margin:0;font-size:17.5px;font-weight:900;color:#fbbf24;">
              شاهد فيديو كيفية استخدام الأداة
            </h3>
          </div>
          <button type="button" id="gdCloseTutorialTopX" title="إلغاء"
            style="width:34px;height:34px;border-radius:9px;border:1px solid #44403c;background:#292524;color:#d6d3d1;font-size:16px;font-weight:900;cursor:pointer;">
            ✕
          </button>
        </div>

        <p style="margin:0 0 12px;font-size:13px;color:#d6d3d1;line-height:1.65;">
          قبل البدء في استخدام <b>محطة إكسيل المساحة الذكية</b>، شاهد هذا الشرح العملي السريع لكيفية رفع الملفات وتصنيف الأكواد والتصدير إلى Civil 3D و AutoCAD، أو اضغط <b>إلغاء</b> للدخول مباشرة:
        </p>

        <div style="position:relative;width:100%;padding-top:56.25%;border-radius:12px;overflow:hidden;border:1.5px solid #38bdf8;background:#000;margin-bottom:14px;">
          <iframe
            src="https://www.youtube.com/embed/${TUTORIAL_VIDEO_ID}?rel=0&modestbranding=1"
            title="شاهد فيديو كيفية استخدام الأداة"
            style="position:absolute;top:0;left:0;width:100%;height:100%;border:0;"
            allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share"
            allowfullscreen>
          </iframe>
        </div>

        <div style="display:flex;flex-wrap:wrap;gap:9px;justify-content:space-between;align-items:center;">
          <a href="${TUTORIAL_VIDEO_URL}" target="_blank" rel="noopener noreferrer"
             style="flex:1;min-width:170px;padding:11px 14px;border-radius:11px;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-size:13.5px;font-weight:900;text-align:center;text-decoration:none;">
            ▶ فتح الفيديو على YouTube
          </a>
          <button type="button" id="gdCancelTutorialBtn"
            style="flex:1;min-width:140px;padding:11px 14px;border-radius:11px;border:1.5px solid #57534e;background:#292524;color:#f5f5f4;font-size:13.5px;font-weight:900;cursor:pointer;font-family:inherit;">
            إلغاء / متابعة للأداة
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);

    const closeAndDismiss = () => {
      markTutorialSeen();
      overlay.remove();
    };

    document.getElementById("gdCancelTutorialBtn").onclick = closeAndDismiss;
    document.getElementById("gdCloseTutorialTopX").onclick = closeAndDismiss;
  }

  function injectRewatchTutorialButton() {
    const headerActions = document.querySelector("#gdAiExcelStudio .gd-ex-header-actions");
    if (!headerActions || document.getElementById("gdRewatchTutorialBtn")) return;

    const btn = document.createElement("button");
    btn.type = "button";
    btn.id = "gdRewatchTutorialBtn";
    btn.className = "gd-theme-bulb-btn";
    btn.style.borderColor = "#38bdf8";
    btn.style.color = "#38bdf8";
    btn.textContent = "▶ شرح الأداة";
    btn.onclick = () => openToolTutorialModal(true);
    headerActions.prepend(btn);
  }

  /* ============================================================================
     إدارة الصلاحيات وفترة التجربة (قفل من 22 نوفمبر لورا / 48 ساعة من 23 نوفمبر)
     ============================================================================ */
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
            const started = Number(d.startedAt);
            const localStart = Number(userTrialsMap[doc.id] || 0);
            userTrialsMap[doc.id] = localStart ? Math.min(localStart, started) : started;
          }
        });
        localStorage.setItem(TRIAL_LOCAL_KEY, JSON.stringify(userTrialsMap));
      }
    } catch (_) {}
  }

  function getOrStartToolTrial(toolId) {
    // 1. أي حساب مسجل يوم 22 نوفمبر 2026 أو قبله: مقفول فوراً بدون فترة سماح إضافية
    if (isLegacyBeforeNov23()) {
      return { active: false, remainingText: "", reason: "legacy_cutoff" };
    }

    const now = Date.now();
    let startedAt = Number(userTrialsMap[toolId] || 0);

    // 2. الحسابات المسجلة من 23 نوفمبر فصاعداً: تحصل على 48 ساعة من وقت تسجيلها أو فتح الأداة
    if (!startedAt) {
      startedAt = userCreatedAtMs >= CUTOFF_NOV_23_MS ? userCreatedAtMs : now;
      userTrialsMap[toolId] = startedAt;
      localStorage.setItem(TRIAL_LOCAL_KEY, JSON.stringify(userTrialsMap));

      if (currentUid) {
        const db = getDb();
        if (db) {
          db.collection("users").doc(currentUid).collection("tool_trials").doc(toolId).set(
            { startedAt, expiresAt: startedAt + TWO_DAYS_MS },
            { merge: true }
          ).catch(() => {});
        }
      }
    }

    const remainingMs = (startedAt + TWO_DAYS_MS) - now;
    if (remainingMs <= 0) {
      return { active: false, remainingText: "", reason: "expired_48h" };
    }

    const remDays = Math.floor(remainingMs / (24 * 60 * 60 * 1000));
    const remHours = Math.floor((remainingMs % (24 * 60 * 60 * 1000)) / (60 * 60 * 1000));
    const remainingText = remDays > 0
      ? `متبقي ${remDays} يوم و ${remHours} ساعة`
      : `متبقي ${Math.max(1, remHours)} ساعة`;

    return { active: true, remainingText, reason: "trial_active" };
  }

  async function syncPermissionsFromCloud() {
    loadLocalData();
    const db = getDb();
    if (!db) {
      enforceToolPermissions(true);
      return;
    }

    try {
      const snap = await db.collection("geodev_config").doc("tools_permissions_v2").get();
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
        await db.collection("geodev_config").doc("tools_permissions_v2").set(
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
          <strong style="font-size:14.5px;color:#d97706;">التحكم في الأدوات والمحطات الحديثة (مجاني / PLUS / PRO / مقفول)</strong>
          <span id="gdCustomToolsSaveStatus" style="font-size:11.5px;font-weight:800;color:#64748b;">
            مقفول على المسجلين من 22 نوفمبر لورا • تجربة 48 ساعة من 23 نوفمبر
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
                  style="min-width:135px;padding:7px 10px;border-radius:9px;border:1.5px solid ${tierColors[cur] || "#64748b"};background:var(--bg,#0f172a);color:${tierColors[cur] || "#fff"};font-weight:900;font-size:12.5px;cursor:pointer;font-family:inherit;">
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

  function hasDirectSubscriptionAccess(requiredTier) {
    if (userTier === "admin") return true;
    if (requiredTier === "free") return true;
    if (requiredTier === "plus") return userTier === "plus" || userTier === "pro";
    if (requiredTier === "pro") return userTier === "pro";
    return false;
  }

  function enforceToolPermissions(forceRefresh = false) {
    injectRewatchTutorialButton();

    // الزائر قبل تسجيل الدخول يبقى في وضع المعاينة الخاص بمحطة الإكسيل
    if (!currentUid) return;

    const displayName = currentUserName || resolveUserDisplayName(null);

    managedTools.forEach(tool => {
      const reqTier = permissionsMap[tool.id] || tool.defaultTier;
      const targets = document.querySelectorAll(tool.selector);

      targets.forEach(el => {
        const lockId = `gdToolLockOverlay_${tool.id}`;
        const trialBannerId = `gdToolTrialBanner_${tool.id}`;
        const existingOverlay = el.querySelector(`#${lockId}`);
        const existingBanner = el.querySelector(`#${trialBannerId}`);

        if (reqTier === "locked" && userTier !== "admin") {
          el.style.display = "none";
          return;
        }

        if (hasDirectSubscriptionAccess(reqTier)) {
          el.style.display = "";
          if (existingOverlay) existingOverlay.remove();
          if (existingBanner) existingBanner.remove();
          return;
        }

        const trial = getOrStartToolTrial(tool.id);
        const stateSig = `${reqTier}_${trial.active ? "trial2d" : "locked"}_${userTier}_${displayName}`;
        if (!forceRefresh && el.dataset.gdAppliedTier === stateSig) return;
        el.dataset.gdAppliedTier = stateSig;

        el.style.display = "";
        const badgeName = reqTier === "plus" ? "باقة PLUS أو PRO" : "باقة PRO";

        if (trial.active) {
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
              <strong style="color:#fbbf24;">أهلاً ${displayName} — معاك يومين تجربة مجانية (${trial.remainingText})</strong>
              <div style="font-size:11.5px;color:#e7e5e4;margin-top:2px;">
                استمتع بكافة المميزات مجاناً الآن — بعد انتهاء الفترة التجريبية تتطلب الأداة الاشتراك في <b>${badgeName}</b>.
              </div>
            </div>
            <a href="https://wa.me/201015912371?text=${encodeURIComponent("مرحباً م. أحمد، أنا " + displayName + " (" + currentUserEmail + ") وأريد الاشتراك في " + badgeName + " على منصة GeoDev")}"
               target="_blank" rel="noopener noreferrer"
               style="padding:7px 13px;border-radius:9px;background:#e09f24;color:#141210;font-weight:900;font-size:12px;text-decoration:none;white-space:nowrap;">
              ترقية الحساب الآن
            </a>
          `;
        } else {
          if (existingBanner) existingBanner.remove();
          el.style.position = "relative";
          el.style.minHeight = "270px";

          if (existingOverlay) existingOverlay.remove();

          const overlay = document.createElement("div");
          overlay.id = lockId;
          overlay.style.cssText = `
            position: absolute; inset: 0; z-index: 50;
            background: rgba(15, 14, 12, 0.95); backdrop-filter: blur(7px);
            border-radius: inherit; border: 2px solid #e09f24;
            display: flex; flex-direction: column;
            align-items: center; justify-content: center; text-align: center;
            padding: 22px 16px; color: #f5f5f4;
          `;
          overlay.innerHTML = `
            <div style="display:inline-block;background:rgba(224,159,36,0.16);border:1.5px solid #e09f24;color:#fbbf24;font-size:12px;font-weight:900;padding:4px 14px;border-radius:999px;margin-bottom:10px;">
              الحساب: ${displayName}
            </div>
            <div style="font-size:13px;font-weight:900;color:#f87171;margin-bottom:5px;">
              انتهت الفترة التجريبية المجانية لحسابك
            </div>
            <h4 style="margin:0 0 8px;font-size:17px;font-weight:900;color:#fbbf24;">
              أداة «${tool.name}» متاحة لمشتركي ${badgeName}
            </h4>
            <p style="margin:0 0 16px;font-size:13px;color:#d6d3d1;max-width:390px;line-height:1.7;">
              مرحباً <b>${displayName}</b>، للاستمرار في استخدام هذه الأداة وجميع أدوات المكتب الفني والمساحة، يرجى تفعيل اشتراكك في <b>${badgeName}</b> الآن:
            </p>
            <div style="display:flex;flex-wrap:wrap;gap:9px;justify-content:center;">
              <a href="https://wa.me/201015912371?text=${encodeURIComponent("مرحباً م. أحمد، أنا " + displayName + " (" + currentUserEmail + ") وأريد تفعيل اشتراك " + badgeName + " على منصة GeoDev")}"
                 target="_blank" rel="noopener noreferrer"
                 style="padding:11px 22px;border-radius:11px;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-weight:900;font-size:13.5px;text-decoration:none;">
                تفعيل ${badgeName} عبر واتساب
              </a>
              <button type="button" onclick="if(typeof showPro==='function')showPro('lock');"
                style="padding:11px 16px;border-radius:11px;border:1.5px solid #e09f24;background:#292524;color:#fbbf24;font-weight:900;font-size:13px;cursor:pointer;font-family:inherit;">
                إدخال كود تفعيل
              </button>
            </div>
          `;
          el.appendChild(overlay);
        }
      });
    });
  }

  async function detectCurrentUserTier(user) {
    currentUid = user ? user.uid : null;
    currentUserEmail = (user && user.email) || sessionStorage.getItem("gd_email") || "";

    if (!user) {
      userTier = "free";
      userCreatedAtMs = 0;
      currentUserName = "";
      enforceToolPermissions(true);
      setTimeout(() => openToolTutorialModal(false), 600);
      return;
    }

    const authCreationMs = user.metadata && user.metadata.creationTime
      ? parseTimestampMs(user.metadata.creationTime)
      : 0;

    userCreatedAtMs = authCreationMs;
    currentUserName = resolveUserDisplayName(user);

    await syncTrialsFromCloud(user.uid);

    const db = getDb();
    if (db) {
      try {
        const snap = await db.collection("users").doc(user.uid).get();
        const data = snap.exists ? snap.data() : {};

        currentUserName = resolveUserDisplayName(user, data);

        const docCreatedMs = parseTimestampMs(data.created || data.createdAt || data.registeredAt);
        if (docCreatedMs > 0) {
          userCreatedAtMs = userCreatedAtMs > 0 ? Math.min(userCreatedAtMs, docCreatedMs) : docCreatedMs;
        }

        if (data.excelTutorialSeenV1 === true) {
          localStorage.setItem(getTutorialStorageKey(), "1");
        }

        if (user.email && user.email.toLowerCase() === ADMIN_EMAIL) {
          userTier = "admin";
        } else {
          const plan = String(data.plan || data.tier || data.subscription || "").toLowerCase();
          const proUntil = parseTimestampMs(data.pro_until || data.proUntil);
          const now = Date.now();

          if ((plan.includes("pro") || data.isPro === true) && (!proUntil || proUntil > now)) {
            userTier = "pro";
          } else if (plan.includes("plus") && (!proUntil || proUntil > now)) {
            userTier = "plus";
          } else {
            userTier = "free";
          }
        }
      } catch (_) {
        userTier = user.email && user.email.toLowerCase() === ADMIN_EMAIL ? "admin" : "free";
      }
    } else {
      userTier = user.email && user.email.toLowerCase() === ADMIN_EMAIL ? "admin" : "free";
    }

    enforceToolPermissions(true);
    setTimeout(() => openToolTutorialModal(false), 650);
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
          enforceToolPermissions(true);
          const survSec = document.getElementById("s-survey") || document.getElementById("s-surv");
          if (survSec && survSec.offsetParent !== null) {
            openToolTutorialModal(false);
          }
        }, 300);
      }
    });
  }

  function init() {
    loadLocalData();
    bindEvents();
    setTimeout(() => {
      renderAdminCustomToolsPanel(false);
      enforceToolPermissions(true);
    }, 550);

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
  // تقليص أي فترة تجريبية قادمة من السيرفر إلى يومين (48 ساعة) من تاريخ التسجيل
  const _origMe = window.me;
  if (typeof _origMe === "function") {
    window.me = function () {
      const u = _origMe.apply(this, arguments);
      if (u && u.trialEnd && u.created) {
        const maxTwoDays = Number(u.created) + TWO_DAYS_MS;
        if (u.trialEnd > maxTwoDays) u.trialEnd = maxTwoDays;
      }
      return u;
    };
  }

})();
