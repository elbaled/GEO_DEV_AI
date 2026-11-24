/* ============================================================================
   GeoDev Guest / Visitor Mode Controller (V1.0)
   - يضيف زر "الدخول كزائر (تجربة مجانية)" في شاشة تسجيل الدخول
   - يسمح للزائر بتجربة كل أداة في المنصة مرة واحدة فقط ليرى النتيجة
   - في المحاولة الثانية لنفس الأداة، يفتح نافذة إنشاء حساب أو تسجيل الدخول
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevGuestModeV1Init) return;
  window.GeoDevGuestModeV1Init = true;

  const GUEST_ACTIVE_KEY = "gd_guest_mode_active_v1";
  const GUEST_USAGE_KEY = "gd_guest_tool_usage_v1";

  let isGuestActive = localStorage.getItem(GUEST_ACTIVE_KEY) === "1";
  let isAuthenticated = false;
  let usageMap = {};

  function loadUsage() {
    try {
      usageMap = JSON.parse(localStorage.getItem(GUEST_USAGE_KEY) || "{}");
    } catch (_) {
      usageMap = {};
    }
  }

  function saveUsage() {
    try {
      localStorage.setItem(GUEST_USAGE_KEY, JSON.stringify(usageMap));
    } catch (_) {}
  }

  /* ============================================================================
     تحديد الأداة التي يحاول الزائر استخدامها من العنصر المضغوط
     ============================================================================ */
  function detectToolFromTarget(target) {
    if (!target) return null;

    // تجاهل أزرار التنقل بين التبويبات الرئيسية وأزرار التحكم في الواجهة
    if (
      target.closest("nav button, [data-t], #gdEdgeDockWrapper, #gdGuestTopBar, #gdGuestAuthPromptModal, #gdCycleThemeBulbBtn, #gdCadZoomInBtn, #gdCadZoomOutBtn, #gdCadZoomResetBtn, [data-code-filter], [data-ex-subtab]")
    ) {
      return null;
    }

    // 1. قسم الإدارة مقفول تماماً على الزوار
    if (target.closest("#s-admin")) {
      return { id: "admin_panel", name: "لوحة الإدارة", limit: 0 };
    }

    // 2. المساعد الذكي (إرسال رسالة أو اختيار سؤال مقترح)
    if (target.closest("#sendBtn, .send-btn, #s-chat button, #s-assistant button")) {
      const btn = target.closest("button");
      if (btn && !btn.closest("#gdGuestAuthPromptModal")) {
        return { id: "ai_assistant", name: "المساعد الذكي (GeoDev AI)", limit: 1 };
      }
    }

    // 3. محطة إكسيل المساحة الذكية (تنفيذ أمر AI أو تصدير أو حساب)
    if (target.closest("#gdAiExcelStudio")) {
      if (target.closest("#gdExcelAiExecBtn, [data-ai-cmd], #gdExAutoClassifyBtn, #gdApplyShiftBtn, #gdCalcCutFillBtn, #gdParseAcadListBtn, [data-export-fmt]")) {
        return { id: "survey_excel_station", name: "محطة إكسيل المساحة الذكية", limit: 1 };
      }
      return null;
    }

    // 4. حاسبة أعمال الصرف والحفر
    if (target.closest("#gdPipeCalcCard")) {
      if (target.closest("button")) {
        return { id: "pipe_trench_calc", name: "حاسبة أعمال الصرف والحفر", limit: 1 };
      }
      return null;
    }

    // 5. مكتبة الليسبات واستوديو الأكواد البرمجية
    if (target.closest("#gdLispLibraryBox, #gdCodeStudioCard, #s-code, #s-prog")) {
      if (target.closest("button, a[download]")) {
        return { id: "code_lisp_studio", name: "استوديو ومكتبة الليسبات والأكواد", limit: 1 };
      }
      return null;
    }

    // 6. حاسبات وأدوات قسم المساحة الهندسية
    if (target.closest("#s-surv, #s-surveying")) {
      if (target.closest("button, .tool-card, .calc-card")) {
        return { id: "surveying_calculators", name: "أدوات وحاسبات المساحة الهندسية", limit: 1 };
      }
      return null;
    }

    // 7. أدوات الجيوماتكس والمختبر والأكاديمية
    if (target.closest("#s-geo, #s-lab, #s-acad, #s-academy")) {
      if (target.closest("button")) {
        return { id: "gis_academy_tools", name: "أدوات الجيوماتكس والأكاديمية", limit: 1 };
      }
    }

    return null;
  }

  /* ============================================================================
     نافذة مطالبة الزائر بإنشاء حساب أو تسجيل الدخول بعد انتهاء التجربة الواحدة
     ============================================================================ */
  function openGuestLimitModal(toolName) {
    const old = document.getElementById("gdGuestAuthPromptModal");
    if (old) old.remove();

    const overlay = document.createElement("div");
    overlay.id = "gdGuestAuthPromptModal";
    overlay.style.cssText = `
      position: fixed; inset: 0; z-index: 999999;
      background: rgba(10, 9, 8, 0.88); backdrop-filter: blur(7px);
      display: flex; align-items: center; justify-content: center;
      padding: 16px; direction: rtl; font-family: inherit;
    `;

    overlay.innerHTML = `
      <div style="width:100%;max-width:440px;background:linear-gradient(165deg,#1c1917,#141210);border:2px solid #e09f24;border-radius:18px;padding:22px;color:#f5f5f4;text-align:center;box-shadow:0 24px 60px rgba(0,0,0,0.8);">
        <div style="display:inline-block;background:rgba(224,159,36,0.16);border:1px solid #e09f24;color:#fbbf24;font-size:11.5px;font-weight:900;padding:4px 12px;border-radius:999px;margin-bottom:10px;">
          انتهت تجربة الزائر لهذه الأداة
        </div>
        <h3 style="margin:0 0 8px;font-size:18px;font-weight:900;color:#fbbf24;">
          أنشئ حساباً مجانياً للمتابعة في استخدام المنصة
        </h3>
        <p style="margin:0 0 16px;font-size:13px;color:#d6d3d1;line-height:1.7;">
          لقد قمت بتجربة <b>«${toolName}»</b> في وضع الزائر. للاستمرار في استخدام هذه الأداة وباقي أدوات <b>GeoDev Studio Pro</b> بلا حدود، يرجى إنشاء حساب مجاني أو تسجيل الدخول الآن:
        </p>
        <div style="display:flex;flex-direction:column;gap:9px;">
          <button type="button" id="gdGuestModalGoSignupBtn"
            style="width:100%;padding:13px;border-radius:11px;border:0;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-size:14px;font-weight:900;cursor:pointer;font-family:inherit;">
            إنشاء حساب مجاني / تسجيل الدخول الآن
          </button>
          <button type="button" id="gdGuestModalCloseBtn"
            style="width:100%;padding:10px;border-radius:10px;border:1px solid #44403c;background:#292524;color:#a8a29e;font-size:12.5px;font-weight:800;cursor:pointer;font-family:inherit;">
            تجربة أداة أخرى كزائر
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);

    document.getElementById("gdGuestModalCloseBtn").onclick = () => overlay.remove();
    overlay.addEventListener("click", e => { if (e.target === overlay) overlay.remove(); });

    document.getElementById("gdGuestModalGoSignupBtn").onclick = () => {
      overlay.remove();
      exitGuestToAuthScreen();
    };
  }

  /* ============================================================================
     تفعيل وضع الزائر أو العودة لشاشة تسجيل الدخول / إنشاء الحساب
     ============================================================================ */
  function getAuthContainers() {
    return Array.from(
      document.querySelectorAll("#authScreen, #loginScreen, #authOverlay, .auth-wrap, .login-wrap")
    );
  }

  function enterGuestMode() {
    isGuestActive = true;
    localStorage.setItem(GUEST_ACTIVE_KEY, "1");

    getAuthContainers().forEach(el => {
      el.dataset.gdPrevDisplay = el.style.display || "";
      el.style.display = "none";
    });

    // إظهار واجهة التطبيق الرئيسية إذا كانت مخفية
    const appMain = document.getElementById("appScreen") || document.getElementById("mainApp") || document.querySelector("main");
    if (appMain && appMain.style.display === "none") {
      appMain.style.display = "block";
    }

    renderGuestStatusBar();
  }

  function exitGuestToAuthScreen() {
    isGuestActive = false;
    localStorage.setItem(GUEST_ACTIVE_KEY, "0");

    const bar = document.getElementById("gdGuestTopBar");
    if (bar) bar.remove();

    const containers = getAuthContainers();
    if (containers.length) {
      containers.forEach(el => {
        el.style.display = el.dataset.gdPrevDisplay && el.dataset.gdPrevDisplay !== "none"
          ? el.dataset.gdPrevDisplay
          : "flex";
      });
      containers[0].scrollIntoView({ behavior: "smooth", block: "center" });
    } else {
      window.location.reload();
    }
  }

  function renderGuestStatusBar() {
    const old = document.getElementById("gdGuestTopBar");
    if (!isGuestActive || isAuthenticated) {
      if (old) old.remove();
      return;
    }
    if (old) return;

    const bar = document.createElement("div");
    bar.id = "gdGuestTopBar";
    bar.style.cssText = `
      direction: rtl; text-align: right;
      background: linear-gradient(90deg, #1c1917, #292524);
      border-bottom: 1.5px solid #e09f24;
      padding: 8px 14px; color: #f5f5f4;
      display: flex; align-items: center; justify-content: space-between;
      flex-wrap: wrap; gap: 8px; font-size: 12px; z-index: 99980; position: relative;
    `;
    bar.innerHTML = `
      <div>
        <strong style="color:#fbbf24;">وضع الزائر (ضيف):</strong>
        <span style="color:#d6d3d1;">يمكنك تجربة كل أداة في المنصة مرة واحدة مجاناً.</span>
      </div>
      <button type="button" id="gdGuestBarCreateAccountBtn"
        style="padding:6px 13px;border-radius:8px;border:0;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-weight:900;font-size:11.5px;cursor:pointer;font-family:inherit;">
        إنشاء حساب / تسجيل الدخول
      </button>
    `;
    document.body.prepend(bar);
  }

  /* ============================================================================
     حقن زر "الدخول كزائر / كضيف" داخل شاشة تسجيل الدخول تلقائياً
     ============================================================================ */
  function injectGuestLoginButton() {
    if (isAuthenticated || document.getElementById("gdEnterAsGuestBtn")) return;

    const authBox =
      document.querySelector("#authScreen .auth-card, #loginScreen .login-card, #authBox") ||
      document.getElementById("authScreen") ||
      document.getElementById("loginScreen") ||
      document.querySelector("form");

    if (!authBox) return;

    const wrap = document.createElement("div");
    wrap.id = "gdGuestEntryWrapper";
    wrap.style.cssText = `
      margin-top: 12px; padding-top: 12px;
      border-top: 1px dashed rgba(224, 159, 36, 0.4);
      text-align: center; direction: rtl;
    `;
    wrap.innerHTML = `
      <button type="button" id="gdEnterAsGuestBtn"
        style="width:100%;padding:12px 14px;border-radius:11px;border:1.5px solid #e09f24;background:rgba(224,159,36,0.12);color:#fbbf24;font-size:13.5px;font-weight:900;cursor:pointer;font-family:inherit;transition:all 0.2s ease;">
        الدخول كزائر / كضيف (تجربة كل أداة مرة واحدة مجاناً)
      </button>
      <div style="font-size:11px;color:#a8a29e;margin-top:5px;">
        تصفّح وجرّب الأدوات الهندسية فوراً بدون الحاجة لإنشاء حساب مسبق
      </div>
    `;
    authBox.appendChild(wrap);
  }

  /* ============================================================================
     مراقبة نقرات الزائر وحساب عدد مرات تجربة كل أداة (1 مرة فقط)
     ============================================================================ */
  function bindEvents() {
    document.addEventListener("click", e => {
      if (e.target.closest("#gdEnterAsGuestBtn")) {
        enterGuestMode();
        return;
      }

      if (e.target.closest("#gdGuestBarCreateAccountBtn")) {
        exitGuestToAuthScreen();
        return;
      }

      // إذا كان المستخدم مسجلاً دخوله بالفعل أو ليس في وضع الزائر، لا نتدخل
      if (isAuthenticated || !isGuestActive) return;

      const tool = detectToolFromTarget(e.target);
      if (!tool) return;

      const usedCount = Number(usageMap[tool.id] || 0);
      if (usedCount >= tool.limit) {
        // استنفد المحاولة الواحدة لهذه الأداة -> نمنع التنفيذ ونعرض نافذة إنشاء الحساب
        e.preventDefault();
        e.stopPropagation();
        e.stopImmediatePropagation();
        openGuestLimitModal(tool.name);
        return;
      }

      // السماح بأول تجربة وتسجيلها
      usageMap[tool.id] = usedCount + 1;
      saveUsage();
    }, true);
  }

  function watchAuthState() {
    const check = setInterval(() => {
      if (window.firebase && firebase.apps && firebase.apps.length) {
        clearInterval(check);
        firebase.auth().onAuthStateChanged(user => {
          isAuthenticated = Boolean(user);
          if (isAuthenticated) {
            isGuestActive = false;
            localStorage.setItem(GUEST_ACTIVE_KEY, "0");
            const bar = document.getElementById("gdGuestTopBar");
            if (bar) bar.remove();
            const entryBtn = document.getElementById("gdGuestEntryWrapper");
            if (entryBtn) entryBtn.remove();
          } else {
            injectGuestLoginButton();
            if (isGuestActive) enterGuestMode();
          }
        });
      }
    }, 400);
    setTimeout(() => clearInterval(check), 12000);
  }

  function init() {
    loadUsage();
    bindEvents();
    injectGuestLoginButton();
    if (isGuestActive) {
      enterGuestMode();
    }
    watchAuthState();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
