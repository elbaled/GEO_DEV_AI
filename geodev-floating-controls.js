/* ============================================================================
   GeoDev Smart Edge Handle & Desktop Mode Controller (V2.0)
   - مقبض سهم باهت على الحافة اليسرى للشاشة (ينفتح عند الضغط ليعرض الأزرار الثلاثة)
   - إصلاح شامل لوضع الكمبيوتر ليملأ عرض الهاتف 100% بدون فراغات أو تشويه
   - زر واتساب الإدارة + زر لوحة الإدارة للأدمن
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevFloatingControlsV20Init) return;
  window.GeoDevFloatingControlsV20Init = true;

  const ADMIN_EMAIL = "eng.ahmedsamir.gis@gmail.com";
  const WHATSAPP_NUMBER = "201015912371";
  const DESKTOP_MODE_KEY = "gd_force_desktop_mode_v2";

  let isDesktopMode = localStorage.getItem(DESKTOP_MODE_KEY) === "1";
  let isAdminUser = false;
  let isDrawerOpen = false;

  /* ============================================================================
     ضبط وضع الكمبيوتر / وضع الهاتف بشكل نظيف ومتجاوب 100%
     ============================================================================ */
  function applyViewportMode(enableDesktop) {
    isDesktopMode = Boolean(enableDesktop);
    localStorage.setItem(DESKTOP_MODE_KEY, isDesktopMode ? "1" : "0");

    let meta = document.querySelector('meta[name="viewport"]');
    if (!meta) {
      meta = document.createElement("meta");
      meta.name = "viewport";
      document.head.appendChild(meta);
    }

    if (isDesktopMode) {
      // عرض 1150px بدون initial-scale مقيد حتى يملأ المتصفح عرض الهاتف بالكامل بلا فراغ جانبي
      meta.setAttribute("content", "width=1150, maximum-scale=5.0, user-scalable=yes");
      document.documentElement.classList.add("gd-desktop-view-active");
      document.body.classList.add("gd-desktop-view-active");
    } else {
      meta.setAttribute("content", "width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes");
      document.documentElement.classList.remove("gd-desktop-view-active");
      document.body.classList.remove("gd-desktop-view-active");
    }

    renderEdgeWidget();
  }

  /* ============================================================================
     حقن التنسيقات لمقبض السهم الباهت والشريط المنبثق
     ============================================================================ */
  function injectStyles() {
    const old = document.getElementById("gdFloatingControlsStyles");
    if (old) old.remove();

    const style = document.createElement("style");
    style.id = "gdFloatingControlsStyles";
    style.textContent = `
      html.gd-desktop-view-active,
      body.gd-desktop-view-active {
        width: 100% !important;
        max-width: 100% !important;
        overflow-x: hidden !important;
      }

      /* الحاوية المثبتة على الحافة اليسرى أسفل الهيدر العلوي */
      #gdEdgeDockWrapper {
        position: fixed;
        left: 0;
        top: calc(env(safe-area-inset-top, 0px) + 74px);
        z-index: 99995;
        display: flex;
        align-items: center;
        direction: ltr;
        font-family: inherit;
        pointer-events: auto;
      }

      /* لسان السهم الباهت الملتصق بحافة الشاشة */
      #gdEdgeArrowHandle {
        width: 24px;
        height: 46px;
        border-radius: 0 12px 12px 0;
        border: 1px solid rgba(224, 159, 36, 0.45);
        border-left: 0;
        background: rgba(24, 22, 20, 0.45);
        backdrop-filter: blur(6px);
        color: rgba(251, 191, 36, 0.85);
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        box-shadow: 2px 4px 14px rgba(0, 0, 0, 0.3);
        transition: all 0.22s ease;
        padding: 0;
      }
      #gdEdgeArrowHandle:hover,
      #gdEdgeDockWrapper.open #gdEdgeArrowHandle {
        background: rgba(28, 25, 23, 0.96);
        border-color: #e09f24;
        color: #fbbf24;
        width: 28px;
      }

      /* درج الأزرار الثلاثة عند فتح السهم */
      #gdEdgeDrawerItems {
        display: flex;
        align-items: center;
        gap: 6px;
        direction: rtl;
        background: rgba(20, 18, 16, 0.95);
        backdrop-filter: blur(10px);
        border: 1.5px solid #e09f24;
        border-radius: 14px;
        padding: 6px 8px;
        margin-left: 6px;
        box-shadow: 0 12px 30px rgba(0, 0, 0, 0.65);
        transform-origin: left center;
        transition: opacity 0.2s ease, transform 0.2s ease;
      }
      #gdEdgeDockWrapper:not(.open) #gdEdgeDrawerItems {
        display: none;
      }

      .gd-edge-action-btn {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        padding: 7px 11px;
        border-radius: 10px;
        border: 1.2px solid #e09f24;
        background: #292524;
        color: #fbbf24;
        font-size: 11.5px;
        font-weight: 900;
        cursor: pointer;
        text-decoration: none;
        white-space: nowrap;
        transition: transform 0.15s ease;
      }
      .gd-edge-action-btn:active {
        transform: scale(0.95);
      }

      .gd-edge-btn-desktop-on {
        background: linear-gradient(135deg, #0284c7, #1d4ed8);
        border-color: #38bdf8;
        color: #ffffff;
      }
      .gd-edge-btn-wa {
        background: linear-gradient(135deg, #16a34a, #15803d);
        border-color: #4ade80;
        color: #ffffff;
      }
      .gd-edge-btn-admin {
        background: linear-gradient(135deg, #e09f24, #b45309);
        border-color: #fde68a;
        color: #141210;
      }

      /* تكبير مقبض السهم والأزرار في وضع الكمبيوتر (1150px) لتظل بنفس الحجم الفيزيائي الواضح على الهاتف */
      html.gd-desktop-view-active #gdEdgeDockWrapper {
        top: 130px;
      }
      html.gd-desktop-view-active #gdEdgeArrowHandle {
        width: 48px;
        height: 92px;
        border-radius: 0 22px 22px 0;
        border-width: 2px;
      }
      html.gd-desktop-view-active #gdEdgeArrowHandle svg {
        width: 26px;
        height: 26px;
      }
      html.gd-desktop-view-active #gdEdgeDrawerItems {
        padding: 12px 16px;
        gap: 12px;
        border-radius: 22px;
        border-width: 2.5px;
        margin-left: 12px;
      }
      html.gd-desktop-view-active .gd-edge-action-btn {
        padding: 14px 22px;
        font-size: 21px;
        border-radius: 16px;
        gap: 10px;
      }
      html.gd-desktop-view-active .gd-edge-action-btn svg {
        width: 22px;
        height: 22px;
      }
    `;
    document.head.appendChild(style);
  }

  function renderEdgeWidget() {
    let wrapper = document.getElementById("gdEdgeDockWrapper");
    if (!wrapper) {
      wrapper = document.createElement("div");
      wrapper.id = "gdEdgeDockWrapper";
      document.body.appendChild(wrapper);
    }

    wrapper.classList.toggle("open", isDrawerOpen);

    const waText = encodeURIComponent("مرحباً م. أحمد سمير، أتواصل معك من منصة GeoDev Studio Pro");
    const waUrl = `https://wa.me/${WHATSAPP_NUMBER}?text=${waText}`;

    const arrowSvg = isDrawerOpen
      ? `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"/></svg>`
      : `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"/></svg>`;

    const modeBtnHtml = isDesktopMode
      ? `<button type="button" id="gdFloatDesktopToggleBtn" class="gd-edge-action-btn gd-edge-btn-desktop-on">
           <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><rect x="7" y="2" width="10" height="20" rx="2"/><line x1="12" y1="18" x2="12.01" y2="18"/></svg>
           <span>وضع الهاتف</span>
         </button>`
      : `<button type="button" id="gdFloatDesktopToggleBtn" class="gd-edge-action-btn">
           <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>
           <span>وضع الكمبيوتر</span>
         </button>`;

    const adminBtnHtml = isAdminUser
      ? `<button type="button" id="gdFloatAdminJumpBtn" class="gd-edge-action-btn gd-edge-btn-admin">
           <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M12 2l3 7h7l-5.5 4.5 2 7L12 16l-6.5 4.5 2-7L2 9h7z"/></svg>
           <span>لوحة الإدارة</span>
         </button>`
      : "";

    wrapper.innerHTML = `
      <button type="button" id="gdEdgeArrowHandle" title="${isDrawerOpen ? "إخفاء الأدوات السريعة" : "إظهار الأدوات السريعة"}">
        ${arrowSvg}
      </button>

      <div id="gdEdgeDrawerItems">
        ${modeBtnHtml}
        <a href="${waUrl}" target="_blank" rel="noopener noreferrer" class="gd-edge-action-btn gd-edge-btn-wa">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"/></svg>
          <span>واتساب الإدارة</span>
        </a>
        ${adminBtnHtml}
      </div>
    `;
  }

  function openAdminSection() {
    isDrawerOpen = false;
    renderEdgeWidget();

    const adminBtn = Array.from(document.querySelectorAll("nav button, [data-t], button, a")).find(
      el => el.dataset?.t === "admin" || (el.textContent || "").trim() === "الإدارة"
    );
    if (adminBtn) adminBtn.click();

    const adminSec = document.getElementById("s-admin");
    if (adminSec) {
      document.querySelectorAll(".sec, section[id^='s-']").forEach(s => {
        if (s.id !== "s-admin") s.style.display = "none";
      });
      adminSec.style.display = "block";
      adminSec.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  function bindEvents() {
    document.addEventListener("click", e => {
      // فتح أو طي مقبض السهم الجانبي
      if (e.target.closest("#gdEdgeArrowHandle")) {
        isDrawerOpen = !isDrawerOpen;
        renderEdgeWidget();
        return;
      }

      // التبديل بين وضع الكمبيوتر ووضع الهاتف
      if (e.target.closest("#gdFloatDesktopToggleBtn")) {
        isDrawerOpen = false;
        applyViewportMode(!isDesktopMode);
        return;
      }

      // الانتقال للوحة الإدارة
      if (e.target.closest("#gdFloatAdminJumpBtn")) {
        openAdminSection();
        return;
      }

      // طي القائمة تلقائياً عند الضغط في أي مكان خارجها
      if (isDrawerOpen && !e.target.closest("#gdEdgeDockWrapper")) {
        isDrawerOpen = false;
        renderEdgeWidget();
      }
    });
  }

  function watchAdminAuth() {
    const check = setInterval(() => {
      if (window.firebase && firebase.apps && firebase.apps.length) {
        clearInterval(check);
        firebase.auth().onAuthStateChanged(user => {
          isAdminUser = Boolean(user && user.email && user.email.toLowerCase() === ADMIN_EMAIL);
          renderEdgeWidget();
        });
      }
    }, 450);
    setTimeout(() => clearInterval(check), 12000);
  }

  function init() {
    injectStyles();
    renderEdgeWidget();
    applyViewportMode(isDesktopMode);
    bindEvents();
    watchAdminAuth();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
