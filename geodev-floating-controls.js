/* ============================================================================
   GeoDev Floating Top Controls (V1.0)
   - زر عائم للتبديل الفوري بين وضع الكمبيوتر (Desktop 1280px) ووضع الهاتف
   - زر عائم للتواصل المباشر عبر واتساب الإدارة (201015912371)
   - زر عائم للانتقال السريع إلى لوحة الإدارة (يظهر لحساب الأدمن تلقائياً)
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevFloatingControlsInit) return;
  window.GeoDevFloatingControlsInit = true;

  const ADMIN_EMAIL = "eng.ahmedsamir.gis@gmail.com";
  const WHATSAPP_NUMBER = "201015912371";
  const DESKTOP_MODE_KEY = "gd_force_desktop_mode_v1";

  let isDesktopMode = localStorage.getItem(DESKTOP_MODE_KEY) === "1";
  let isAdminUser = false;

  /* ============================================================================
     تطبيق أو إلغاء وضع الكمبيوتر (Desktop Viewport 1280px)
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
      meta.setAttribute("content", "width=1280, initial-scale=0.32, maximum-scale=5.0, user-scalable=yes");
      document.documentElement.classList.add("gd-desktop-view-active");
      document.body.classList.add("gd-desktop-view-active");
    } else {
      meta.setAttribute("content", "width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes");
      document.documentElement.classList.remove("gd-desktop-view-active");
      document.body.classList.remove("gd-desktop-view-active");
    }

    updateDesktopToggleButtonUI();
  }

  function updateDesktopToggleButtonUI() {
    const btn = document.getElementById("gdFloatDesktopToggleBtn");
    if (!btn) return;
    if (isDesktopMode) {
      btn.style.background = "linear-gradient(135deg, #0284c7, #1d4ed8)";
      btn.style.borderColor = "#38bdf8";
      btn.style.color = "#ffffff";
      btn.innerHTML = `
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
          <rect x="7" y="2" width="10" height="20" rx="2"/><line x1="12" y1="18" x2="12.01" y2="18"/>
        </svg>
        <span>وضع الهاتف</span>
      `;
    } else {
      btn.style.background = "rgba(28, 25, 23, 0.92)";
      btn.style.borderColor = "#e09f24";
      btn.style.color = "#fbbf24";
      btn.innerHTML = `
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
          <rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/>
        </svg>
        <span>وضع الكمبيوتر</span>
      `;
    }
  }

  /* ============================================================================
     حقن التنسيقات والأزرار العائمة في أعلى الشاشة
     ============================================================================ */
  function injectStyles() {
    if (document.getElementById("gdFloatingControlsStyles")) return;
    const style = document.createElement("style");
    style.id = "gdFloatingControlsStyles";
    style.textContent = `
      html.gd-desktop-view-active,
      body.gd-desktop-view-active {
        min-width: 1280px !important;
        width: 1280px !important;
        overflow-x: auto !important;
      }

      #gdTopFloatingCluster {
        position: fixed;
        top: 10px;
        left: 10px;
        z-index: 99990;
        display: flex;
        align-items: center;
        gap: 6px;
        direction: rtl;
        font-family: inherit;
        flex-wrap: wrap;
      }

      .gd-float-pill-btn {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        padding: 7px 11px;
        border-radius: 999px;
        border: 1.5px solid #e09f24;
        background: rgba(28, 25, 23, 0.92);
        backdrop-filter: blur(8px);
        color: #fbbf24;
        font-size: 11.5px;
        font-weight: 900;
        cursor: pointer;
        text-decoration: none;
        box-shadow: 0 6px 18px rgba(0, 0, 0, 0.45);
        transition: transform 0.18s ease, box-shadow 0.18s ease;
        white-space: nowrap;
      }
      .gd-float-pill-btn:active {
        transform: scale(0.96);
      }

      .gd-float-wa-btn {
        background: linear-gradient(135deg, #16a34a, #15803d);
        border-color: #4ade80;
        color: #ffffff;
      }

      .gd-float-admin-btn {
        background: linear-gradient(135deg, #e09f24, #b45309);
        border-color: #fde68a;
        color: #141210;
      }

      /* تكبير الأزرار قليلاً عند تفعيل وضع الكمبيوتر 1280px لتظل واضحة وسهلة اللمس */
      html.gd-desktop-view-active #gdTopFloatingCluster {
        top: 16px;
        left: 16px;
        gap: 10px;
      }
      html.gd-desktop-view-active .gd-float-pill-btn {
        padding: 11px 18px;
        font-size: 15px;
      }
    `;
    document.head.appendChild(style);
  }

  function renderFloatingBar() {
    let bar = document.getElementById("gdTopFloatingCluster");
    if (!bar) {
      bar = document.createElement("div");
      bar.id = "gdTopFloatingCluster";
      document.body.appendChild(bar);
    }

    const waText = encodeURIComponent("مرحباً م. أحمد سمير، أتواصل معك من منصة GeoDev Studio Pro");
    const waUrl = `https://wa.me/${WHATSAPP_NUMBER}?text=${waText}`;

    bar.innerHTML = `
      <button type="button" id="gdFloatDesktopToggleBtn" class="gd-float-pill-btn" title="تبديل عرض المنصة بين وضع الكمبيوتر ووضع الهاتف">
        <span>وضع الكمبيوتر</span>
      </button>

      <a href="${waUrl}" target="_blank" rel="noopener noreferrer" class="gd-float-pill-btn gd-float-wa-btn" title="تواصل مباشر مع الإدارة على واتساب">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
          <path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"/>
        </svg>
        <span>واتساب الإدارة</span>
      </a>

      ${isAdminUser ? `
        <button type="button" id="gdFloatAdminJumpBtn" class="gd-float-pill-btn gd-float-admin-btn" title="الانتقال الفوري إلى لوحة الإدارة">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
            <path d="M12 2l3 7h7l-5.5 4.5 2 7L12 16l-6.5 4.5 2-7L2 9h7z"/>
          </svg>
          <span>لوحة الإدارة</span>
        </button>
      ` : ""}
    `;

    updateDesktopToggleButtonUI();
  }

  function openAdminSection() {
    const adminBtn = Array.from(document.querySelectorAll("nav button, [data-t], button, a")).find(
      el => el.dataset?.t === "admin" || (el.textContent || "").trim() === "الإدارة"
    );
    if (adminBtn) {
      adminBtn.click();
    }

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
      if (e.target.closest("#gdFloatDesktopToggleBtn")) {
        applyViewportMode(!isDesktopMode);
        return;
      }
      if (e.target.closest("#gdFloatAdminJumpBtn")) {
        openAdminSection();
      }
    });
  }

  function watchAdminAuth() {
    const check = setInterval(() => {
      if (window.firebase && firebase.apps && firebase.apps.length) {
        clearInterval(check);
        firebase.auth().onAuthStateChanged(user => {
          isAdminUser = Boolean(user && user.email && user.email.toLowerCase() === ADMIN_EMAIL);
          renderFloatingBar();
        });
      }
    }, 450);
    setTimeout(() => clearInterval(check), 12000);
  }

  function init() {
    injectStyles();
    renderFloatingBar();
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
