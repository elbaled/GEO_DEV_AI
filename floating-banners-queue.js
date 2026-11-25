/* ============================================================================
   GeoDev Floating Banners Queue (V2.0 - Students & Engineers 20% Discount Offer)
   - يعرض إشعار عروض خصم 20% للطلاب والمهندسين والخريجين
   - زران مخصصان للواتساب (طلاب / خريجين ومهندسين) مع إدراج اسم المستخدم تلقائياً
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevFloatingBannersV2Init) return;
  window.GeoDevFloatingBannersV2Init = true;

  const WHATSAPP_NUMBER = "201015912371";
  const DISMISS_KEY = "gd_offer_20_banner_dismissed_v2";

  function getCurrentUserName() {
    try {
      const sName = sessionStorage.getItem("gd_name");
      if (sName && sName.trim() && sName.trim() !== "ضيف") {
        return sName.trim();
      }
      if (window.firebase && firebase.apps && firebase.apps.length && firebase.auth().currentUser) {
        const u = firebase.auth().currentUser;
        if (u.displayName && u.displayName.trim()) return u.displayName.trim();
        if (u.email) return u.email.split("@")[0];
      }
    } catch (_) {}
    return "مستخدم المنصة";
  }

  function buildWhatsAppLink(category) {
    const userName = getCurrentUserName();
    const message =
      category === "students"
        ? `مرحباً م. أحمد، أنا ${userName} بستفسر عن عرض الطلاب (خصم 20%) في منصة GeoDev Studio Pro`
        : `مرحباً م. أحمد، أنا ${userName} بستفسر عن عرض الخريجين والمهندسين (خصم 20%) في منصة GeoDev Studio Pro`;

    return `https://wa.me/${WHATSAPP_NUMBER}?text=${encodeURIComponent(message)}`;
  }

  function injectStyles() {
    if (document.getElementById("gdOffer20BannerStyles")) return;
    const style = document.createElement("style");
    style.id = "gdOffer20BannerStyles";
    style.textContent = `
      #gdFloatingOfferBanner {
        position: fixed;
        top: calc(env(safe-area-inset-top, 0px) + 58px);
        left: 50%;
        transform: translateX(-50%);
        z-index: 99985;
        width: min(460px, calc(100vw - 24px));
        background: var(--cd, #1e1b18);
        border: 1.5px solid #e09f24;
        border-radius: 16px;
        padding: 12px 14px;
        box-shadow: 0 14px 36px rgba(0, 0, 0, 0.45);
        direction: rtl;
        text-align: right;
        font-family: inherit;
        color: var(--tx, #f5f5f4);
        animation: gdBannerSlideDown 0.35s ease;
      }

      @keyframes gdBannerSlideDown {
        from { opacity: 0; transform: translate(-50%, -14px); }
        to { opacity: 1; transform: translate(-50%, 0); }
      }

      .gd-ob-top {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 10px;
      }

      .gd-ob-badge {
        display: inline-block;
        background: rgba(224, 159, 36, 0.18);
        border: 1px solid #e09f24;
        color: #d97706;
        font-size: 10.5px;
        font-weight: 900;
        padding: 2px 9px;
        border-radius: 999px;
        margin-bottom: 4px;
      }

      .gd-ob-title {
        margin: 0 0 3px;
        font-size: 14.5px;
        font-weight: 900;
        color: var(--tx, #f5f5f4);
      }

      .gd-ob-sub {
        margin: 0;
        font-size: 12px;
        color: var(--mt, #a8a29e);
        line-height: 1.5;
      }

      .gd-ob-close {
        background: rgba(148, 163, 184, 0.14);
        border: 1px solid var(--ln, #38332e);
        color: var(--mt, #a8a29e);
        border-radius: 8px;
        padding: 5px 9px;
        font-size: 11.5px;
        font-weight: 800;
        cursor: pointer;
        white-space: nowrap;
        font-family: inherit;
      }

      .gd-ob-actions {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 8px;
        margin-top: 10px;
      }

      .gd-ob-wa-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
        padding: 9px 10px;
        border-radius: 10px;
        font-size: 12.5px;
        font-weight: 900;
        text-decoration: none;
        transition: transform 0.15s ease;
        white-space: nowrap;
      }
      .gd-ob-wa-btn:active {
        transform: scale(0.97);
      }

      .gd-ob-btn-students {
        background: linear-gradient(135deg, #22c55e, #16a34a);
        color: #ffffff;
        border: 1px solid #4ade80;
      }

      .gd-ob-btn-grads {
        background: linear-gradient(135deg, #e09f24, #b45309);
        color: #141210;
        border: 1px solid #fde68a;
      }
    `;
    document.head.appendChild(style);
  }

  function renderOfferBanner() {
    if (sessionStorage.getItem(DISMISS_KEY) === "1") return;
    if (document.getElementById("gdFloatingOfferBanner")) return;

    const banner = document.createElement("div");
    banner.id = "gdFloatingOfferBanner";

    const waSvg = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"/></svg>`;

    banner.innerHTML = `
      <div class="gd-ob-top">
        <div>
          <span class="gd-ob-badge">عرض الاشتراك • خصم 20%</span>
          <h4 class="gd-ob-title">عروض خصم 20% للطلاب والمهندسين</h4>
          <p class="gd-ob-sub">اسأل عن عرض الطلاب أو عرض المهندسين والخريجين</p>
        </div>
        <button type="button" id="gdCloseOfferBannerBtn" class="gd-ob-close">إغلاق ✕</button>
      </div>

      <div class="gd-ob-actions">
        <a id="gdWaStudentOfferLink" href="${buildWhatsAppLink("students")}" target="_blank" rel="noopener noreferrer" class="gd-ob-wa-btn gd-ob-btn-students">
          ${waSvg}
          <span>طلاب</span>
        </a>
        <a id="gdWaGradOfferLink" href="${buildWhatsAppLink("graduates")}" target="_blank" rel="noopener noreferrer" class="gd-ob-wa-btn gd-ob-btn-grads">
          ${waSvg}
          <span>خريجين / مهندسين</span>
        </a>
      </div>
    `;

    document.body.appendChild(banner);

    document.getElementById("gdCloseOfferBannerBtn").onclick = () => {
      sessionStorage.setItem(DISMISS_KEY, "1");
      banner.remove();
    };

    // تحديث اسم المستخدم لحظة الضغط لضمان ظهور الاسم الصحيح في رسالة الواتساب
    document.getElementById("gdWaStudentOfferLink").addEventListener("click", function () {
      this.href = buildWhatsAppLink("students");
    });
    document.getElementById("gdWaGradOfferLink").addEventListener("click", function () {
      this.href = buildWhatsAppLink("graduates");
    });
  }

  function init() {
    injectStyles();
    setTimeout(renderOfferBanner, 1200);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
