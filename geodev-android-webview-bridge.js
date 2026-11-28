/* ============================================================================
   GeoDev Android WebView & APK Bridge (V1.0)
   - يحل مشكلة عدم تنزيل/حفظ الملفات (CSV / DXF / SDR / STR / XLSX / KML) داخل تطبيق الأندرويد
   - يوفر الحفظ المباشر ومشاركة الملفات عبر قائمة أندرويد الأصلية (WhatsApp / Drive / Files)
   - يضيف زر تحديث فوري (Refresh) لكسر الكاش داخل التطبيق
   - يحسن أداء WebView ويمنع الثقل أثناء معاينة نقاط الإكسيل
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevAndroidBridgeV1Init) return;
  window.GeoDevAndroidBridgeV1Init = true;

  // كشف العمل داخل تطبيق أندرويد (WebView / APK)
  function isAndroidAppWebView() {
    const ua = (navigator.userAgent || "").toLowerCase();
    const isWv = ua.includes("; wv)") || (ua.includes("android") && ua.includes("version/"));
    const isWrapper =
      ua.includes("median") ||
      ua.includes("gonative") ||
      ua.includes("appcreator") ||
      typeof window.median !== "undefined" ||
      typeof window.gonative !== "undefined" ||
      window.matchMedia("(display-mode: standalone)").matches;
    return isWv || isWrapper;
  }

  // مخزن مؤقت لربط روابط blob: بالملف الفعلي وبياناته
  const blobRegistry = new Map();
  const origCreateObjectURL = URL.createObjectURL.bind(URL);

  URL.createObjectURL = function (obj) {
    const url = origCreateObjectURL(obj);
    if (obj instanceof Blob) {
      blobRegistry.set(url, obj);
      if (blobRegistry.size > 25) {
        const firstKey = blobRegistry.keys().next().value;
        blobRegistry.delete(firstKey);
      }
    }
    return url;
  };

  function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  function getMimeByFilename(filename) {
    const ext = String(filename || "").split(".").pop().toLowerCase();
    const map = {
      csv: "text/csv",
      txt: "text/plain",
      dxf: "application/dxf",
      sdr: "text/plain",
      str: "text/plain",
      gsi: "text/plain",
      dat: "text/plain",
      scr: "text/plain",
      xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      xls: "application/vnd.ms-excel",
      geojson: "application/geo+json",
      json: "application/json",
      kml: "application/vnd.google-earth.kml+xml",
      html: "text/html",
      pdf: "application/pdf",
      zip: "application/zip"
    };
    return map[ext] || "application/octet-stream";
  }

  // نافذة حفظ ومشاركة الملف المتوافقة 100% مع تطبيق الأندرويد
  async function handleWebViewFileExport(blob, rawFilename) {
    const filename = rawFilename || `GeoDev_Export_${Date.now()}.txt`;
    const mime = blob.type || getMimeByFilename(filename);
    const typedBlob = new Blob([blob], { type: mime });
    const fileObj = new File([typedBlob], filename, { type: mime });
    const dataUrl = await blobToBase64(typedBlob);

    const old = document.getElementById("gdAndroidExportModal");
    if (old) old.remove();

    const modal = document.createElement("div");
    modal.id = "gdAndroidExportModal";
    modal.style.cssText = `
      position: fixed; inset: 0; z-index: 999999;
      background: rgba(10, 9, 8, 0.88); backdrop-filter: blur(6px);
      display: flex; align-items: center; justify-content: center;
      padding: 16px; direction: rtl; font-family: inherit;
    `;

    const canNativeShare = Boolean(navigator.canShare && navigator.canShare({ files: [fileObj] }));

    modal.innerHTML = `
      <div style="width:100%;max-width:420px;background:#1e1b18;border:2px solid #e09f24;border-radius:18px;padding:18px;color:#f5f5f4;text-align:center;box-shadow:0 24px 60px rgba(0,0,0,0.85);">
        <div style="display:inline-block;background:rgba(34,197,94,0.16);border:1px solid #22c55e;color:#4ade80;font-size:11.5px;font-weight:900;padding:3px 12px;border-radius:999px;margin-bottom:8px;">
          الملف جاهز للحفظ والمشاركة
        </div>
        <h3 style="margin:0 0 6px;font-size:16px;font-weight:900;color:#fbbf24;word-break:break-all;direction:ltr;">
          ${filename}
        </h3>
        <p style="margin:0 0 14px;font-size:12.5px;color:#d6d3d1;line-height:1.6;">
          اختر طريقة الحفظ المناسبة لتطبيق الأندرويد:
        </p>

        <div style="display:flex;flex-direction:column;gap:9px;">
          ${canNativeShare ? `
            <button type="button" id="gdNativeShareFileBtn"
              style="width:100%;padding:13px;border-radius:11px;border:0;background:linear-gradient(135deg,#22c55e,#15803d);color:#fff;font-size:14px;font-weight:900;cursor:pointer;font-family:inherit;">
              حفظ في الهاتف أو مشاركة (واتساب / Drive / مدير الملفات)
            </button>
          ` : ""}

          <a id="gdDataUrlDownloadBtn" href="${dataUrl}" download="${filename}"
            style="display:block;width:100%;box-sizing:border-box;padding:12px;border-radius:11px;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-size:13.5px;font-weight:900;text-decoration:none;">
            تنزيل مباشر إلى مجلد Downloads
          </a>

          <button type="button" id="gdCopyFileContentBtn"
            style="width:100%;padding:10px;border-radius:10px;border:1px solid #38bdf8;background:#0f172a;color:#38bdf8;font-size:12.5px;font-weight:800;cursor:pointer;font-family:inherit;">
            نسخ محتوى الملف كنص (للملفات النصية والـ CSV)
          </button>

          <button type="button" id="gdCloseAndroidExportBtn"
            style="width:100%;padding:9px;border-radius:10px;border:1px solid #44403c;background:#292524;color:#a8a29e;font-size:12px;font-weight:800;cursor:pointer;font-family:inherit;">
            إغلاق
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(modal);

    const shareBtn = document.getElementById("gdNativeShareFileBtn");
    if (shareBtn) {
      shareBtn.onclick = async () => {
        try {
          await navigator.share({
            files: [fileObj],
            title: filename,
            text: "ملف مساحي مُصدّر من منصة GeoDev Studio Pro"
          });
        } catch (_) {}
      };
    }

    document.getElementById("gdCopyFileContentBtn").onclick = async () => {
      try {
        const txt = await typedBlob.text();
        await navigator.clipboard.writeText(txt);
        if (typeof toast === "function") toast("تم نسخ محتوى الملف بالكامل إلى الحافظة", "success");
      } catch (_) {
        if (typeof toast === "function") toast("هذا الملف ثنائي (Excel/ZIP)، استخدم زر الحفظ أو المشاركة", "warn");
      }
    };

    document.getElementById("gdCloseAndroidExportBtn").onclick = () => modal.remove();
  }

  // اعتراض نقرات تحميل روابط blob: داخل تطبيق الأندرويد
  document.addEventListener(
    "click",
    e => {
      const link = e.target.closest("a[download], a[href^='blob:']");
      if (!link) return;
      const href = link.getAttribute("href") || "";
      if (!href.startsWith("blob:")) return;

      if (isAndroidAppWebView()) {
        e.preventDefault();
        e.stopPropagation();
        const blob = blobRegistry.get(href);
        const filename = link.getAttribute("download") || `GeoDev_${Date.now()}.txt`;
        if (blob) {
          handleWebViewFileExport(blob, filename);
        } else {
          fetch(href)
            .then(r => r.blob())
            .then(b => handleWebViewFileExport(b, filename))
            .catch(() => {});
        }
      }
    },
    true
  );

  // إضافة زر تحديث سريع (Refresh + مسح الكاش) في الهيدر العلوي لتطبيق الأندرويد
  function injectAppRefreshButton() {
    const header = document.querySelector("header");
    if (!header || document.getElementById("gdAppForceRefreshBtn")) return;

    const btn = document.createElement("button");
    btn.type = "button";
    btn.id = "gdAppForceRefreshBtn";
    btn.title = "تحديث التطبيق وجلب آخر التعديلات";
    btn.style.cssText = `
      background: #292524;
      color: #fbbf24;
      border: 1px solid #e09f24;
      border-radius: 9px;
      padding: 6px 10px;
      font-size: 12px;
      font-weight: 900;
      cursor: pointer;
      font-family: inherit;
      white-space: nowrap;
    `;
    btn.textContent = "⟳ تحديث";

    btn.onclick = async () => {
      btn.textContent = "⏳...";
      try {
        if ("serviceWorker" in navigator) {
          const regs = await navigator.serviceWorker.getRegistrations();
          for (const reg of regs) await reg.unregister();
        }
        if ("caches" in window) {
          const keys = await caches.keys();
          for (const k of keys) await caches.delete(k);
        }
      } catch (_) {}
      const cleanUrl = window.location.origin + window.location.pathname + "?refresh=" + Date.now();
      window.location.replace(cleanUrl);
    };

    header.appendChild(btn);
  }

  // تفعيل سحب الشاشة لأسفل للتحديث (Pull-to-Refresh)
  function enablePullToRefresh() {
    let startY = 0;
    let pulling = false;

    window.addEventListener(
      "touchstart",
      e => {
        if (window.scrollY <= 2 && e.touches.length === 1) {
          startY = e.touches[0].clientY;
          pulling = true;
        }
      },
      { passive: true }
    );

    window.addEventListener(
      "touchend",
      e => {
        if (!pulling) return;
        pulling = false;
        const endY = e.changedTouches[0]?.clientY || 0;
        if (endY - startY > 140 && window.scrollY <= 2) {
          const refreshBtn = document.getElementById("gdAppForceRefreshBtn");
          if (refreshBtn) refreshBtn.click();
          else window.location.reload();
        }
      },
      { passive: true }
    );
  }

  // معالجة تنزيلات مكتبة SheetJS (XLSX.writeFile) داخل تطبيق الأندرويد
  function patchSheetJSForAndroid() {
    if (!isAndroidAppWebView()) return;
    const iv = setInterval(() => {
      if (window.XLSX && window.XLSX.writeFile && !window.XLSX._gdPatched) {
        window.XLSX._gdPatched = true;
        const origWriteFile = window.XLSX.writeFile.bind(window.XLSX);
        window.XLSX.writeFile = function (wb, filename, opts) {
          try {
            const wbout = window.XLSX.write(wb, { bookType: "xlsx", type: "array", ...(opts || {}) });
            const blob = new Blob([wbout], {
              type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            });
            handleWebViewFileExport(blob, filename || `GeoDev_${Date.now()}.xlsx`);
          } catch (_) {
            origWriteFile(wb, filename, opts);
          }
        };
        clearInterval(iv);
      }
    }, 800);
    setTimeout(() => clearInterval(iv), 20000);
  }

  function init() {
    injectAppRefreshButton();
    enablePullToRefresh();
    patchSheetJSForAndroid();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
