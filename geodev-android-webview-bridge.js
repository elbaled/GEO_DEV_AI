/* ============================================================================
   GeoDev Android WebView & APK Bridge (V2.0 - Real Download + Custom Filename)
   - يتيح تعديل اسم الملف قبل الحفظ (مثل Area_1 أو Road_2) مع تثبيت الامتداد (.xlsx / .csv / .dxf / .sdr / .kml)
   - يتجاوز حظر WebView لروابط data: و blob: عبر توليد رابط تنزيل HTTP حقيقي يفتح في مدير التنزيلات
   - يدعم المشاركة المباشرة في أندرويد (واتساب / Google Drive / مدير الملفات)
   - زر تحديث فوري (Refresh) لكسر الكاش داخل التطبيق
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevAndroidBridgeV2Init) return;
  window.GeoDevAndroidBridgeV2Init = true;

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

  function getMimeByExt(ext) {
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
    return map[ext.toLowerCase()] || "application/octet-stream";
  }

  function splitFilename(rawFilename) {
    const clean = String(rawFilename || `GeoDev_${Date.now()}.txt`).trim();
    const lastDot = clean.lastIndexOf(".");
    if (lastDot <= 0) {
      return { baseName: clean, ext: "txt" };
    }
    return {
      baseName: clean.slice(0, lastDot),
      ext: clean.slice(lastDot + 1).toLowerCase()
    };
  }

  function sanitizeBaseName(name, fallback) {
    const cleaned = String(name || "")
      .trim()
      .replace(/[\\/:*?"<>|]+/g, "_")
      .replace(/\s+/g, "_");
    return cleaned || fallback;
  }

  // رفع مؤقت عبر خدمة نقل ملفات HTTP لتوليد رابط تنزيل حقيقي يقبله Android DownloadManager
  async function uploadToTempHttpDownload(blob, finalFilename) {
    const formData = new FormData();
    formData.append("file", blob, finalFilename);

    // المحاولة الأولى: tmpfiles.org (يعطي رابط تنزيل مباشر حقيقي)
    try {
      const res = await fetch("https://tmpfiles.org/api/v1/upload", {
        method: "POST",
        body: formData
      });
      if (res.ok) {
        const json = await res.json();
        const pageUrl = json?.data?.url;
        if (pageUrl) {
          // تحويل رابط العرض إلى رابط تنزيل مباشر
          return pageUrl.replace("https://tmpfiles.org/", "https://tmpfiles.org/dl/");
        }
      }
    } catch (_) {}

    // المحاولة الثانية: file.io
    try {
      const res2 = await fetch("https://file.io/?expires=1d", {
        method: "POST",
        body: formData
      });
      if (res2.ok) {
        const json2 = await res2.json();
        if (json2?.link) return json2.link;
      }
    } catch (_) {}

    return null;
  }

  // نافذة تعديل اسم الملف + التنزيل الفعلي والمشاركة في تطبيق الأندرويد
  async function handleWebViewFileExport(blob, rawFilename) {
    const { baseName, ext } = splitFilename(rawFilename);
    const mime = blob.type || getMimeByExt(ext);
    const typedBlob = new Blob([blob], { type: mime });

    const old = document.getElementById("gdAndroidExportModal");
    if (old) old.remove();

    const modal = document.createElement("div");
    modal.id = "gdAndroidExportModal";
    modal.style.cssText = `
      position: fixed; inset: 0; z-index: 999999;
      background: rgba(10, 9, 8, 0.90); backdrop-filter: blur(7px);
      display: flex; align-items: center; justify-content: center;
      padding: 16px; direction: rtl; font-family: inherit;
    `;

    modal.innerHTML = `
      <div style="width:100%;max-width:430px;background:#1e1b18;border:2px solid #e09f24;border-radius:18px;padding:18px;color:#f5f5f4;text-align:right;box-shadow:0 24px 60px rgba(0,0,0,0.85);">
        <div style="text-align:center;margin-bottom:10px;">
          <span style="display:inline-block;background:rgba(34,197,94,0.16);border:1px solid #22c55e;color:#4ade80;font-size:11.5px;font-weight:900;padding:3px 12px;border-radius:999px;">
            تجهيز الملف للحفظ والتنزيل
          </span>
        </div>

        <label style="display:block;font-size:12.5px;font-weight:900;color:#fbbf24;margin-bottom:6px;">
          اسم الملف (يمكنك تعديل الاسم مع بقاء الصيغة ثابتة):
        </label>

        <div style="display:flex;align-items:stretch;direction:ltr;border:1.5px solid #e09f24;border-radius:11px;overflow:hidden;background:#141210;margin-bottom:12px;">
          <input type="text" id="gdCustomFileBaseInput" value="${baseName}"
            placeholder="Area_1"
            style="flex:1;min-width:0;border:0;background:transparent;color:#ffffff;padding:11px 12px;font-size:14px;font-weight:800;font-family:monospace;outline:none;margin:0;">
          <span style="display:flex;align-items:center;padding:0 13px;background:#292524;color:#fbbf24;font-weight:900;font-size:14px;font-family:monospace;border-left:1px solid #44403c;user-select:none;">
            .${ext}
          </span>
        </div>

        <div id="gdAndroidDownloadStatus" style="display:none;margin-bottom:10px;padding:8px 10px;border-radius:9px;background:rgba(56,189,248,0.14);border:1px solid #38bdf8;color:#38bdf8;font-size:12px;font-weight:800;text-align:center;"></div>

        <div style="display:flex;flex-direction:column;gap:9px;">
          <button type="button" id="gdRealHttpDownloadBtn"
            style="width:100%;padding:13px;border-radius:11px;border:0;background:linear-gradient(135deg,#e09f24,#b45309);color:#141210;font-size:14px;font-weight:900;cursor:pointer;font-family:inherit;">
            تنزيل مباشر إلى مجلد Downloads
          </button>

          <button type="button" id="gdNativeShareFileBtn"
            style="width:100%;padding:12px;border-radius:11px;border:0;background:linear-gradient(135deg,#22c55e,#15803d);color:#ffffff;font-size:13.5px;font-weight:900;cursor:pointer;font-family:inherit;">
            حفظ في الهاتف أو مشاركة (واتساب / Drive / الملفات)
          </button>

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

    const nameInput = document.getElementById("gdCustomFileBaseInput");
    const statusBox = document.getElementById("gdAndroidDownloadStatus");

    const getFinalFilename = () => {
      const safeBase = sanitizeBaseName(nameInput.value, baseName);
      return `${safeBase}.${ext}`;
    };

    // 1. زر التنزيل المباشر عبر رابط HTTP حقيقي (يعمل في Android WebView و DownloadManager)
    document.getElementById("gdRealHttpDownloadBtn").onclick = async function () {
      const finalName = getFinalFilename();
      this.disabled = true;
      this.textContent = "جاري تجهيز رابط التنزيل المباشر...";
      statusBox.style.display = "block";
      statusBox.textContent = `جاري تحضير الملف (${finalName}) للتنزيل...`;

      try {
        const directUrl = await uploadToTempHttpDownload(typedBlob, finalName);
        if (directUrl) {
          statusBox.style.borderColor = "#22c55e";
          statusBox.style.color = "#4ade80";
          statusBox.innerHTML = `بدأ تنزيل <b>${finalName}</b> — إذا لم يبدأ تلقائياً <a href="${directUrl}" target="_blank" rel="noopener" style="color:#fbbf24;text-decoration:underline;">اضغط هنا للتحميل</a>`;

          // تشغيل التنزيل عبر رابط HTTPS حقيقي يتعرف عليه نظام الأندرويد فوراً
          const a = document.createElement("a");
          a.href = directUrl;
          a.download = finalName;
          a.target = "_blank";
          a.rel = "noopener";
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
        } else {
          statusBox.style.borderColor = "#f59e0b";
          statusBox.style.color = "#fbbf24";
          statusBox.textContent = "تعذر الاتصال بخادم التنزيل، استخدم زر «حفظ في الهاتف أو مشاركة» بالأسفل.";
        }
      } catch (_) {
        statusBox.textContent = "استخدم زر «حفظ في الهاتف أو مشاركة» بالأسفل لحفظ الملف مباشرة.";
      } finally {
        this.disabled = false;
        this.textContent = "تنزيل مباشر إلى مجلد Downloads";
      }
    };

    // 2. زر المشاركة والحفظ في الهاتف (Android Share Sheet)
    document.getElementById("gdNativeShareFileBtn").onclick = async () => {
      const finalName = getFinalFilename();
      const fileObj = new File([typedBlob], finalName, { type: mime });
      try {
        if (navigator.canShare && navigator.canShare({ files: [fileObj] })) {
          await navigator.share({
            files: [fileObj],
            title: finalName,
            text: `ملف ${finalName} من منصة GeoDev Studio Pro`
          });
          return;
        }
      } catch (_) {}

      // بديل تلقائي إذا لم يدعم WebView مشاركة الملفات
      document.getElementById("gdRealHttpDownloadBtn").click();
    };

    // 3. نسخ المحتوى النصي
    document.getElementById("gdCopyFileContentBtn").onclick = async () => {
      try {
        const txt = await typedBlob.text();
        await navigator.clipboard.writeText(txt);
        if (typeof toast === "function") toast("تم نسخ محتوى الملف بالكامل إلى الحافظة", "success");
      } catch (_) {
        if (typeof toast === "function") toast("هذا الملف ثنائي (Excel/ZIP)، استخدم زر التنزيل المباشر", "warn");
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
    patchSheetJSForAndroid();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
