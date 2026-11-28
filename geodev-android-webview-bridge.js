/* ============================================================================
   GeoDev Direct Device File Saver & Permission Bridge (V3.0 - Zero External Links)
   - بدون أي مواقع وسيطة أو روابط خارجية نهائياً
   - يطلب إذن حفظ الملفات على ذاكرة الهاتف مباشرة عند فتح التطبيق (File System Access API)
   - يتيح تعديل اسم الملف مع بقاء الصيغة ثابتة (.xlsx / .csv / .dxf / .sdr / .kml)
   - يحفظ الملف مباشرة داخل الهاتف أو عبر قائمة حفظ أندرويد الأصلية
   ============================================================================ */
(function () {
  "use strict";
  if (window.GeoDevAndroidBridgeV3Init) return;
  window.GeoDevAndroidBridgeV3Init = true;

  const STORAGE_PERM_KEY = "gd_device_storage_perm_granted_v1";
  let savedDirectoryHandle = null;

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
    if (lastDot <= 0) return { baseName: clean, ext: "txt" };
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

  /* ============================================================================
     1. رسالة طلب إذن حفظ الملفات على ذاكرة الهاتف فور فتح التطبيق
     ============================================================================ */
  function promptDeviceStoragePermissionOnStartup() {
    if (!isAndroidAppWebView()) return;
    if (localStorage.getItem(STORAGE_PERM_KEY) === "1") return;
    if (document.getElementById("gdStoragePermDialog")) return;

    const modal = document.createElement("div");
    modal.id = "gdStoragePermDialog";
    modal.style.cssText = `
      position: fixed; inset: 0; z-index: 999999;
      background: rgba(10, 9, 8, 0.85); backdrop-filter: blur(6px);
      display: flex; align-items: center; justify-content: center;
      padding: 16px; direction: rtl; font-family: inherit;
    `;

    modal.innerHTML = `
      <div style="width:100%;max-width:400px;background:#1e1b18;border:2px solid #e09f24;border-radius:18px;padding:20px;color:#f5f5f4;text-align:center;box-shadow:0 24px 60px rgba(0,0,0,0.85);">
        <div style="width:48px;height:48px;margin:0 auto 10px;border-radius:14px;background:rgba(224,159,36,0.16);border:1.5px solid #e09f24;display:grid;place-items:center;color:#fbbf24;font-weight:900;font-size:18px;">
          DIR
        </div>
        <h3 style="margin:0 0 8px;font-size:16.5px;font-weight:900;color:#fbbf24;">
          السماح بحفظ الملفات في ذاكرة الهاتف
        </h3>
        <p style="margin:0 0 16px;font-size:13px;color:#d6d3d1;line-height:1.7;">
          هل تريد السماح لتطبيق <b>GeoDev Studio Pro</b> بحفظ ملفات الإكسيل والأوتوكاد والتوتال ستيشن مباشرة داخل مجلدات هاتفك؟
        </p>
        <div style="display:flex;gap:9px;">
          <button type="button" id="gdAllowDeviceStorageBtn"
            style="flex:1;padding:12px;border-radius:11px;border:0;background:linear-gradient(135deg,#22c55e,#15803d);color:#fff;font-size:14px;font-weight:900;cursor:pointer;font-family:inherit;">
            موافق (السماح بالحفظ)
          </button>
          <button type="button" id="gdDenyDeviceStorageBtn"
            style="padding:12px 16px;border-radius:11px;border:1px solid #44403c;background:#292524;color:#a8a29e;font-size:13px;font-weight:800;cursor:pointer;font-family:inherit;">
            لاحقاً
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(modal);

    document.getElementById("gdAllowDeviceStorageBtn").onclick = async () => {
      localStorage.setItem(STORAGE_PERM_KEY, "1");
      modal.remove();
      try {
        if (typeof window.showDirectoryPicker === "function") {
          savedDirectoryHandle = await window.showDirectoryPicker({ mode: "readwrite" });
        }
      } catch (_) {}
      if (typeof toast === "function") {
        toast("تم تفعيل الحفظ المباشر على الهاتف بنجاح", "success");
      }
    };

    document.getElementById("gdDenyDeviceStorageBtn").onclick = () => {
      localStorage.setItem(STORAGE_PERM_KEY, "1");
      modal.remove();
    };
  }

  /* ============================================================================
     2. حفظ الملف مباشرة داخل الهاتف بدون مغادرة التطبيق أو فتح أي موقع خارجي
     ============================================================================ */
  async function saveBlobDirectlyToDevice(typedBlob, finalName, mime, statusBox) {
    // المسار أ: الكتابة المباشرة في المجلد المصرح به على الهاتف (إذا كان مدعوماً)
    if (savedDirectoryHandle) {
      try {
        const fileHandle = await savedDirectoryHandle.getFileHandle(finalName, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(typedBlob);
        await writable.close();
        statusBox.style.display = "block";
        statusBox.style.borderColor = "#22c55e";
        statusBox.style.color = "#4ade80";
        statusBox.innerHTML = `تم حفظ الملف <b>${finalName}</b> مباشرة داخل مجلد الهاتف بنجاح.`;
        return true;
      } catch (_) {}
    }

    // المسار ب: نافذة "حفظ باسم" في نظام أندرويد (showSaveFilePicker)
    if (typeof window.showSaveFilePicker === "function") {
      try {
        const { ext } = splitFilename(finalName);
        const handle = await window.showSaveFilePicker({
          suggestedName: finalName,
          types: [{ description: "GeoDev Survey File", accept: { [mime]: ["." + ext] } }]
        });
        const writable = await handle.createWritable();
        await writable.write(typedBlob);
        await writable.close();
        statusBox.style.display = "block";
        statusBox.style.borderColor = "#22c55e";
        statusBox.style.color = "#4ade80";
        statusBox.innerHTML = `تم حفظ <b>${finalName}</b> في ذاكرة الهاتف بنجاح.`;
        return true;
      } catch (err) {
        if (err && err.name === "AbortError") return true;
      }
    }

    // المسار ج: واجهة تطبيق أندرويد المدمجة (إن وجدت في APK)
    if (window.Android && typeof window.Android.saveFile === "function") {
      try {
        const reader = new FileReader();
        reader.onloadend = () => {
          window.Android.saveFile(reader.result, finalName, mime);
        };
        reader.readAsDataURL(typedBlob);
        statusBox.style.display = "block";
        statusBox.style.borderColor = "#22c55e";
        statusBox.style.color = "#4ade80";
        statusBox.innerHTML = `تم إرسال <b>${finalName}</b> إلى مجلد التنزيلات في الهاتف.`;
        return true;
      } catch (_) {}
    }

    // المسار د: قائمة الحفظ والمشاركة الأصلية في نظام أندرويد (Save to Files / Drive / WhatsApp)
    const fileObj = new File([typedBlob], finalName, { type: mime });
    if (navigator.canShare && navigator.canShare({ files: [fileObj] })) {
      try {
        await navigator.share({
          files: [fileObj],
          title: finalName
        });
        statusBox.style.display = "block";
        statusBox.style.borderColor = "#22c55e";
        statusBox.style.color = "#4ade80";
        statusBox.innerHTML = `تم تجهيز <b>${finalName}</b> — اختر «حفظ في الملفات» أو التطبيق المطلوب.`;
        return true;
      } catch (err) {
        if (err && err.name === "AbortError") return true;
      }
    }

    // المسار هـ: التنزيل المحلي المباشر للمتصفح (blob:)
    const localBlobUrl = origCreateObjectURL(typedBlob);
    const a = document.createElement("a");
    a.href = localBlobUrl;
    a.download = finalName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(localBlobUrl), 3000);

    statusBox.style.display = "block";
    statusBox.style.borderColor = "#fbbf24";
    statusBox.style.color = "#fbbf24";
    statusBox.innerHTML = `تم طلب حفظ <b>${finalName}</b> محلياً على الجهاز.`;
    return true;
  }

  /* ============================================================================
     3. نافذة تعديل اسم الملف (مع تثبيت الصيغة) والحفظ المباشر على الهاتف
     ============================================================================ */
  function handleWebViewFileExport(blob, rawFilename) {
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
            حفظ الملف مباشرة في الهاتف
          </span>
        </div>

        <label style="display:block;font-size:12.5px;font-weight:900;color:#fbbf24;margin-bottom:6px;">
          اكتب اسم الملف (الصيغة ثابتة تلقائياً):
        </label>

        <div style="display:flex;align-items:stretch;direction:ltr;border:1.5px solid #e09f24;border-radius:11px;overflow:hidden;background:#141210;margin-bottom:12px;">
          <input type="text" id="gdCustomFileBaseInput" value="${baseName}"
            placeholder="Area_1"
            style="flex:1;min-width:0;border:0;background:transparent;color:#ffffff;padding:11px 12px;font-size:14px;font-weight:800;font-family:monospace;outline:none;margin:0;">
          <span style="display:flex;align-items:center;padding:0 13px;background:#292524;color:#fbbf24;font-weight:900;font-size:14px;font-family:monospace;border-left:1px solid #44403c;user-select:none;">
            .${ext}
          </span>
        </div>

        <div id="gdAndroidDownloadStatus" style="display:none;margin-bottom:10px;padding:9px 10px;border-radius:9px;background:rgba(34,197,94,0.14);border:1px solid #22c55e;color:#4ade80;font-size:12px;font-weight:800;text-align:center;"></div>

        <div style="display:flex;flex-direction:column;gap:9px;">
          <button type="button" id="gdDirectSaveToPhoneBtn"
            style="width:100%;padding:13px;border-radius:11px;border:0;background:linear-gradient(135deg,#22c55e,#15803d);color:#ffffff;font-size:14.5px;font-weight:900;cursor:pointer;font-family:inherit;">
            حفظ في الهاتف مباشرة
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

    document.getElementById("gdDirectSaveToPhoneBtn").onclick = async function () {
      const finalName = getFinalFilename();
      this.disabled = true;
      this.textContent = "جاري الحفظ في الهاتف...";
      await saveBlobDirectlyToDevice(typedBlob, finalName, mime, statusBox);
      this.disabled = false;
      this.textContent = "حفظ في الهاتف مباشرة";
    };

    document.getElementById("gdCopyFileContentBtn").onclick = async () => {
      try {
        const txt = await typedBlob.text();
        await navigator.clipboard.writeText(txt);
        if (typeof toast === "function") toast("تم نسخ محتوى الملف بالكامل إلى الحافظة", "success");
      } catch (_) {
        if (typeof toast === "function") toast("هذا الملف ثنائي (Excel)، استخدم زر «حفظ في الهاتف مباشرة»", "warn");
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

  // زر تحديث سريع (Refresh + مسح الكاش) في الهيدر العلوي
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
    setTimeout(promptDeviceStoragePermissionOnStartup, 1500);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
