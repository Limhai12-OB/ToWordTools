/* ToWord front-end: upload → convert → edit (TinyMCE) → download. */
(() => {
  "use strict";

  // ------------------------------------------------------------------ i18n
  const T = {
    en: {
      tagline: "PDF → editable Word · Khmer supported",
      heroTitle: "Convert PDF to Word and edit it right here",
      heroSub: "Works with Khmer and English documents, including scanned PDFs (OCR).",
      dropTitle: "Drop your file here, or click to choose",
      dropSub: "PDF, images (JPG/PNG), DOCX, DOC, ODT, RTF · up to 60 MB",
      options: "Conversion options", modeLabel: "Mode",
      modeAuto: "Automatic", modeAutoHint: "Picks layout or OCR for you",
      modeLayout: "Keep layout", modeLayoutHint: "Best for digital PDFs: keeps tables, images, positions",
      modeOcr: "OCR (read text from image)", modeOcrHint: "For scanned PDFs (plain text, layout is not kept)",
      modeHybrid: "Keep layout + fix Khmer with OCR", modeHybridHint: "For Khmer PDFs made in Word: keeps the form layout, re-reads Khmer text from the page",
      modeHybridShort: "Best for Khmer PDFs from Word", modeLayoutShort: "Use the PDF's own text", modeOcrShort: "Plain text, for scanned pages",
      reconvert: "Convert again", modeBadge_hybrid: "Layout + Khmer OCR",
      noticeHybrid: "This PDF's Khmer text was broken, so it was re-read from the page with OCR while keeping the layout. Please proofread names and numbers.",
      langLabel: "OCR language", langBoth: "Khmer + English", langKhm: "Khmer only", langEng: "English only",
      fontLabel: "Khmer font in Word",
      resumeText: "Continue editing your last document?", resumeBtn: "Open",
      f1t: "Khmer-ready", f1: "Repairs broken Khmer text from PDFs and sets proper Khmer fonts for Word.",
      f2t: "Scanned PDFs", f2: "Built-in Khmer + English OCR turns page images into editable text.",
      f3t: "Edit, then download", f3: "Edit in the browser and download as Word (.docx) or PDF.",
      tryAnother: "Try another file", compare: "Show original", hideCompare: "Hide original",
      reconvertOcr: "Try OCR", reconvertLayout: "Try layout mode",
      download: "Download", dlDocx: "Word (.docx) — with my edits", dlPdf: "PDF — with my edits",
      dlOrig: "Word (.docx) — original conversion", dlOrigHint: "Exact layout, without browser edits",
      newFile: "New file", uploading: "Uploading…", saving: "Saving…", saved: "All changes saved",
      saveFailed: "Could not save draft", preparing: "Preparing download…",
      modeBadge_layout: "Layout mode", modeBadge_ocr: "OCR", modeBadge_docx: "Word file", modeBadge_office: "Converted",
      noticeOcr: "Text was read with OCR. Please proofread — especially Khmer subscripts and vowels.",
      noticeLayout: "If any Khmer text looks wrong (broken vowels or missing letters), try OCR mode.",
      noticeAutoOcr: "This PDF's text layer was unusable ({why}), so OCR was used. Please proofread.",
      unsupported: "Unsupported file type.", failed: "Conversion failed",
      confirmNew: "Start over with a new file? Your edits to this one are saved as a draft.",
    },
    km: {
      tagline: "PDF → Word ដែលអាចកែបាន · គាំទ្រភាសាខ្មែរ",
      heroTitle: "បម្លែង PDF ទៅជា Word ហើយកែសម្រួលនៅទីនេះ",
      heroSub: "ដំណើរការជាមួយឯកសារភាសាខ្មែរ និងអង់គ្លេស រួមទាំង PDF ដែលស្កេន (OCR)។",
      dropTitle: "ទម្លាក់ឯកសារនៅទីនេះ ឬចុចដើម្បីជ្រើសរើស",
      dropSub: "PDF, រូបភាព (JPG/PNG), DOCX, DOC, ODT, RTF · ធំបំផុត 60 MB",
      options: "ជម្រើសបម្លែង", modeLabel: "របៀប",
      modeAuto: "ស្វ័យប្រវត្តិ", modeAutoHint: "ជ្រើសរើសប្លង់ ឬ OCR ដោយស្វ័យប្រវត្តិ",
      modeLayout: "រក្សាប្លង់", modeLayoutHint: "ល្អសម្រាប់ PDF ឌីជីថល៖ រក្សាតារាង រូបភាព ទីតាំង",
      modeOcr: "OCR (អានអក្សរពីរូបភាព)", modeOcrHint: "សម្រាប់ PDF ស្កេន (អក្សរធម្មតា មិនរក្សាប្លង់)",
      modeHybrid: "រក្សាប្លង់ + កែអក្សរខ្មែរដោយ OCR", modeHybridHint: "សម្រាប់ PDF ខ្មែរដែលធ្វើពី Word៖ រក្សាប្លង់ទម្រង់ ហើយអានអក្សរខ្មែរឡើងវិញពីទំព័រ",
      modeHybridShort: "ល្អបំផុតសម្រាប់ PDF ខ្មែរពី Word", modeLayoutShort: "ប្រើអក្សរដើមរបស់ PDF", modeOcrShort: "អក្សរធម្មតា សម្រាប់ទំព័រស្កេន",
      reconvert: "បម្លែងម្តងទៀត", modeBadge_hybrid: "ប្លង់ + OCR ខ្មែរ",
      noticeHybrid: "អក្សរខ្មែរក្នុង PDF នេះខូច ដូច្នេះត្រូវបានអានឡើងវិញដោយ OCR ដោយរក្សាប្លង់។ សូមពិនិត្យឈ្មោះ និងលេខឡើងវិញ។",
      langLabel: "ភាសា OCR", langBoth: "ខ្មែរ + អង់គ្លេស", langKhm: "ខ្មែរតែប៉ុណ្ណោះ", langEng: "អង់គ្លេសតែប៉ុណ្ណោះ",
      fontLabel: "ពុម្ពអក្សរខ្មែរក្នុង Word",
      resumeText: "បន្តកែសម្រួលឯកសារចុងក្រោយរបស់អ្នក?", resumeBtn: "បើក",
      f1t: "គាំទ្រខ្មែរ", f1: "ជួសជុលអក្សរខ្មែរដែលខូចពី PDF និងកំណត់ពុម្ពអក្សរខ្មែរត្រឹមត្រូវសម្រាប់ Word។",
      f2t: "PDF ស្កេន", f2: "OCR ខ្មែរ + អង់គ្លេស បម្លែងរូបភាពទំព័រទៅជាអក្សរដែលអាចកែបាន។",
      f3t: "កែ រួចទាញយក", f3: "កែសម្រួលក្នុងកម្មវិធីរុករក ហើយទាញយកជា Word (.docx) ឬ PDF។",
      tryAnother: "សាកល្បងឯកសារផ្សេង", compare: "បង្ហាញឯកសារដើម", hideCompare: "លាក់ឯកសារដើម",
      reconvertOcr: "សាកល្បង OCR", reconvertLayout: "សាកល្បងរក្សាប្លង់",
      download: "ទាញយក", dlDocx: "Word (.docx) — ជាមួយការកែសម្រួល", dlPdf: "PDF — ជាមួយការកែសម្រួល",
      dlOrig: "Word (.docx) — លទ្ធផលបម្លែងដើម", dlOrigHint: "ប្លង់ត្រឹមត្រូវ គ្មានការកែសម្រួល",
      newFile: "ឯកសារថ្មី", uploading: "កំពុងផ្ទុកឡើង…", saving: "កំពុងរក្សាទុក…", saved: "បានរក្សាទុកការផ្លាស់ប្តូរទាំងអស់",
      saveFailed: "មិនអាចរក្សាទុកបាន", preparing: "កំពុងរៀបចំឯកសារ…",
      modeBadge_layout: "រក្សាប្លង់", modeBadge_ocr: "OCR", modeBadge_docx: "ឯកសារ Word", modeBadge_office: "បានបម្លែង",
      noticeOcr: "អក្សរត្រូវបានអានដោយ OCR។ សូមពិនិត្យឡើងវិញ ជាពិសេសជើងអក្សរ និងស្រៈ។",
      noticeLayout: "ប្រសិនបើអក្សរខ្មែរមើលទៅខុស (ស្រៈខូច ឬបាត់អក្សរ) សូមសាកល្បង OCR។",
      noticeAutoOcr: "អក្សរក្នុង PDF នេះមិនអាចប្រើបាន ({why}) ដូច្នេះបានប្រើ OCR។ សូមពិនិត្យឡើងវិញ។",
      unsupported: "មិនគាំទ្រប្រភេទឯកសារនេះទេ។", failed: "ការបម្លែងបរាជ័យ",
      confirmNew: "ចាប់ផ្តើមឯកសារថ្មី? ការកែសម្រួលនេះត្រូវបានរក្សាទុកជាសេចក្តីព្រាង។",
    },
  };

  const store = {
    get(k) { try { return localStorage.getItem(k); } catch { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch { /* ignore */ } },
    del(k) { try { localStorage.removeItem(k); } catch { /* ignore */ } },
  };
  let lang = store.get("toword.lang") || "en";
  const t = (k, vars) => {
    let s = (T[lang] && T[lang][k]) || T.en[k] || k;
    if (vars) for (const [a, b] of Object.entries(vars)) s = s.replace(`{${a}}`, b);
    return s;
  };
  function applyI18n() {
    document.documentElement.lang = lang === "km" ? "km" : "en";
    document.querySelectorAll("[data-i18n]").forEach((el) => { el.textContent = t(el.dataset.i18n); });
    $("#langToggle").textContent = lang === "km" ? "English" : "ខ្មែរ";
    if (job) updateEditorChrome();
  }

  // ------------------------------------------------------------------ utils
  const $ = (s) => document.querySelector(s);
  const views = { upload: $("#viewUpload"), progress: $("#viewProgress"), editor: $("#viewEditor") };
  function show(name) {
    for (const [k, el] of Object.entries(views)) el.hidden = k !== name;
    document.body.classList.toggle("editing", name === "editor");
  }
  let toastTimer;
  function toast(msg, ms = 3000) {
    const el = $("#toast");
    el.textContent = msg; el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { el.hidden = true; }, ms);
  }
  async function api(path, opts = {}) {
    const res = await fetch(path, opts);
    if (!res.ok) {
      let msg = res.statusText;
      try { msg = (await res.json()).detail || msg; } catch { /* ignore */ }
      throw new Error(msg);
    }
    return res;
  }
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // ------------------------------------------------------------------ state
  let config = { fonts: [], accept: [] };
  let job = null;          // status object from the server
  let editor = null;       // TinyMCE instance
  let dirty = false, saveTimer = null, loadingContent = false;

  // ------------------------------------------------------------------ init
  async function init() {
    applyI18n();
    $("#langToggle").addEventListener("click", () => {
      lang = lang === "km" ? "en" : "km"; store.set("toword.lang", lang); applyI18n();
    });
    try {
      config = await (await api("/api/config")).json();
    } catch { /* server down: upload will show error */ }
    const sel = $("#fontSelect");
    const savedFont = store.get("toword.font");
    for (const f of config.fonts || []) {
      const o = document.createElement("option");
      o.value = f; o.textContent = f; o.style.fontFamily = `'${f}'`;
      if (f === savedFont) o.selected = true;
      sel.appendChild(o);
    }
    sel.addEventListener("change", () => store.set("toword.font", sel.value));
    $("#fileInput").accept = (config.accept || []).join(",");

    setupDropzone();
    $("#progBack").addEventListener("click", goHome);
    $("#btnNew").addEventListener("click", async () => { await flushSave(); goHome(); });
    $("#brand").addEventListener("click", async (e) => { e.preventDefault(); await flushSave(); goHome(); });
    $("#btnCompare").addEventListener("click", toggleCompare);
    const rbtn = $("#btnReconvert"), rmenu = $("#reconvertMenu");
    rbtn.addEventListener("click", (e) => { e.stopPropagation(); $("#downloadMenu").hidden = true; rmenu.hidden = !rmenu.hidden; });
    document.addEventListener("click", () => { rmenu.hidden = true; });
    rmenu.addEventListener("click", (e) => {
      const b = e.target.closest("button[data-mode]");
      if (b) { rmenu.hidden = true; reconvert(b.dataset.mode); }
    });
    setupDownloadMenu();
    window.addEventListener("hashchange", route);
    window.addEventListener("beforeunload", (e) => { if (dirty) { flushSave(); e.preventDefault(); } });
    route();
  }

  function goHome() {
    location.hash = "";
    job = null;
    show("upload");
    showResume();
  }

  async function showResume() {
    const last = store.get("toword.lastJob");
    const box = $("#resumeBox");
    box.hidden = true;
    if (!last) return;
    try {
      const st = await (await api(`/api/jobs/${last}`)).json();
      if (st.status !== "done") return;
      $("#resumeName").textContent = st.filename;
      $("#resumeBtn").onclick = () => { location.hash = `job=${last}`; };
      box.hidden = false;
    } catch { store.del("toword.lastJob"); }
  }

  function route() {
    const m = location.hash.match(/job=([a-f0-9]+)/);
    if (m) openJob(m[1]);
    else { show("upload"); showResume(); }
  }

  // ------------------------------------------------------------------ upload
  function setupDropzone() {
    const dz = $("#dropzone"), input = $("#fileInput");
    input.addEventListener("change", () => { if (input.files[0]) upload(input.files[0]); input.value = ""; });
    dz.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
    ["dragenter", "dragover"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); }));
    ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); }));
    dz.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) upload(f); });
    // Allow dropping anywhere on the upload page.
    document.addEventListener("dragover", (e) => e.preventDefault());
    document.addEventListener("drop", (e) => {
      e.preventDefault();
      if (!views.upload.hidden && e.dataTransfer.files[0] && !dz.contains(e.target)) upload(e.dataTransfer.files[0]);
    });
  }

  function upload(file) {
    const ext = "." + file.name.split(".").pop().toLowerCase();
    if (config.accept?.length && !config.accept.includes(ext)) { toast(t("unsupported")); return; }
    const fd = new FormData();
    fd.append("file", file);
    fd.append("mode", document.querySelector("input[name=mode]:checked").value);
    fd.append("lang", $("#ocrLang").value);
    fd.append("font", $("#fontSelect").value);

    resetProgress(file.name);
    setProgress(0, t("uploading"));
    const xhr = new XMLHttpRequest();
    xhr.open("POST", "/api/upload");
    xhr.upload.onprogress = (e) => { if (e.lengthComputable) setProgress(Math.round(e.loaded / e.total * 100), t("uploading")); };
    xhr.onload = () => {
      let body = {};
      try { body = JSON.parse(xhr.responseText); } catch { /* ignore */ }
      if (xhr.status >= 400) return failProgress(body.detail || xhr.statusText);
      location.hash = `job=${body.id}`;
    };
    xhr.onerror = () => failProgress("Network error");
    xhr.send(fd);
  }

  function resetProgress(name) {
    show("progress");
    $(".progress-card").classList.remove("failed");
    $("#progName").textContent = name;
    $("#progError").hidden = true;
    $("#progBack").hidden = true;
  }
  function setProgress(pct, msg) {
    $("#progBar").style.width = `${pct}%`;
    $("#progMsg").textContent = msg;
  }
  function failProgress(msg) {
    show("progress");
    $(".progress-card").classList.add("failed");
    $("#progError").textContent = `${t("failed")}: ${msg}`;
    $("#progError").hidden = false;
    $("#progBack").hidden = false;
  }

  // ------------------------------------------------------------------ job
  async function openJob(id) {
    let st;
    try { st = await (await api(`/api/jobs/${id}`)).json(); }
    catch (e) { failProgress(e.message); return; }
    if (st.status !== "done") {
      resetProgress(st.filename);
      while (st.status !== "done") {
        if (st.status === "error") return failProgress(st.error);
        setProgress(st.progress || 0, st.message || "…");
        await sleep(700);
        if (!location.hash.includes(id)) return; // user navigated away
        try { st = await (await api(`/api/jobs/${id}`)).json(); }
        catch (e) { return failProgress(e.message); }
      }
      setProgress(100, st.message);
    }
    job = st;
    store.set("toword.lastJob", id);
    const html = await (await api(`/api/jobs/${id}/html`)).text();
    show("editor");
    closeCompare();
    updateEditorChrome();
    await ensureEditor();
    applyPageSetup(job.page);
    loadingContent = true;
    editor.setContent(html);
    editor.undoManager.clear();
    loadingContent = false;
    dirty = false;
    setSaveState(job.has_draft ? t("saved") : "");
  }

  function updateEditorChrome() {
    if (!job) return;
    $("#docName").textContent = job.filename;
    $("#docName").title = job.filename;
    const used = job.mode_used || "layout";
    $("#modeBadge").textContent = t(`modeBadge_${used}`);
    const isPdf = job.ext === ".pdf";
    $("#btnReconvert").parentElement.hidden = !isPdf;
    document.querySelectorAll("#reconvertMenu button").forEach((b) => { b.hidden = b.dataset.mode === used; });
    const n = $("#notice");
    n.hidden = false;
    if (used === "hybrid") {
      n.textContent = t("noticeHybrid");
    } else if (used === "ocr" && job.mode === "auto" && job.quality?.reasons?.length) {
      n.textContent = t("noticeAutoOcr", { why: job.quality.reasons.join(", ") });
    } else if (used === "ocr") {
      n.textContent = t("noticeOcr");
    } else if (isPdf && (job.quality?.khmer_chars || 0) > 0) {
      n.textContent = t("noticeLayout");
    } else {
      n.hidden = true;
    }
    $("#btnCompare").textContent = $("#paneOriginal").hidden ? t("compare") : t("hideCompare");
  }

  async function reconvert(mode) {
    if (!job) return;
    await flushSave();
    try {
      const res = await (await api(`/api/jobs/${job.id}/reconvert`, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ mode, lang: $("#ocrLang").value }),
      })).json();
      location.hash = `job=${res.id}`;
    } catch (e) { toast(e.message); }
  }

  // ------------------------------------------------------------------ compare
  function toggleCompare() {
    const pane = $("#paneOriginal");
    if (pane.hidden) {
      const frame = $("#originalFrame");
      const src = `/api/jobs/${job.id}/source`;
      if (frame.dataset.src !== src) {
        if (job.ext === ".pdf") frame.src = src;
        else if ([".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"].includes(job.ext)) {
          frame.srcdoc = `<body style="margin:0;background:#888;display:flex;justify-content:center"><img src="${src}" style="max-width:100%;height:auto"></body>`;
        } else frame.src = src;
        frame.dataset.src = src;
      }
      pane.hidden = false;
    } else pane.hidden = true;
    updateEditorChrome();
  }
  function closeCompare() {
    $("#paneOriginal").hidden = true;
    const f = $("#originalFrame");
    f.removeAttribute("src"); f.removeAttribute("srcdoc"); delete f.dataset.src;
  }

  // ------------------------------------------------------------------ editor
  const KHMER_FONT_FAMILIES = [
    ["Khmer OS Battambang", "'Khmer OS Battambang', Battambang"],
    ["Khmer OS Siemreap", "'Khmer OS Siemreap', Siemreap"],
    ["Khmer OS", "'Khmer OS'"],
    ["Khmer OS Muol Light", "'Khmer OS Muol Light', Moul"],
    ["Khmer Muol", "'Khmer Muol', Moul"],
    ["Noto Sans Khmer", "'Noto Sans Khmer'"],
    ["Noto Serif Khmer", "'Noto Serif Khmer'"],
    ["Battambang", "Battambang"],
    ["Kantumruy Pro", "'Kantumruy Pro'"],
    ["Hanuman", "Hanuman"],
    ["Moul", "Moul"],
    ["Koulen", "Koulen"],
  ];
  const LATIN = [
    ["Arial", "arial, helvetica, sans-serif"],
    ["Calibri", "calibri, sans-serif"],
    ["Cambria", "cambria, serif"],
    ["Courier New", "'courier new', courier, monospace"],
    ["Georgia", "georgia, serif"],
    ["Tahoma", "tahoma, sans-serif"],
    ["Times New Roman", "'times new roman', times, serif"],
    ["Verdana", "verdana, sans-serif"],
  ];

  function ensureEditor() {
    if (editor) return Promise.resolve(editor);
    const fontFormats = [...KHMER_FONT_FAMILIES, ...LATIN].map(([n, v]) => `${n}=${v}`).join("; ");
    return tinymce.init({
      selector: "#editor",
      license_key: "gpl",
      base_url: "/vendor/tinymce",
      suffix: ".min",
      height: "100%",
      menubar: "file edit view insert format table tools",
      plugins: "advlist autolink lists link image charmap searchreplace visualblocks code fullscreen table pagebreak wordcount quickbars",
      toolbar:
        "undo redo | blocks fontfamily fontsize | bold italic underline strikethrough | forecolor backcolor | " +
        "alignleft aligncenter alignright alignjustify | lineheight | bullist numlist outdent indent | " +
        "table image link pagebreak | removeformat searchreplace code fullscreen",
      toolbar_mode: "sliding",
      quickbars_insert_toolbar: false,
      quickbars_selection_toolbar: "bold italic underline | fontsize forecolor | quicklink",
      font_family_formats: fontFormats,
      font_size_formats: "8pt 9pt 10pt 11pt 12pt 13pt 14pt 16pt 18pt 20pt 22pt 24pt 28pt 32pt 36pt 48pt",
      line_height_formats: "1 1.15 1.5 1.7 2 2.5 3",
      promotion: false,
      branding: false,
      convert_urls: false,
      paste_data_images: true,
      automatic_uploads: false,
      browser_spellcheck: true,
      contextmenu: "link image table",
      table_default_styles: { "border-collapse": "collapse", width: "100%" },
      table_default_attributes: { border: "1" },
      content_css: [
        "https://fonts.googleapis.com/css2?family=Battambang:wght@400;700&family=Hanuman:wght@400;700&family=Kantumruy+Pro:wght@400;700&family=Koulen&family=Moul&family=Noto+Sans+Khmer:wght@400;700&family=Noto+Serif+Khmer:wght@400;700&family=Siemreap&display=swap",
      ],
      content_style: `
        html { background: #e7e9ee; }
        body {
          box-sizing: border-box; max-width: 21cm; min-height: 29.7cm; margin: 24px auto; padding: 2cm 2cm;
          background: #fff; color: #111; box-shadow: 0 1px 3px rgba(0,0,0,.15), 0 6px 20px rgba(0,0,0,.08);
          font-family: 'Khmer OS Battambang', Battambang, 'Noto Sans Khmer', 'Times New Roman', serif;
          font-size: 12pt; line-height: 1.5; word-wrap: break-word;
        }
        @media (max-width: 22cm) { body { margin: 0; padding: 16px !important; min-height: 100%; box-shadow: none; } }
        p { margin: 0 0 4pt; }
        h1, h2, h3 { line-height: 1.5; }
        table { border-collapse: collapse; }
        td, th { padding: 2px 6px; vertical-align: top; }
        table:not([border="0"]) td:not([style*="border"]) { border: 1px dashed #c7ccd6; }
        td > p:last-child { margin-bottom: 0; }
        img { max-width: 100%; height: auto; }
        .tab { white-space: pre; }
        .mce-pagebreak { border-top: 2px dashed #9aa3b5; width: 100%; height: 6px; margin: 16px 0; }
      `,
      setup(ed) {
        ed.on("input change undo redo SetContent", (e) => {
          if (loadingContent || e.type === "setcontent") return;
          markDirty();
        });
        ed.addShortcut("meta+s", "Save draft", () => flushSave());
      },
    }).then(([ed]) => { editor = ed; return ed; });
  }

  // Mirror the Word page (size + margins) so lines wrap like in Word.
  function applyPageSetup(pg) {
    const b = editor.getBody();
    if (!pg) { b.style.removeProperty("max-width"); b.style.removeProperty("padding"); return; }
    b.style.maxWidth = `${pg.width}pt`;
    b.style.minHeight = `${pg.height}pt`;
    b.style.padding = `${pg.top}pt ${pg.right}pt ${pg.bottom}pt ${pg.left}pt`;
  }

  // ------------------------------------------------------------------ saving
  function setSaveState(s) { $("#saveState").textContent = s; }
  function markDirty() {
    dirty = true;
    setSaveState(t("saving"));
    clearTimeout(saveTimer);
    saveTimer = setTimeout(flushSave, 1500);
  }
  async function flushSave() {
    clearTimeout(saveTimer);
    if (!dirty || !job || !editor) return;
    dirty = false;
    try {
      await api(`/api/jobs/${job.id}/draft`, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ html: editor.getContent() }), keepalive: true,
      });
      setSaveState(t("saved"));
    } catch { dirty = true; setSaveState(t("saveFailed")); }
  }

  // ------------------------------------------------------------------ download
  function setupDownloadMenu() {
    const btn = $("#btnDownload"), menu = $("#downloadMenu");
    btn.addEventListener("click", (e) => { e.stopPropagation(); $("#reconvertMenu").hidden = true; menu.hidden = !menu.hidden; });
    document.addEventListener("click", () => { menu.hidden = true; });
    menu.addEventListener("click", async (e) => {
      const b = e.target.closest("button[data-dl]");
      if (!b) return;
      menu.hidden = true;
      await download(b.dataset.dl);
    });
  }

  function filenameFrom(res, fallback) {
    const cd = res.headers.get("Content-Disposition") || "";
    const star = cd.match(/filename\*=UTF-8''([^;]+)/i);
    if (star) return decodeURIComponent(star[1]);
    const plain = cd.match(/filename="([^"]+)"/i);
    return plain ? plain[1] : fallback;
  }

  async function download(kind) {
    if (!job) return;
    const btn = $("#btnDownload");
    btn.disabled = true;
    toast(t("preparing"), 60000);
    try {
      let res;
      if (kind === "original") {
        res = await api(`/api/jobs/${job.id}/converted.docx`);
      } else {
        await flushSave();
        res = await api(`/api/jobs/${job.id}/export`, {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ html: editor.getContent(), font: job.font || null, format: kind }),
        });
      }
      const blob = await res.blob();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = filenameFrom(res, `document.${kind === "pdf" ? "pdf" : "docx"}`);
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 10000);
      $("#toast").hidden = true;
    } catch (e) {
      toast(`${t("failed")}: ${e.message}`, 5000);
    } finally {
      btn.disabled = false;
    }
  }

  init();
})();
