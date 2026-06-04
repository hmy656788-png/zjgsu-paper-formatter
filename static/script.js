/**
 * 学术论文自动排版工具 - 前端交互逻辑
 */
(function () {
  "use strict";

  const $ = (sel) => document.querySelector(sel);
  const tabUpload = $("#tabUpload");
  const tabPaste = $("#tabPaste");
  const tabConcat = $("#tabConcat");
  const uploadSection = $("#uploadSection");
  const pasteSection = $("#pasteSection");
  const pasteArea = $("#pasteArea");
  const btnFormatText = $("#btnFormatText");
  const uploadZone = $("#uploadZone");
  const fileInput = $("#fileInput");
  const filePreview = $("#filePreview");
  const fileName = $("#fileName");
  const fileSize = $("#fileSize");
  const fileRemove = $("#fileRemove");
  const btnFormat = $("#btnFormat");
  const btnFormatLabel = $("#btnFormatLabel");
  const processingSection = $("#processingSection");
  const processingLive = $("#processingLive");
  const resultSection = $("#resultSection");
  const errorSection = $("#errorSection");
  const errorMessage = $("#errorMessage");
  const btnDownload = $("#btnDownload");
  const btnReset = $("#btnReset");
  const btnRetry = $("#btnRetry");
  const retryLabel = $("#retryLabel");

  // 封面（可选）
  const coverZone = $("#coverZone");
  const coverInput = $("#coverInput");
  const coverFileCard = $("#coverFileCard");
  const coverFileName = $("#coverFileName");
  const coverFileSize = $("#coverFileSize");
  const coverFileRemove = $("#coverFileRemove");
  const coverMetaEnabled = $("#coverMetaEnabled");
  const coverMetaPanel = $("#coverMetaPanel");
  const coverTitleInput = $("#coverTitleInput");
  const collegeInput = $("#collegeInput");
  const teacherInput = $("#teacherInput");
  const classNameInput = $("#classNameInput");
  const studentNameInput = $("#studentNameInput");
  const studentIdInput = $("#studentIdInput");

  // 拼接文档
  const concatSection = $("#concatSection");
  const concatZoneFirst = $("#concatZoneFirst");
  const concatZoneSecond = $("#concatZoneSecond");
  const concatFirstInput = $("#concatFirstInput");
  const concatSecondInput = $("#concatSecondInput");
  const concatFirstCard = $("#concatFirstCard");
  const concatSecondCard = $("#concatSecondCard");
  const concatFirstName = $("#concatFirstName");
  const concatSecondName = $("#concatSecondName");
  const concatFirstSize = $("#concatFirstSize");
  const concatSecondSize = $("#concatSecondSize");
  const concatFirstRemove = $("#concatFirstRemove");
  const concatSecondRemove = $("#concatSecondRemove");
  const concatRestartPage = $("#concatRestartPage");
  const btnConcat = $("#btnConcat");

  const statInputSize = $("#statInputSize");
  const statOutputSize = $("#statOutputSize");
  const statElapsed = $("#statElapsed");
  const previewHighlights = $("#previewHighlights");
  const previewOutline = $("#previewOutline");
  const outlinePanel = $("#outlinePanel");
  const resultPreview = $("#resultPreview");
  const highlightsSubtitle = $("#highlightsSubtitle");
  const resultTitle = $("#resultTitle");
  const btnDownloadLabel = $("#btnDownloadLabel");
  const btnResetLabel = $("#btnResetLabel");
  const processingTitle = $("#processingTitle");
  const processingHint = $("#processingHint");
  const steps = [$("#step1"), $("#step2"), $("#step3"), $("#step4")];
  const stepLabelEls = steps.map((s) => (s ? s.querySelector("span") : null));
  const FORMAT_STEP_LABELS = ["解析文档结构", "识别标题层级", "应用排版规则", "生成输出文档"];
  const CONCAT_STEP_LABELS = ["读取两个文档", "保留各自排版", "拼接文档内容", "生成输出文档"];

  const outlineLevelLabels = {
    title: "论文标题",
    h1: "一级标题",
    h2: "二级标题",
    h3: "三级标题",
    section: "章节标题",
    references: "参考文献",
    english_abstract_heading: "英文摘要",
    abstract: "摘要",
  };

  let selectedFile = null;
  let selectedCover = null;
  let selectedConcatFirst = null;
  let selectedConcatSecond = null;
  let currentMode = "format"; // "format" | "text" | "concat"
  let downloadUrl = "";
  let downloadName = "";
  let stepAnimationTimer = null;
  let resultRevealTimer = null;
  let currentRequestController = null;
  let currentEventSource = null;
  let isSubmitting = false;

  // ====== 背景粒子 ======
  function initParticles() {
    const container = $("#bgParticles");
    for (let i = 0; i < 30; i++) {
      const p = document.createElement("div");
      p.classList.add("particle");
      p.style.left = Math.random() * 100 + "%";
      p.style.animationDuration = 8 + Math.random() * 12 + "s";
      p.style.animationDelay = Math.random() * 10 + "s";
      p.style.width = p.style.height = 1.5 + Math.random() * 2.5 + "px";
      p.style.opacity = 0.15 + Math.random() * 0.35;
      container.appendChild(p);
    }
  }

  // ====== 签名动画 ======
  function initSignatureTyping() {
    const signs = document.querySelectorAll(".sign-name");
    signs.forEach((el) => {
      const text = el.textContent.trim();
      if (!text) return;
      el.dataset.text = text;
      el.textContent = "";
      const u = document.createElement("span");
      u.className = "sign-underline";
      u.setAttribute("aria-hidden", "true");
      el.appendChild(u);
    });
    signs.forEach((el) => {
      if (el.dataset.typing === "true") return;
      el.dataset.typing = "true";
      const chars = Array.from(el.dataset.text || "");
      const u = el.querySelector(".sign-underline");
      let i = 0;
      (function t() {
        if (!u) return;
        if (i < chars.length) { el.insertBefore(document.createTextNode(chars[i++]), u); window.setTimeout(t, 80 + Math.random() * 40); }
        else el.classList.add("typed");
      })();
    });
  }

  // ====== UI 切换 ======
  function showSection(sec) {
    [uploadSection, pasteSection, concatSection, filePreview, processingSection, resultSection, errorSection].forEach((s) => { if (s) s.classList.add("hidden"); });
    if (sec) sec.classList.remove("hidden");
  }

  function setActiveTab(activeTab) {
    [tabUpload, tabPaste, tabConcat].forEach((t) => { if (t) t.classList.toggle("active", t === activeTab); });
  }

  function showUpload() {
    setActiveTab(tabUpload); currentMode = "format";
    updateRetryLabel();
    if (selectedFile) { showPreview(); return; }
    showSection(uploadSection);
  }

  function showPaste() {
    setActiveTab(tabPaste); currentMode = "text";
    updateRetryLabel();
    showSection(pasteSection);
  }

  function showConcat() {
    setActiveTab(tabConcat); currentMode = "concat";
    updateRetryLabel();
    showSection(concatSection);
  }

  function showPreview() { updateRetryLabel(); showSection(filePreview); }

  function createEmptyPreviewMessage(text) {
    const empty = document.createElement("p");
    empty.className = "preview-empty";
    empty.textContent = text;
    return empty;
  }

  function renderPreview(preview) {
    if (!previewHighlights || !previewOutline) return;

    previewHighlights.textContent = "";
    previewOutline.textContent = "";

    const highlights = Array.isArray(preview && preview.highlights) ? preview.highlights : [];
    const outline = Array.isArray(preview && preview.outline) ? preview.outline : [];

    if (!highlights.length) {
      previewHighlights.appendChild(createEmptyPreviewMessage("这次排版已经完成，但暂时没有可展示的预览摘要。"));
    } else {
      highlights.forEach((item) => {
        const card = document.createElement("article");
        card.className = "preview-item";

        const eyebrow = document.createElement("span");
        eyebrow.className = "preview-item-eyebrow";
        eyebrow.textContent = item.eyebrow || "排版动作";

        const title = document.createElement("h5");
        title.className = "preview-item-title";
        title.textContent = item.title || "已完成自动处理";

        const description = document.createElement("p");
        description.className = "preview-item-desc";
        description.textContent = item.description || "";

        card.appendChild(eyebrow);
        card.appendChild(title);
        card.appendChild(description);
        previewHighlights.appendChild(card);
      });
    }

    if (!outline.length) {
      previewOutline.appendChild(createEmptyPreviewMessage("这份文档没有识别到可展示的标题结构，正文仍已完成统一排版。"));
      return;
    }

    const list = document.createElement("div");
    list.className = "outline-list";

    outline.forEach((item) => {
      const row = document.createElement("div");
      row.className = "outline-item";
      if (item.level) {
        row.dataset.level = item.level;
      }

      const label = document.createElement("span");
      label.className = "outline-level";
      label.textContent = outlineLevelLabels[item.level] || "结构";

      const text = document.createElement("span");
      text.className = "outline-text";
      text.textContent = item.text || "";

      row.appendChild(label);
      row.appendChild(text);
      list.appendChild(row);
    });

    previewOutline.appendChild(list);
  }

  function setProcessingLive(message) {
    if (processingLive) processingLive.textContent = message || "正在准备任务...";
  }

  function setStepLabels(labels) {
    stepLabelEls.forEach((el, i) => { if (el && labels[i]) el.textContent = labels[i]; });
  }

  function showProcessing(useAnimatedSteps = false) {
    stopResultReveal(); showSection(processingSection);
    const isConcat = currentMode === "concat";
    setStepLabels(isConcat ? CONCAT_STEP_LABELS : FORMAT_STEP_LABELS);
    if (processingTitle) processingTitle.textContent = isConcat ? "正在拼接中..." : "正在排版中...";
    if (processingHint) processingHint.textContent = isConcat
      ? "保持两个文档原有排版，合并为一个文件"
      : "智能识别论文结构，应用浙工商学术排版规范";
    resetSteps();
    setStepProgress(1);
    setProcessingLive("正在准备任务...");
    if (useAnimatedSteps) animateSteps();
  }

  function showResult(data) {
    stopResultReveal(); completeSteps(); showSection(resultSection);
    const isConcat = currentMode === "concat";
    if (resultTitle) resultTitle.textContent = isConcat ? "拼接完成！" : "排版完成！";
    if (btnDownloadLabel) btnDownloadLabel.textContent = isConcat ? "下载拼接文档" : "下载排版文档";
    if (btnResetLabel) btnResetLabel.textContent = isConcat ? "继续拼接其他文档" : "继续排版其他文档";
    if (outlinePanel) outlinePanel.classList.toggle("hidden", isConcat);
    if (resultPreview) resultPreview.classList.toggle("is-single", isConcat);
    if (highlightsSubtitle) highlightsSubtitle.textContent = isConcat
      ? "保持两个文档各自的排版形态，仅把它们合并为一个文件"
      : "把真正做过的页面设置、页眉页码和结构识别结果直接展示出来";
    statInputSize.textContent = data.stats.input_size;
    statOutputSize.textContent = data.stats.output_size;
    statElapsed.textContent = data.stats.elapsed;
    downloadUrl = data.download_url + "?name=" + encodeURIComponent(data.download_name);
    downloadName = data.download_name;
    if (data.preview || data.summary) {
      renderPreview(data.preview || data.summary);
    }
  }

  function showError(msg) {
    stopResultReveal(); stopStepAnimation(); closeProgressStream(); showSection(errorSection); errorMessage.textContent = msg;
  }

  async function parseApiResponse(res, fb) {
    const ct = res.headers.get("content-type") || "";
    try {
      if (ct.includes("application/json")) return await res.json();
      const t = (await res.text()).trim();
      if (!t) return { success: res.ok, error: fb };
      try { return JSON.parse(t); } catch { return { success: res.ok, error: t }; }
    } catch { return { success: false, error: fb }; }
  }

  // ====== 动画 ======
  function stopResultReveal() { if (resultRevealTimer !== null) { clearTimeout(resultRevealTimer); resultRevealTimer = null; } }
  function stopStepAnimation() { if (stepAnimationTimer !== null) { clearInterval(stepAnimationTimer); stepAnimationTimer = null; } }
  function closeProgressStream() { if (currentEventSource) { currentEventSource.close(); currentEventSource = null; } }

  function syncBusyState() {
    [tabUpload, tabPaste, tabConcat, pasteArea, btnFormatText, fileRemove, btnFormat, coverFileRemove, concatFirstRemove, concatSecondRemove, concatRestartPage].forEach((el) => { if (el) el.disabled = isSubmitting; });
    [coverMetaEnabled, coverTitleInput, collegeInput, teacherInput, classNameInput, studentNameInput, studentIdInput].forEach((el) => { if (el) el.disabled = isSubmitting; });
    if (uploadZone) { if (isSubmitting) uploadZone.classList.remove("drag-over"); uploadZone.classList.toggle("is-busy", isSubmitting); }
    if (coverZone) { if (isSubmitting) coverZone.classList.remove("drag-over"); coverZone.classList.toggle("is-busy", isSubmitting); }
    [concatZoneFirst, concatZoneSecond].forEach((z) => { if (z) { if (isSubmitting) z.classList.remove("drag-over"); z.classList.toggle("is-busy", isSubmitting); } });
    updateConcatButton();
  }

  function updateRetryLabel() {
    if (!retryLabel) return;
    if (tabConcat && tabConcat.classList.contains("active")) { retryLabel.textContent = "重新选择文档"; return; }
    if (tabPaste.classList.contains("active")) { retryLabel.textContent = "返回继续编辑"; return; }
    retryLabel.textContent = selectedFile ? "重新选择文件" : "重新上传";
  }

  function updateFormatButton() {
    if (!btnFormatLabel) return;
    if (selectedCover) { btnFormatLabel.textContent = "合并排版"; return; }
    if (coverMetaEnabled && coverMetaEnabled.checked) { btnFormatLabel.textContent = "生成封面并排版"; return; }
    btnFormatLabel.textContent = "开始排版";
  }

  function toggleCoverMetaPanel() {
    if (!coverMetaPanel || !coverMetaEnabled) return;
    coverMetaPanel.classList.toggle("hidden", !coverMetaEnabled.checked);
    updateFormatButton();
  }

  function collectCoverMeta(formData) {
    if (!coverMetaEnabled || !coverMetaEnabled.checked || !formData) return;
    formData.append("generate_cover", "1");
    [
      ["cover_title", coverTitleInput],
      ["college", collegeInput],
      ["teacher", teacherInput],
      ["class_name", classNameInput],
      ["student_name", studentNameInput],
      ["student_id", studentIdInput],
    ].forEach(([key, input]) => {
      const value = input && input.value ? input.value.trim() : "";
      if (value) formData.append(key, value);
    });
  }

  function beginRequest() {
    if (isSubmitting) return null;
    stopResultReveal(); closeProgressStream(); isSubmitting = true; syncBusyState();
    currentRequestController = new AbortController();
    return currentRequestController;
  }
  function finishRequest(c) { if (currentRequestController === c) currentRequestController = null; isSubmitting = false; syncBusyState(); }
  function abortActiveRequest() {
    if (currentRequestController) { currentRequestController.abort(); currentRequestController = null; }
    closeProgressStream();
    isSubmitting = false;
    syncBusyState();
  }

  function resetSteps() { stopStepAnimation(); steps.forEach((s) => s.classList.remove("active", "done")); }
  function completeSteps() { stopStepAnimation(); steps.forEach((s) => { s.classList.remove("active"); s.classList.add("done"); }); }
  function setStepProgress(step) {
    stopStepAnimation();
    steps.forEach((s, index) => {
      s.classList.remove("active", "done");
      if (index + 1 < step) s.classList.add("done");
      else if (index + 1 === step) s.classList.add("active");
    });
  }
  function animateSteps() {
    stopStepAnimation(); let cur = 1;
    stepAnimationTimer = setInterval(() => {
      if (cur > 0 && cur <= steps.length) { steps[cur - 1].classList.remove("active"); steps[cur - 1].classList.add("done"); }
      if (cur < steps.length) { steps[cur].classList.add("active"); cur++; } else stopStepAnimation();
    }, 600);
  }
  function scheduleResult(data) { stopResultReveal(); resultRevealTimer = setTimeout(() => { resultRevealTimer = null; showResult(data); }, 500); }

  async function fetchAsyncJobResult(resultUrl) {
    const res = await fetch(resultUrl);
    const data = await parseApiResponse(res, "获取排版结果失败，请稍后重试。");
    if (!res.ok || !data.success) throw new Error(data.error || "获取排版结果失败，请稍后重试。");
    return data;
  }

  function handleProgressEvent(payload) {
    if (payload && typeof payload.step === "number") setStepProgress(payload.step);
    const liveMessage = payload && payload.detail ? `${payload.message} · ${payload.detail}` : payload && payload.message;
    setProcessingLive(liveMessage || "服务器正在处理文档...");
  }

  function openProgressStream(eventsUrl, resultUrl, controller) {
    closeProgressStream();
    let isTerminal = false;
    const source = new EventSource(eventsUrl);
    currentEventSource = source;

    source.addEventListener("progress", (event) => {
      try {
        handleProgressEvent(JSON.parse(event.data || "{}"));
      } catch {
        setProcessingLive("服务器正在处理文档...");
      }
    });

    source.addEventListener("complete", async () => {
      if (isTerminal) return;
      isTerminal = true;
      closeProgressStream();
      setProcessingLive("排版完成，正在整理结果...");
      try {
        const result = await fetchAsyncJobResult(resultUrl);
        completeSteps();
        scheduleResult(result);
      } catch (err) {
        showError((err && err.message) || "获取排版结果失败，请稍后重试。");
      } finally {
        finishRequest(controller);
      }
    });

    source.addEventListener("failed", (event) => {
      if (isTerminal) return;
      isTerminal = true;
      closeProgressStream();
      let message = "排版处理失败，请稍后重试。";
      try {
        const payload = JSON.parse(event.data || "{}");
        message = payload.message || message;
      } catch {}
      showError(message);
      finishRequest(controller);
    });

    source.onerror = () => {
      if (isTerminal || currentEventSource !== source) return;
      isTerminal = true;
      closeProgressStream();
      showError("进度连接中断，请检查网络后重试。");
      finishRequest(controller);
    };
  }

  async function startAsyncJob(endpoint, fetchOptions, controller, fallbackMessage) {
    const res = await fetch(endpoint, fetchOptions);
    if (res.status === 413) { showError(PAYLOAD_LIMIT_MESSAGE); return false; }
    const data = await parseApiResponse(res, fallbackMessage);
    if (!res.ok || !data.success) {
      showError(data.error || fallbackMessage);
      return false;
    }

    setProcessingLive("任务已创建，正在连接实时进度流...");
    openProgressStream(data.events_url, data.result_url, controller);
    return true;
  }

  // ====== 工具 ======
  // 部署在 Vercel 上时，单次请求体有约 4.5MB 的平台硬限制（FUNCTION_PAYLOAD_TOO_LARGE）。
  // 这里提前拦截，避免用户看到天书般的平台报错；并对 413 做兜底处理。
  const MAX_UPLOAD_BYTES = 4.4 * 1024 * 1024;
  const PAYLOAD_LIMIT_MESSAGE = "文档体积超过服务器单次上传上限（约 4.5MB）。请压缩正文中的图片后再传。";
  function formatSize(b) { if (b < 1024) return b + " B"; if (b < 1048576) return (b / 1024).toFixed(1) + " KB"; return (b / 1048576).toFixed(2) + " MB"; }
  function exceedsUploadLimit(totalBytes) {
    if (totalBytes > MAX_UPLOAD_BYTES) {
      showError(`两个文档合计约 ${formatSize(totalBytes)}，${PAYLOAD_LIMIT_MESSAGE}`);
      return true;
    }
    return false;
  }
  function validateDocx(f) {
    if (!f.name.toLowerCase().endsWith(".docx")) { showError("仅支持 .docx 格式的 Word 文档"); return false; }
    if (f.size > MAX_UPLOAD_BYTES) { showError(`该文档约 ${formatSize(f.size)}，${PAYLOAD_LIMIT_MESSAGE}`); return false; }
    return true;
  }

  // ====== 文件选择 ======
  function selectFile(file) {
    if (!file || !validateDocx(file)) return;
    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatSize(file.size);
    updateRetryLabel(); updateFormatButton(); showPreview();
  }

  // ====== 封面选择 ======
  function selectCover(file) {
    if (!file || !validateDocx(file)) return;
    selectedCover = file;
    coverFileName.textContent = file.name;
    coverFileSize.textContent = formatSize(file.size);
    coverZone.classList.add("hidden");
    coverFileCard.classList.remove("hidden");
    updateFormatButton();
  }

  function removeCover() {
    selectedCover = null;
    coverInput.value = "";
    coverZone.classList.remove("hidden");
    coverFileCard.classList.add("hidden");
    updateFormatButton();
  }

  // ====== 拼接文档：文件选择 ======
  function updateConcatButton() {
    if (!btnConcat) return;
    btnConcat.disabled = isSubmitting || !selectedConcatFirst || !selectedConcatSecond;
  }

  function selectConcatFile(slot, file) {
    if (!file || !validateDocx(file)) return;
    const isFirst = slot === "first";
    if (isFirst) selectedConcatFirst = file; else selectedConcatSecond = file;
    (isFirst ? concatFirstName : concatSecondName).textContent = file.name;
    (isFirst ? concatFirstSize : concatSecondSize).textContent = formatSize(file.size);
    (isFirst ? concatZoneFirst : concatZoneSecond).classList.add("hidden");
    (isFirst ? concatFirstCard : concatSecondCard).classList.remove("hidden");
    updateConcatButton();
  }

  function removeConcatFile(slot) {
    const isFirst = slot === "first";
    if (isFirst) { selectedConcatFirst = null; if (concatFirstInput) concatFirstInput.value = ""; }
    else { selectedConcatSecond = null; if (concatSecondInput) concatSecondInput.value = ""; }
    (isFirst ? concatZoneFirst : concatZoneSecond).classList.remove("hidden");
    (isFirst ? concatFirstCard : concatSecondCard).classList.add("hidden");
    updateConcatButton();
  }

  function resetConcatSlots() {
    selectedConcatFirst = null;
    selectedConcatSecond = null;
    if (concatFirstInput) concatFirstInput.value = "";
    if (concatSecondInput) concatSecondInput.value = "";
    [concatZoneFirst, concatZoneSecond].forEach((z) => { if (z) z.classList.remove("hidden"); });
    [concatFirstCard, concatSecondCard].forEach((c) => { if (c) c.classList.add("hidden"); });
    updateConcatButton();
  }

  // ====== 重置 ======
  function resetAll() {
    abortActiveRequest(); stopResultReveal(); stopStepAnimation();
    selectedFile = null; selectedCover = null;
    downloadUrl = ""; downloadName = "";
    fileInput.value = ""; coverInput.value = "";
    coverZone.classList.remove("hidden");
    coverFileCard.classList.add("hidden");
    resetConcatSlots();
    updateRetryLabel(); updateFormatButton();
    if (tabConcat && tabConcat.classList.contains("active")) showConcat();
    else if (tabPaste.classList.contains("active")) showPaste();
    else showUpload();
  }

  // ====== 上传排版（自动判断是否合并） ======
  async function uploadAndFormatLegacy() {
    if (!selectedFile || isSubmitting) return;
    if (selectedCover && exceedsUploadLimit(selectedFile.size + selectedCover.size)) return;
    const controller = beginRequest();
    if (!controller) return;
    showProcessing(true);

    const formData = new FormData();
    let endpoint;

    if (selectedCover) {
      // 有封面 → 调合并接口
      formData.append("cover", selectedCover);
      formData.append("body", selectedFile);
      endpoint = "/api/format_merge";
    } else {
      // 无封面 → 普通排版
      formData.append("file", selectedFile);
      collectCoverMeta(formData);
      endpoint = "/api/format";
    }

    try {
      const res = await fetch(endpoint, { method: "POST", body: formData, signal: controller.signal });
      if (res.status === 413) { showError(PAYLOAD_LIMIT_MESSAGE); return; }
      const data = await parseApiResponse(res, "排版处理失败，请稍后重试。");
      if (!res.ok || !data.success) { showError(data.error || "排版处理失败"); return; }
      completeSteps(); scheduleResult(data);
    } catch (err) {
      if (err && err.name === "AbortError") return;
      showError("网络连接失败，请检查网络后重试。");
    } finally { finishRequest(controller); }
  }

  async function uploadAndFormat() {
    if (!selectedFile || isSubmitting) return;
    if (selectedCover && exceedsUploadLimit(selectedFile.size + selectedCover.size)) return;
    if (typeof EventSource === "undefined") return uploadAndFormatLegacy();

    const controller = beginRequest();
    if (!controller) return;
    showProcessing();

    const formData = new FormData();
    let endpoint;

    if (selectedCover) {
      formData.append("cover", selectedCover);
      formData.append("body", selectedFile);
      endpoint = "/api/format_merge_async";
    } else {
      formData.append("file", selectedFile);
      collectCoverMeta(formData);
      endpoint = "/api/format_async";
    }

    let handedOff = false;
    try {
      handedOff = await startAsyncJob(
        endpoint,
        { method: "POST", body: formData, signal: controller.signal },
        controller,
        "排版任务创建失败，请稍后重试。"
      );
    } catch (err) {
      if (err && err.name === "AbortError") return;
      showError("网络连接失败，请检查网络后重试。");
    } finally {
      if (!handedOff) finishRequest(controller);
    }
  }

  // ====== 文字排版 ======
  async function submitPastedTextLegacy() {
    const text = pasteArea.value.trim();
    if (!text) { showError("文本内容不能为空"); return; }
    if (isSubmitting) return;
    const controller = beginRequest();
    if (!controller) return;
    showProcessing(true);
    try {
      const res = await fetch("/api/format_text", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text }), signal: controller.signal });
      const data = await parseApiResponse(res, "排版失败");
      if (!res.ok || !data.success) { showError(data.error || "排版失败"); return; }
      completeSteps(); scheduleResult(data);
    } catch (err) {
      if (err && err.name === "AbortError") return;
      showError("网络连接失败");
    } finally { finishRequest(controller); }
  }

  async function submitPastedText() {
    const text = pasteArea.value.trim();
    if (!text) { showError("文本内容不能为空"); return; }
    if (isSubmitting) return;
    if (typeof EventSource === "undefined") return submitPastedTextLegacy();

    const controller = beginRequest();
    if (!controller) return;
    showProcessing();

    let handedOff = false;
    try {
      handedOff = await startAsyncJob(
        "/api/format_text_async",
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text }),
          signal: controller.signal,
        },
        controller,
        "排版任务创建失败，请稍后重试。"
      );
    } catch (err) {
      if (err && err.name === "AbortError") return;
      showError("网络连接失败，请检查网络后重试。");
    } finally {
      if (!handedOff) finishRequest(controller);
    }
  }

  // ====== 拼接文档（保持各自排版，合并为一个文件） ======
  async function submitConcatLegacy() {
    if (!selectedConcatFirst || !selectedConcatSecond || isSubmitting) return;
    if (exceedsUploadLimit(selectedConcatFirst.size + selectedConcatSecond.size)) return;
    currentMode = "concat";
    const controller = beginRequest();
    if (!controller) return;
    showProcessing(true);

    const formData = new FormData();
    formData.append("first", selectedConcatFirst);
    formData.append("second", selectedConcatSecond);
    formData.append("restart_page_number", concatRestartPage && concatRestartPage.checked ? "1" : "0");

    try {
      const res = await fetch("/api/concat", { method: "POST", body: formData, signal: controller.signal });
      if (res.status === 413) { showError(PAYLOAD_LIMIT_MESSAGE); return; }
      const data = await parseApiResponse(res, "文档拼接失败，请稍后重试。");
      if (!res.ok || !data.success) { showError(data.error || "文档拼接失败"); return; }
      completeSteps(); scheduleResult(data);
    } catch (err) {
      if (err && err.name === "AbortError") return;
      showError("网络连接失败，请检查网络后重试。");
    } finally { finishRequest(controller); }
  }

  async function submitConcat() {
    if (!selectedConcatFirst || !selectedConcatSecond) { showError("请先分别选择两个 .docx 文档"); return; }
    if (isSubmitting) return;
    if (exceedsUploadLimit(selectedConcatFirst.size + selectedConcatSecond.size)) return;
    currentMode = "concat";
    if (typeof EventSource === "undefined") return submitConcatLegacy();

    const controller = beginRequest();
    if (!controller) return;
    showProcessing();

    const formData = new FormData();
    formData.append("first", selectedConcatFirst);
    formData.append("second", selectedConcatSecond);
    formData.append("restart_page_number", concatRestartPage && concatRestartPage.checked ? "1" : "0");

    let handedOff = false;
    try {
      handedOff = await startAsyncJob(
        "/api/concat_async",
        { method: "POST", body: formData, signal: controller.signal },
        controller,
        "拼接任务创建失败，请稍后重试。"
      );
    } catch (err) {
      if (err && err.name === "AbortError") return;
      showError("网络连接失败，请检查网络后重试。");
    } finally {
      if (!handedOff) finishRequest(controller);
    }
  }

  // ====== 事件绑定 ======
  tabUpload.addEventListener("click", showUpload);
  tabPaste.addEventListener("click", showPaste);
  if (tabConcat) tabConcat.addEventListener("click", showConcat);
  btnFormatText.addEventListener("click", submitPastedText);
  uploadZone.addEventListener("click", () => { if (!isSubmitting) fileInput.click(); });
  fileInput.addEventListener("change", (e) => { if (e.target.files.length > 0) selectFile(e.target.files[0]); });

  uploadZone.addEventListener("dragover", (e) => { e.preventDefault(); if (!isSubmitting) uploadZone.classList.add("drag-over"); });
  uploadZone.addEventListener("dragleave", (e) => { e.preventDefault(); uploadZone.classList.remove("drag-over"); });
  uploadZone.addEventListener("drop", (e) => { e.preventDefault(); uploadZone.classList.remove("drag-over"); if (!isSubmitting && e.dataTransfer.files.length > 0) selectFile(e.dataTransfer.files[0]); });

  fileRemove.addEventListener("click", resetAll);
  btnFormat.addEventListener("click", uploadAndFormat);

  // 封面事件
  coverZone.addEventListener("click", () => { if (!isSubmitting) coverInput.click(); });
  coverInput.addEventListener("change", (e) => { if (e.target.files.length > 0) selectCover(e.target.files[0]); });
  coverFileRemove.addEventListener("click", removeCover);
  if (coverMetaEnabled) coverMetaEnabled.addEventListener("change", toggleCoverMetaPanel);
  [coverTitleInput, collegeInput, teacherInput, classNameInput, studentNameInput, studentIdInput].forEach((input) => {
    if (input) input.addEventListener("input", updateFormatButton);
  });

  // 拼接文档事件
  if (btnConcat) btnConcat.addEventListener("click", submitConcat);
  [["first", concatZoneFirst, concatFirstInput, concatFirstRemove], ["second", concatZoneSecond, concatSecondInput, concatSecondRemove]].forEach(([slot, zone, input, removeBtn]) => {
    if (zone && input) {
      zone.addEventListener("click", () => { if (!isSubmitting) input.click(); });
      zone.addEventListener("dragover", (e) => { e.preventDefault(); if (!isSubmitting) zone.classList.add("drag-over"); });
      zone.addEventListener("dragleave", (e) => { e.preventDefault(); zone.classList.remove("drag-over"); });
      zone.addEventListener("drop", (e) => { e.preventDefault(); zone.classList.remove("drag-over"); if (!isSubmitting && e.dataTransfer.files.length > 0) selectConcatFile(slot, e.dataTransfer.files[0]); });
      input.addEventListener("change", (e) => { if (e.target.files.length > 0) selectConcatFile(slot, e.target.files[0]); });
    }
    if (removeBtn) removeBtn.addEventListener("click", () => removeConcatFile(slot));
  });

  btnDownload.addEventListener("click", () => {
    if (downloadUrl) { const a = document.createElement("a"); a.href = downloadUrl; a.download = downloadName; document.body.appendChild(a); a.click(); a.remove(); }
  });
  btnReset.addEventListener("click", resetAll);
  btnRetry.addEventListener("click", resetAll);

  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("drop", (e) => e.preventDefault());
  document.addEventListener("paste", (e) => {
    if (!isSubmitting && !uploadSection.classList.contains("hidden") && e.clipboardData && e.clipboardData.files.length > 0) {
      e.preventDefault(); selectFile(e.clipboardData.files[0]);
    }
  });

  syncBusyState(); updateRetryLabel(); updateFormatButton(); updateConcatButton(); toggleCoverMetaPanel(); initParticles(); initSignatureTyping();
})();
