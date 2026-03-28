/**
 * 学术论文自动排版工具 - 前端交互逻辑
 * 处理文件上传、拖拽、API 交互和 UI 状态管理
 */

(function () {
  "use strict";

  // ====== DOM 元素引用 ======
  const $ = (sel) => document.querySelector(sel);
  const tabUpload = $("#tabUpload");
  const tabMerge = $("#tabMerge");
  const tabPaste = $("#tabPaste");
  const uploadSection = $("#uploadSection");
  const pasteSection = $("#pasteSection");
  const mergeSection = $("#mergeSection");
  const pasteArea = $("#pasteArea");
  const btnFormatText = $("#btnFormatText");
  const uploadZone = $("#uploadZone");
  const fileInput = $("#fileInput");
  const filePreview = $("#filePreview");
  const fileName = $("#fileName");
  const fileSize = $("#fileSize");
  const fileRemove = $("#fileRemove");
  const btnFormat = $("#btnFormat");
  const processingSection = $("#processingSection");
  const resultSection = $("#resultSection");
  const errorSection = $("#errorSection");
  const errorMessage = $("#errorMessage");
  const btnDownload = $("#btnDownload");
  const btnReset = $("#btnReset");
  const btnRetry = $("#btnRetry");
  const retryLabel = $("#retryLabel");

  // 合并模式 DOM
  const mergeCoverZone = $("#mergeCoverZone");
  const mergeCoverInput = $("#mergeCoverInput");
  const mergeCoverFileEl = $("#mergeCoverFile");
  const mergeCoverName = $("#mergeCoverName");
  const mergeCoverSize = $("#mergeCoverSize");
  const mergeCoverRemove = $("#mergeCoverRemove");
  const mergeBodyZone = $("#mergeBodyZone");
  const mergeBodyInput = $("#mergeBodyInput");
  const mergeBodyFileEl = $("#mergeBodyFile");
  const mergeBodyName = $("#mergeBodyName");
  const mergeBodySize = $("#mergeBodySize");
  const mergeBodyRemove = $("#mergeBodyRemove");
  const btnMerge = $("#btnMerge");

  // 统计数据
  const statInputSize = $("#statInputSize");
  const statOutputSize = $("#statOutputSize");
  const statElapsed = $("#statElapsed");

  // 处理步骤
  const steps = [$("#step1"), $("#step2"), $("#step3"), $("#step4")];

  // ====== 状态 ======
  let selectedFile = null;
  let downloadUrl = "";
  let downloadName = "";
  let stepAnimationTimer = null;
  let resultRevealTimer = null;
  let currentRequestController = null;
  let isSubmitting = false;

  // 合并模式状态
  let selectedCoverFile = null;
  let selectedBodyFile = null;

  // ====== 初始化背景粒子 ======
  function initParticles() {
    const container = $("#bgParticles");
    const count = 30;
    for (let i = 0; i < count; i++) {
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

  // ====== 签名打字动画 ======
  function initSignatureTyping() {
    const signs = document.querySelectorAll(".sign-name");
    if (!signs.length) return;
    signs.forEach((el) => {
      const text = el.textContent.trim();
      if (!text) return;
      el.dataset.text = text;
      el.textContent = "";
      const underline = document.createElement("span");
      underline.className = "sign-underline";
      underline.setAttribute("aria-hidden", "true");
      el.appendChild(underline);
    });
    function typeText(el) {
      if (el.dataset.typing === "true") return;
      el.dataset.typing = "true";
      const chars = Array.from(el.dataset.text || "");
      const underline = el.querySelector(".sign-underline");
      let index = 0;
      function typeChar() {
        if (!underline) return;
        if (index < chars.length) {
          el.insertBefore(document.createTextNode(chars[index]), underline);
          index += 1;
          window.setTimeout(typeChar, 80 + Math.random() * 40);
          return;
        }
        el.classList.add("typed");
      }
      typeChar();
    }
    signs.forEach((el) => typeText(el));
  }

  // ====== UI 状态切换 ======
  function showSection(section) {
    [uploadSection, pasteSection, mergeSection, filePreview, processingSection, resultSection, errorSection].forEach((s) => {
      if (s) s.classList.add("hidden");
    });
    if (section) section.classList.remove("hidden");
  }

  function deactivateAllTabs() {
    tabUpload.classList.remove("active");
    tabMerge.classList.remove("active");
    tabPaste.classList.remove("active");
  }

  function showUpload() {
    deactivateAllTabs();
    tabUpload.classList.add("active");
    updateRetryLabel();
    if (selectedFile) { showPreview(); return; }
    showSection(uploadSection);
  }

  function showPaste() {
    deactivateAllTabs();
    tabPaste.classList.add("active");
    updateRetryLabel();
    showSection(pasteSection);
  }

  function showMerge() {
    deactivateAllTabs();
    tabMerge.classList.add("active");
    updateRetryLabel();
    showSection(mergeSection);
  }

  function showPreview() {
    updateRetryLabel();
    showSection(filePreview);
  }

  function showProcessing() {
    stopResultReveal();
    showSection(processingSection);
    resetSteps();
    steps[0].classList.add("active");
    animateSteps();
  }

  function showResult(data) {
    stopResultReveal();
    completeSteps();
    showSection(resultSection);
    statInputSize.textContent = data.stats.input_size;
    statOutputSize.textContent = data.stats.output_size;
    statElapsed.textContent = data.stats.elapsed;
    downloadUrl = data.download_url + "?name=" + encodeURIComponent(data.download_name);
    downloadName = data.download_name;
  }

  function showError(msg) {
    stopResultReveal();
    stopStepAnimation();
    showSection(errorSection);
    errorMessage.textContent = msg;
  }

  async function parseApiResponse(response, fallbackMessage) {
    const contentType = response.headers.get("content-type") || "";
    try {
      if (contentType.includes("application/json")) return await response.json();
      const text = (await response.text()).trim();
      if (!text) return { success: response.ok, error: fallbackMessage };
      try { return JSON.parse(text); } catch { return { success: response.ok, error: text }; }
    } catch (err) {
      console.error("Response parse error:", err);
      return { success: false, error: fallbackMessage };
    }
  }

  // ====== 步骤动画 ======
  function stopResultReveal() { if (resultRevealTimer !== null) { window.clearTimeout(resultRevealTimer); resultRevealTimer = null; } }
  function stopStepAnimation() { if (stepAnimationTimer !== null) { window.clearInterval(stepAnimationTimer); stepAnimationTimer = null; } }

  function syncBusyState() {
    [tabUpload, tabMerge, tabPaste, pasteArea, btnFormatText, fileRemove, btnFormat].forEach((el) => {
      if (el) el.disabled = isSubmitting;
    });
    if (btnMerge) btnMerge.disabled = isSubmitting || !selectedCoverFile || !selectedBodyFile;
    if (uploadZone) {
      if (isSubmitting) uploadZone.classList.remove("drag-over");
      uploadZone.classList.toggle("is-busy", isSubmitting);
      uploadZone.setAttribute("aria-disabled", String(isSubmitting));
    }
    [mergeCoverZone, mergeBodyZone].forEach((z) => {
      if (z) {
        z.classList.toggle("is-busy", isSubmitting);
        if (isSubmitting) z.classList.remove("drag-over");
      }
    });
  }

  function updateRetryLabel() {
    if (!retryLabel) return;
    if (tabPaste.classList.contains("active")) { retryLabel.textContent = "返回继续编辑"; return; }
    if (tabMerge.classList.contains("active")) { retryLabel.textContent = "返回继续选择"; return; }
    retryLabel.textContent = selectedFile ? "重新选择文件" : "重新上传";
  }

  function beginRequest() {
    if (isSubmitting) return null;
    stopResultReveal();
    isSubmitting = true;
    syncBusyState();
    currentRequestController = new AbortController();
    return currentRequestController;
  }

  function finishRequest(controller) {
    if (currentRequestController === controller) currentRequestController = null;
    isSubmitting = false;
    syncBusyState();
  }

  function abortActiveRequest() {
    if (currentRequestController) { currentRequestController.abort(); currentRequestController = null; }
    isSubmitting = false;
    syncBusyState();
  }

  function resetSteps() { stopStepAnimation(); steps.forEach((s) => { s.classList.remove("active", "done"); }); }
  function completeSteps() { stopStepAnimation(); steps.forEach((s) => { s.classList.remove("active"); s.classList.add("done"); }); }

  function animateSteps() {
    stopStepAnimation();
    let current = 1;
    stepAnimationTimer = window.setInterval(() => {
      if (current > 0 && current <= steps.length) { steps[current - 1].classList.remove("active"); steps[current - 1].classList.add("done"); }
      if (current < steps.length) { steps[current].classList.add("active"); current++; } else { stopStepAnimation(); }
    }, 600);
  }

  function scheduleResult(data) {
    stopResultReveal();
    resultRevealTimer = window.setTimeout(() => { resultRevealTimer = null; showResult(data); }, 500);
  }

  // ====== 文件工具 ======
  function formatSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(2) + " MB";
  }

  function validateDocx(file) {
    if (!file.name.toLowerCase().endsWith(".docx")) { showError("仅支持 .docx 格式的 Word 文档，请重新选择文件。"); return false; }
    if (file.size > 50 * 1024 * 1024) { showError("文件大小超过 50MB 限制，请压缩后重试。"); return false; }
    return true;
  }

  // ====== 选中文件（单文件模式） ======
  function selectFile(file) {
    if (!file || !validateDocx(file)) return;
    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatSize(file.size);
    updateRetryLabel();
    showPreview();
  }

  // ====== 合并模式文件选择 ======
  function selectMergeCover(file) {
    if (!file || !validateDocx(file)) return;
    selectedCoverFile = file;
    mergeCoverName.textContent = file.name;
    mergeCoverSize.textContent = formatSize(file.size);
    mergeCoverZone.classList.add("hidden");
    mergeCoverFileEl.classList.remove("hidden");
    updateMergeButton();
  }

  function selectMergeBody(file) {
    if (!file || !validateDocx(file)) return;
    selectedBodyFile = file;
    mergeBodyName.textContent = file.name;
    mergeBodySize.textContent = formatSize(file.size);
    mergeBodyZone.classList.add("hidden");
    mergeBodyFileEl.classList.remove("hidden");
    updateMergeButton();
  }

  function removeMergeCover() {
    selectedCoverFile = null;
    mergeCoverInput.value = "";
    mergeCoverZone.classList.remove("hidden");
    mergeCoverFileEl.classList.add("hidden");
    updateMergeButton();
  }

  function removeMergeBody() {
    selectedBodyFile = null;
    mergeBodyInput.value = "";
    mergeBodyZone.classList.remove("hidden");
    mergeBodyFileEl.classList.add("hidden");
    updateMergeButton();
  }

  function updateMergeButton() {
    if (btnMerge) btnMerge.disabled = !selectedCoverFile || !selectedBodyFile || isSubmitting;
  }

  // ====== 重置 ======
  function resetAll() {
    abortActiveRequest();
    stopResultReveal();
    stopStepAnimation();
    selectedFile = null;
    downloadUrl = "";
    downloadName = "";
    fileInput.value = "";
    // 重置合并状态
    selectedCoverFile = null;
    selectedBodyFile = null;
    if (mergeCoverInput) mergeCoverInput.value = "";
    if (mergeBodyInput) mergeBodyInput.value = "";
    if (mergeCoverZone) mergeCoverZone.classList.remove("hidden");
    if (mergeCoverFileEl) mergeCoverFileEl.classList.add("hidden");
    if (mergeBodyZone) mergeBodyZone.classList.remove("hidden");
    if (mergeBodyFileEl) mergeBodyFileEl.classList.add("hidden");
    updateMergeButton();
    updateRetryLabel();
    if (tabMerge.classList.contains("active")) { showMerge(); }
    else if (tabPaste.classList.contains("active")) { showPaste(); }
    else { showUpload(); }
  }

  // ====== 上传并排版（单文件） ======
  async function uploadAndFormat() {
    if (!selectedFile || isSubmitting) return;
    const controller = beginRequest();
    if (!controller) return;
    showProcessing();
    const formData = new FormData();
    formData.append("file", selectedFile);
    try {
      const response = await fetch("/api/format", { method: "POST", body: formData, signal: controller.signal });
      const data = await parseApiResponse(response, "排版处理失败，请稍后重试。");
      if (!response.ok || !data.success) { showError(data.error || "排版处理失败，请稍后重试。"); return; }
      completeSteps();
      scheduleResult(data);
    } catch (err) {
      if (err && err.name === "AbortError") return;
      console.error("Upload error:", err);
      showError("网络连接失败，请检查网络后重试。");
    } finally { finishRequest(controller); }
  }

  // ====== 提交黏贴的文字 ======
  async function submitPastedText() {
    const text = pasteArea.value.trim();
    if (!text) { showError("文本内容不能为空，请黏贴点文字吧。"); return; }
    if (isSubmitting) return;
    const controller = beginRequest();
    if (!controller) return;
    showProcessing();
    try {
      const response = await fetch("/api/format_text", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text: text }), signal: controller.signal });
      const data = await parseApiResponse(response, "排版处理失败，请稍后重试。");
      if (!response.ok || !data.success) { showError(data.error || "排版处理失败，请稍后重试。"); return; }
      completeSteps();
      scheduleResult(data);
    } catch (err) {
      if (err && err.name === "AbortError") return;
      console.error("Text processing error:", err);
      showError("网络连接失败，请检查网络后重试。");
    } finally { finishRequest(controller); }
  }

  // ====== 合并排版 ======
  async function uploadAndMerge() {
    if (!selectedCoverFile || !selectedBodyFile || isSubmitting) return;
    const controller = beginRequest();
    if (!controller) return;
    showProcessing();
    const formData = new FormData();
    formData.append("cover", selectedCoverFile);
    formData.append("body", selectedBodyFile);
    try {
      const response = await fetch("/api/format_merge", { method: "POST", body: formData, signal: controller.signal });
      const data = await parseApiResponse(response, "合并排版失败，请稍后重试。");
      if (!response.ok || !data.success) { showError(data.error || "合并排版失败，请稍后重试。"); return; }
      completeSteps();
      scheduleResult(data);
    } catch (err) {
      if (err && err.name === "AbortError") return;
      console.error("Merge error:", err);
      showError("网络连接失败，请检查网络后重试。");
    } finally { finishRequest(controller); }
  }

  // ====== 事件绑定 ======
  tabUpload.addEventListener("click", showUpload);
  tabMerge.addEventListener("click", showMerge);
  tabPaste.addEventListener("click", showPaste);
  btnFormatText.addEventListener("click", submitPastedText);
  uploadZone.addEventListener("click", () => { if (!isSubmitting) fileInput.click(); });
  fileInput.addEventListener("change", (e) => { if (e.target.files.length > 0) selectFile(e.target.files[0]); });

  // 拖拽 - 主上传区
  uploadZone.addEventListener("dragover", (e) => { e.preventDefault(); if (!isSubmitting) uploadZone.classList.add("drag-over"); });
  uploadZone.addEventListener("dragleave", (e) => { e.preventDefault(); uploadZone.classList.remove("drag-over"); });
  uploadZone.addEventListener("drop", (e) => { e.preventDefault(); uploadZone.classList.remove("drag-over"); if (!isSubmitting && e.dataTransfer.files.length > 0) selectFile(e.dataTransfer.files[0]); });

  fileRemove.addEventListener("click", resetAll);
  btnFormat.addEventListener("click", uploadAndFormat);

  // 合并模式事件
  mergeCoverZone.addEventListener("click", () => { if (!isSubmitting) mergeCoverInput.click(); });
  mergeBodyZone.addEventListener("click", () => { if (!isSubmitting) mergeBodyInput.click(); });
  mergeCoverInput.addEventListener("change", (e) => { if (e.target.files.length > 0) selectMergeCover(e.target.files[0]); });
  mergeBodyInput.addEventListener("change", (e) => { if (e.target.files.length > 0) selectMergeBody(e.target.files[0]); });
  mergeCoverRemove.addEventListener("click", removeMergeCover);
  mergeBodyRemove.addEventListener("click", removeMergeBody);
  btnMerge.addEventListener("click", uploadAndMerge);

  // 合并区域拖拽
  [mergeCoverZone, mergeBodyZone].forEach((zone, idx) => {
    zone.addEventListener("dragover", (e) => { e.preventDefault(); if (!isSubmitting) zone.classList.add("drag-over"); });
    zone.addEventListener("dragleave", (e) => { e.preventDefault(); zone.classList.remove("drag-over"); });
    zone.addEventListener("drop", (e) => {
      e.preventDefault(); zone.classList.remove("drag-over");
      if (!isSubmitting && e.dataTransfer.files.length > 0) { idx === 0 ? selectMergeCover(e.dataTransfer.files[0]) : selectMergeBody(e.dataTransfer.files[0]); }
    });
  });

  // 下载 / 重置
  btnDownload.addEventListener("click", () => {
    if (downloadUrl) { const a = document.createElement("a"); a.href = downloadUrl; a.download = downloadName; document.body.appendChild(a); a.click(); a.remove(); }
  });
  btnReset.addEventListener("click", resetAll);
  btnRetry.addEventListener("click", resetAll);

  // 全局拖拽防止浏览器打开文件
  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("drop", (e) => e.preventDefault());

  // 全局粘贴
  document.addEventListener("paste", (e) => {
    if (!isSubmitting && !uploadSection.classList.contains("hidden") && e.clipboardData && e.clipboardData.files.length > 0) {
      e.preventDefault(); selectFile(e.clipboardData.files[0]);
    }
  });

  // ====== 初始化 ======
  syncBusyState();
  updateRetryLabel();
  updateMergeButton();
  initParticles();
  initSignatureTyping();
})();
