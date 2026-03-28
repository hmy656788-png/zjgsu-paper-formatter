/**
 * 学术论文自动排版工具 - 前端交互逻辑
 */
(function () {
  "use strict";

  const $ = (sel) => document.querySelector(sel);
  const tabUpload = $("#tabUpload");
  const tabPaste = $("#tabPaste");
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

  const statInputSize = $("#statInputSize");
  const statOutputSize = $("#statOutputSize");
  const statElapsed = $("#statElapsed");
  const steps = [$("#step1"), $("#step2"), $("#step3"), $("#step4")];

  let selectedFile = null;
  let selectedCover = null;
  let downloadUrl = "";
  let downloadName = "";
  let stepAnimationTimer = null;
  let resultRevealTimer = null;
  let currentRequestController = null;
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
    [uploadSection, pasteSection, filePreview, processingSection, resultSection, errorSection].forEach((s) => { if (s) s.classList.add("hidden"); });
    if (sec) sec.classList.remove("hidden");
  }

  function showUpload() {
    tabUpload.classList.add("active"); tabPaste.classList.remove("active");
    updateRetryLabel();
    if (selectedFile) { showPreview(); return; }
    showSection(uploadSection);
  }

  function showPaste() {
    tabUpload.classList.remove("active"); tabPaste.classList.add("active");
    updateRetryLabel();
    showSection(pasteSection);
  }

  function showPreview() { updateRetryLabel(); showSection(filePreview); }

  function showProcessing() {
    stopResultReveal(); showSection(processingSection);
    resetSteps(); steps[0].classList.add("active"); animateSteps();
  }

  function showResult(data) {
    stopResultReveal(); completeSteps(); showSection(resultSection);
    statInputSize.textContent = data.stats.input_size;
    statOutputSize.textContent = data.stats.output_size;
    statElapsed.textContent = data.stats.elapsed;
    downloadUrl = data.download_url + "?name=" + encodeURIComponent(data.download_name);
    downloadName = data.download_name;
  }

  function showError(msg) { stopResultReveal(); stopStepAnimation(); showSection(errorSection); errorMessage.textContent = msg; }

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

  function syncBusyState() {
    [tabUpload, tabPaste, pasteArea, btnFormatText, fileRemove, btnFormat].forEach((el) => { if (el) el.disabled = isSubmitting; });
    if (uploadZone) { if (isSubmitting) uploadZone.classList.remove("drag-over"); uploadZone.classList.toggle("is-busy", isSubmitting); }
  }

  function updateRetryLabel() {
    if (!retryLabel) return;
    if (tabPaste.classList.contains("active")) { retryLabel.textContent = "返回继续编辑"; return; }
    retryLabel.textContent = selectedFile ? "重新选择文件" : "重新上传";
  }

  function updateFormatButton() {
    if (!btnFormatLabel) return;
    btnFormatLabel.textContent = selectedCover ? "合并排版" : "开始排版";
  }

  function beginRequest() {
    if (isSubmitting) return null;
    stopResultReveal(); isSubmitting = true; syncBusyState();
    currentRequestController = new AbortController();
    return currentRequestController;
  }
  function finishRequest(c) { if (currentRequestController === c) currentRequestController = null; isSubmitting = false; syncBusyState(); }
  function abortActiveRequest() { if (currentRequestController) { currentRequestController.abort(); currentRequestController = null; } isSubmitting = false; syncBusyState(); }

  function resetSteps() { stopStepAnimation(); steps.forEach((s) => s.classList.remove("active", "done")); }
  function completeSteps() { stopStepAnimation(); steps.forEach((s) => { s.classList.remove("active"); s.classList.add("done"); }); }
  function animateSteps() {
    stopStepAnimation(); let cur = 1;
    stepAnimationTimer = setInterval(() => {
      if (cur > 0 && cur <= steps.length) { steps[cur - 1].classList.remove("active"); steps[cur - 1].classList.add("done"); }
      if (cur < steps.length) { steps[cur].classList.add("active"); cur++; } else stopStepAnimation();
    }, 600);
  }
  function scheduleResult(data) { stopResultReveal(); resultRevealTimer = setTimeout(() => { resultRevealTimer = null; showResult(data); }, 500); }

  // ====== 工具 ======
  function formatSize(b) { if (b < 1024) return b + " B"; if (b < 1048576) return (b / 1024).toFixed(1) + " KB"; return (b / 1048576).toFixed(2) + " MB"; }
  function validateDocx(f) {
    if (!f.name.toLowerCase().endsWith(".docx")) { showError("仅支持 .docx 格式的 Word 文档"); return false; }
    if (f.size > 52428800) { showError("文件大小超过 50MB 限制"); return false; }
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

  // ====== 重置 ======
  function resetAll() {
    abortActiveRequest(); stopResultReveal(); stopStepAnimation();
    selectedFile = null; selectedCover = null;
    downloadUrl = ""; downloadName = "";
    fileInput.value = ""; coverInput.value = "";
    coverZone.classList.remove("hidden");
    coverFileCard.classList.add("hidden");
    updateRetryLabel(); updateFormatButton();
    if (tabPaste.classList.contains("active")) showPaste(); else showUpload();
  }

  // ====== 上传排版（自动判断是否合并） ======
  async function uploadAndFormat() {
    if (!selectedFile || isSubmitting) return;
    const controller = beginRequest();
    if (!controller) return;
    showProcessing();

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
      endpoint = "/api/format";
    }

    try {
      const res = await fetch(endpoint, { method: "POST", body: formData, signal: controller.signal });
      const data = await parseApiResponse(res, "排版处理失败，请稍后重试。");
      if (!res.ok || !data.success) { showError(data.error || "排版处理失败"); return; }
      completeSteps(); scheduleResult(data);
    } catch (err) {
      if (err && err.name === "AbortError") return;
      showError("网络连接失败，请检查网络后重试。");
    } finally { finishRequest(controller); }
  }

  // ====== 文字排版 ======
  async function submitPastedText() {
    const text = pasteArea.value.trim();
    if (!text) { showError("文本内容不能为空"); return; }
    if (isSubmitting) return;
    const controller = beginRequest();
    if (!controller) return;
    showProcessing();
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

  // ====== 事件绑定 ======
  tabUpload.addEventListener("click", showUpload);
  tabPaste.addEventListener("click", showPaste);
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

  syncBusyState(); updateRetryLabel(); updateFormatButton(); initParticles(); initSignatureTyping();
})();
