/**
 * 学术论文自动排版工具 - 前端交互逻辑
 * 处理文件上传、拖拽、API 交互和 UI 状态管理
 */

(function () {
  "use strict";

  // ====== DOM 元素引用 ======
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
  const processingSection = $("#processingSection");
  const resultSection = $("#resultSection");
  const errorSection = $("#errorSection");
  const errorMessage = $("#errorMessage");
  const btnDownload = $("#btnDownload");
  const btnReset = $("#btnReset");
  const btnRetry = $("#btnRetry");
  const retryLabel = $("#retryLabel");

  // 统计数据
  const statInputSize = $("#statInputSize");
  const statOutputSize = $("#statOutputSize");
  const statElapsed = $("#statElapsed");
  const previewHighlights = $("#previewHighlights");
  const previewOutline = $("#previewOutline");

  // 处理步骤
  const steps = [$("#step1"), $("#step2"), $("#step3"), $("#step4")];
  const outlineLevelLabels = {
    title: "论文标题",
    h1: "一级标题",
    h2: "二级标题",
    h3: "三级标题",
    section: "章节标题",
    references: "参考文献",
  };

  // ====== 状态 ======
  let selectedFile = null;
  let downloadUrl = "";
  let downloadName = "";
  let stepAnimationTimer = null;
  let resultRevealTimer = null;
  let currentRequestController = null;
  let isSubmitting = false;

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
    [uploadSection, pasteSection, filePreview, processingSection, resultSection, errorSection].forEach((s) => {
      if (s) s.classList.add("hidden");
    });
    if (section) section.classList.remove("hidden");
  }

  function showUpload() {
    tabUpload.classList.add("active");
    tabPaste.classList.remove("active");
    updateRetryLabel();
    if (selectedFile) {
      showPreview();
      return;
    }

    showSection(uploadSection);
  }

  function showPaste() {
    tabUpload.classList.remove("active");
    tabPaste.classList.add("active");
    updateRetryLabel();
    showSection(pasteSection);
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
    renderPreview(data.preview);
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
      if (contentType.includes("application/json")) {
        return await response.json();
      }

      const text = (await response.text()).trim();
      if (!text) {
        return { success: response.ok, error: fallbackMessage };
      }

      try {
        return JSON.parse(text);
      } catch {
        return { success: response.ok, error: text };
      }
    } catch (err) {
      console.error("Response parse error:", err);
      return { success: false, error: fallbackMessage };
    }
  }

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

  // ====== 步骤动画 ======
  function stopResultReveal() {
    if (resultRevealTimer !== null) {
      window.clearTimeout(resultRevealTimer);
      resultRevealTimer = null;
    }
  }

  function stopStepAnimation() {
    if (stepAnimationTimer !== null) {
      window.clearInterval(stepAnimationTimer);
      stepAnimationTimer = null;
    }
  }

  function syncBusyState() {
    [tabUpload, tabPaste, pasteArea, btnFormatText, fileRemove, btnFormat].forEach((el) => {
      if (el) el.disabled = isSubmitting;
    });

    if (uploadZone) {
      if (isSubmitting) {
        uploadZone.classList.remove("drag-over");
      }
      uploadZone.classList.toggle("is-busy", isSubmitting);
      uploadZone.setAttribute("aria-disabled", String(isSubmitting));
    }
  }

  function updateRetryLabel() {
    if (!retryLabel) return;

    if (tabPaste.classList.contains("active")) {
      retryLabel.textContent = "返回继续编辑";
      return;
    }

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
    if (currentRequestController === controller) {
      currentRequestController = null;
    }

    isSubmitting = false;
    syncBusyState();
  }

  function abortActiveRequest() {
    if (currentRequestController) {
      currentRequestController.abort();
      currentRequestController = null;
    }

    isSubmitting = false;
    syncBusyState();
  }

  function resetSteps() {
    stopStepAnimation();
    steps.forEach((s) => {
      s.classList.remove("active", "done");
    });
  }

  function completeSteps() {
    stopStepAnimation();
    steps.forEach((s) => {
      s.classList.remove("active");
      s.classList.add("done");
    });
  }

  function animateSteps() {
    stopStepAnimation();

    let current = 1;
    stepAnimationTimer = window.setInterval(() => {
      if (current > 0 && current <= steps.length) {
        steps[current - 1].classList.remove("active");
        steps[current - 1].classList.add("done");
      }

      if (current < steps.length) {
        steps[current].classList.add("active");
        current++;
      } else {
        stopStepAnimation();
      }
    }, 600);
  }

  function scheduleResult(data) {
    stopResultReveal();
    resultRevealTimer = window.setTimeout(() => {
      resultRevealTimer = null;
      showResult(data);
    }, 500);
  }

  // ====== 文件大小格式化 ======
  function formatSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(2) + " MB";
  }

  // ====== 选中文件 ======
  function selectFile(file) {
    if (!file) return;

    if (!file.name.toLowerCase().endsWith(".docx")) {
      showError("仅支持 .docx 格式的 Word 文档，请重新选择文件。");
      return;
    }

    if (file.size > 50 * 1024 * 1024) {
      showError("文件大小超过 50MB 限制，请压缩后重试。");
      return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatSize(file.size);
    updateRetryLabel();
    showPreview();
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
    renderPreview(null);
    updateRetryLabel();
    
    // 恢复到当前激活的选项卡视图
    if (tabPaste.classList.contains("active")) {
      showPaste();
    } else {
      showUpload();
    }
  }

  // ====== 上传并排版 ======
  async function uploadAndFormat() {
    if (!selectedFile || isSubmitting) return;

    const controller = beginRequest();
    if (!controller) return;

    showProcessing();

    const formData = new FormData();
    formData.append("file", selectedFile);

    try {
      const response = await fetch("/api/format", {
        method: "POST",
        body: formData,
        signal: controller.signal,
      });

      const data = await parseApiResponse(response, "排版处理失败，请稍后重试。");

      if (!response.ok || !data.success) {
        showError(data.error || "排版处理失败，请稍后重试。");
        return;
      }

      // 完成所有步骤动画
      completeSteps();

      // 延迟一下显示结果，让最后一步动画可见
      scheduleResult(data);
    } catch (err) {
      if (err && err.name === "AbortError") {
        return;
      }
      console.error("Upload error:", err);
      showError("网络连接失败，请检查网络后重试。");
    } finally {
      finishRequest(controller);
    }
  }

  // ====== 提交黏贴的文字 ======
  async function submitPastedText() {
    const text = pasteArea.value.trim();
    if (!text) {
      showError("文本内容不能为空，请黏贴点文字吧。");
      return;
    }

    if (isSubmitting) return;

    const controller = beginRequest();
    if (!controller) return;

    showProcessing();

    try {
      const response = await fetch("/api/format_text", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text: text }),
        signal: controller.signal,
      });

      const data = await parseApiResponse(response, "排版处理失败，请稍后重试。");

      if (!response.ok || !data.success) {
        showError(data.error || "排版处理失败，请稍后重试。");
        return;
      }

      completeSteps();

      scheduleResult(data);
    } catch (err) {
      if (err && err.name === "AbortError") {
        return;
      }
      console.error("Text processing error:", err);
      showError("网络连接失败，请检查网络后重试。");
    } finally {
      finishRequest(controller);
    }
  }

  // ====== 事件绑定 ======

  // 选项卡切换
  tabUpload.addEventListener("click", showUpload);
  tabPaste.addEventListener("click", showPaste);

  // 文字模式下点击格式化按钮
  btnFormatText.addEventListener("click", submitPastedText);

  // 点击上传区域 → 打开文件选择器
  uploadZone.addEventListener("click", () => {
    if (isSubmitting) return;
    fileInput.click();
  });

  // 文件选择
  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      selectFile(e.target.files[0]);
    }
  });

  // 拖拽事件
  uploadZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    if (isSubmitting) return;
    uploadZone.classList.add("drag-over");
  });

  uploadZone.addEventListener("dragleave", (e) => {
    e.preventDefault();
    if (isSubmitting) return;
    uploadZone.classList.remove("drag-over");
  });

  uploadZone.addEventListener("drop", (e) => {
    e.preventDefault();
    if (isSubmitting) return;
    uploadZone.classList.remove("drag-over");
    if (e.dataTransfer.files.length > 0) {
      selectFile(e.dataTransfer.files[0]);
    }
  });

  // 移除文件
  fileRemove.addEventListener("click", resetAll);

  // 开始排版
  btnFormat.addEventListener("click", uploadAndFormat);

  // 下载
  btnDownload.addEventListener("click", () => {
    if (downloadUrl) {
      const a = document.createElement("a");
      a.href = downloadUrl;
      a.download = downloadName;
      document.body.appendChild(a);
      a.click();
      a.remove();
    }
  });

  // 重新开始
  btnReset.addEventListener("click", resetAll);
  btnRetry.addEventListener("click", resetAll);

  // 全局拖拽防止浏览器打开文件
  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("drop", (e) => e.preventDefault());

  // ====== 全局粘贴事件（Ctrl+V / Cmd+V） ======
  document.addEventListener("paste", (e) => {
    // 只有在上传界面可见时，才处理粘贴操作
    if (!isSubmitting && !uploadSection.classList.contains("hidden") && e.clipboardData && e.clipboardData.files.length > 0) {
      e.preventDefault();
      selectFile(e.clipboardData.files[0]);
    }
  });

  // ====== 初始化 ======
  syncBusyState();
  updateRetryLabel();
  initParticles();
  initSignatureTyping();
})();
