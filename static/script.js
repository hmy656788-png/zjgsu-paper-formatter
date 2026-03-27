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
    showSection(uploadSection);
  }

  function showPaste() {
    tabUpload.classList.remove("active");
    tabPaste.classList.add("active");
    showSection(pasteSection);
  }

  function showPreview() {
    showSection(filePreview);
  }

  function showProcessing() {
    showSection(processingSection);
    // 重置步骤
    steps.forEach((s) => {
      s.classList.remove("active", "done");
    });
    steps[0].classList.add("active");
    animateSteps();
  }

  function showResult(data) {
    showSection(resultSection);
    statInputSize.textContent = data.stats.input_size;
    statOutputSize.textContent = data.stats.output_size;
    statElapsed.textContent = data.stats.elapsed;
    downloadUrl = data.download_url + "?name=" + encodeURIComponent(data.download_name);
    downloadName = data.download_name;
  }

  function showError(msg) {
    showSection(errorSection);
    errorMessage.textContent = msg;
  }

  // ====== 步骤动画 ======
  function animateSteps() {
    let current = 0;
    const interval = setInterval(() => {
      if (current > 0) {
        steps[current - 1].classList.remove("active");
        steps[current - 1].classList.add("done");
      }
      if (current < steps.length) {
        steps[current].classList.add("active");
        current++;
      } else {
        clearInterval(interval);
      }
    }, 600);
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
    showPreview();
  }

  // ====== 重置 ======
  function resetAll() {
    selectedFile = null;
    downloadUrl = "";
    downloadName = "";
    fileInput.value = "";
    
    // 恢复到当前激活的选项卡视图
    if (tabPaste.classList.contains("active")) {
      showPaste();
    } else {
      showUpload();
    }
  }

  // ====== 上传并排版 ======
  async function uploadAndFormat() {
    if (!selectedFile) return;

    showProcessing();

    const formData = new FormData();
    formData.append("file", selectedFile);

    try {
      const response = await fetch("/api/format", {
        method: "POST",
        body: formData,
      });

      const data = await response.json();

      if (!response.ok || !data.success) {
        showError(data.error || "排版处理失败，请稍后重试。");
        return;
      }

      // 完成所有步骤动画
      steps.forEach((s) => {
        s.classList.remove("active");
        s.classList.add("done");
      });

      // 延迟一下显示结果，让最后一步动画可见
      setTimeout(() => showResult(data), 500);
    } catch (err) {
      console.error("Upload error:", err);
      showError("网络连接失败，请检查网络后重试。");
    }
  }

  // ====== 提交黏贴的文字 ======
  async function submitPastedText() {
    const text = pasteArea.value.trim();
    if (!text) {
      showError("文本内容不能为空，请黏贴点文字吧。");
      return;
    }

    showProcessing();

    try {
      const response = await fetch("/api/format_text", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text: text }),
      });

      const data = await response.json();

      if (!response.ok || !data.success) {
        showError(data.error || "排版处理失败，请稍后重试。");
        return;
      }

      steps.forEach((s) => {
        s.classList.remove("active");
        s.classList.add("done");
      });

      setTimeout(() => showResult(data), 500);
    } catch (err) {
      console.error("Text processing error:", err);
      showError("网络连接失败，请检查网络后重试。");
    }
  }

  // ====== 事件绑定 ======

  // 选项卡切换
  tabUpload.addEventListener("click", showUpload);
  tabPaste.addEventListener("click", showPaste);

  // 文字模式下点击格式化按钮
  btnFormatText.addEventListener("click", submitPastedText);

  // 点击上传区域 → 打开文件选择器
  uploadZone.addEventListener("click", () => fileInput.click());

  // 文件选择
  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      selectFile(e.target.files[0]);
    }
  });

  // 拖拽事件
  uploadZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    uploadZone.classList.add("drag-over");
  });

  uploadZone.addEventListener("dragleave", (e) => {
    e.preventDefault();
    uploadZone.classList.remove("drag-over");
  });

  uploadZone.addEventListener("drop", (e) => {
    e.preventDefault();
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
    if (!uploadSection.classList.contains("hidden") && e.clipboardData && e.clipboardData.files.length > 0) {
      e.preventDefault();
      selectFile(e.clipboardData.files[0]);
    }
  });

  // ====== 初始化 ======
  initParticles();
  initSignatureTyping();
})();
