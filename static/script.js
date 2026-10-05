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
  const fileCard = $("#fileCard");
  const fileName = $("#fileName");
  const fileSize = $("#fileSize");
  const fileRemove = $("#fileRemove");
  const fileSelectionStatus = $("#fileSelectionStatus");
  const btnFormat = $("#btnFormat");
  const btnFormatLabel = $("#btnFormatLabel");
  const processingSection = $("#processingSection");
  const processingLive = $("#processingLive");
  const resultSection = $("#resultSection");
  const errorSection = $("#errorSection");
  const errorTitle = $("#errorTitle");
  const errorMessage = $("#errorMessage");
  const btnDownload = $("#btnDownload");
  const btnReset = $("#btnReset");
  const btnRetry = $("#btnRetry");
  const btnDiscardPending = $("#btnDiscardPending");
  const btnCancelRequest = $("#btnCancelRequest");
  const retryLabel = $("#retryLabel");

  // 封面（可选）
  const coverZone = $("#coverZone");
  const coverInput = $("#coverInput");
  const coverFileCard = $("#coverFileCard");
  const coverFileName = $("#coverFileName");
  const coverFileSize = $("#coverFileSize");
  const coverFileRemove = $("#coverFileRemove");
  const coverBuilder = $("#coverBuilder");
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
  const concatFeatureNotice = $("#concatFeatureNotice");

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
  const stepLabelEls = steps.map((s) => (s ? s.querySelector(".step-label") : null));
  const FORMAT_STEP_LABELS = ["解析文档结构", "识别标题层级", "应用排版规则", "生成输出文档"];
  const CONCAT_STEP_LABELS = ["读取两个文档", "保留各自排版", "拼接文档内容", "生成输出文档"];
  const JOB_ID_PATTERN = "[0-9a-f]{32}";
  const CLIENT_JOB_ID_HEADER = "X-Job-ID";
  const PENDING_RESULT_STORAGE_KEY = "zjgsu-paper-formatter.pending-result.v1";
  const DOWNLOAD_URL_PATTERN = new RegExp(`^/api/download/${JOB_ID_PATTERN}_output\\.docx$`);
  const MAX_ERROR_MESSAGE_LENGTH = 240;
  const JOB_CREATION_TIMEOUT_MS = 3 * 60 * 1000;
  const LEGACY_PROCESSING_TIMEOUT_MS = 5 * 60 * 1000;
  const RESULT_FETCH_TIMEOUT_MS = 30 * 1000;
  const RESULT_FETCH_ATTEMPTS = 3;
  const RESULT_FETCH_RETRY_DELAY_MS = 1000;
  const POLL_REQUEST_TIMEOUT_MS = 20 * 1000;
  const POLL_RESULT_ATTEMPTS = 160;
  const POLL_RESULT_DELAY_MS = 3000;
  const POLL_TRANSIENT_FAILURE_LIMIT = 3;
  const PROGRESS_IDLE_TIMEOUT_MS = 5 * 60 * 1000;
  const JOB_WAIT_TIMEOUT_MS = 10 * 60 * 1000;
  const HEALTH_FEATURE_CHECK_TIMEOUT_MS = 4 * 1000;
  const DOCXCOMPOSE_MIN_VERSION = "2.1.0";
  const CONCAT_FEATURE_DISABLED_MESSAGE = "当前部署未启用 docxcompose，拼接文档功能暂不可用。";
  const CONCAT_FEATURE_DISABLED_MESSAGE_MISSING = "当前部署未安装或未启用 docxcompose，拼接文档功能暂不可用。";
  const CONCAT_FEATURE_DISABLED_MESSAGE_VERSION_UNKNOWN = "当前部署未能识别 docxcompose 版本信息，请联系管理员后重试。";
  const CONCAT_FEATURE_DISABLED_MESSAGE_BY_REASON = {
    dependency_missing: CONCAT_FEATURE_DISABLED_MESSAGE_MISSING,
    version_unknown: CONCAT_FEATURE_DISABLED_MESSAGE_VERSION_UNKNOWN,
  };
  const DOCXCOMPOSE_MIN_VERSION_SOURCE_LABELS = {
    runtime: "运行时配置",
    requirements_in: "requirements.in",
    requirements_txt: "requirements.txt",
    unavailable: "未配置",
  };
  // POST 超时后，服务端仍可能在完成上传解析或创建任务；给恢复轮询
  // 独立留出额外窗口，避免恰好跨过创建超时边界时把短暂 404 当成终态。
  const PROVISIONAL_RESULT_GRACE_MS = JOB_CREATION_TIMEOUT_MS + 2 * 60 * 1000;

  function on(el, type, handler, options) {
    if (el) el.addEventListener(type, handler, options);
  }

  function prefersReducedMotion() {
    return typeof window.matchMedia === "function"
      && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  }

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
  let currentProgressIdleTimer = null;
  let currentJobWaitTimer = null;
  let pendingResultUrl = "";
  let pendingResultMode = "format";
  let pendingResultCreatedAt = 0;
  let pendingResultProvisional = false;
  let isSubmitting = false;
  let canCancelRequest = false;
  let serverSupportsConcat = true;
  let docxcomposeSupportReason = null;
  let docxcomposeSupportMinVersion = DOCXCOMPOSE_MIN_VERSION;
  let docxcomposeSupportMinVersionSource = "unavailable";
  let docxcomposeSupportMinVersionSourceLabel = "未配置";

  // ====== 背景粒子 ======
  function initParticles() {
    const container = $("#bgParticles");
    const compactOrTouch = typeof window.matchMedia === "function"
      && window.matchMedia("(max-width: 560px), (hover: none) and (pointer: coarse)").matches;
    if (!container || prefersReducedMotion() || compactOrTouch) return;
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
    const reducedMotion = prefersReducedMotion();
    signs.forEach((el) => {
      const text = el.textContent.trim();
      if (!text) return;
      el.dataset.text = text;
      el.textContent = "";
      const u = document.createElement("span");
      u.className = "sign-underline";
      u.setAttribute("aria-hidden", "true");
      el.appendChild(u);
      if (reducedMotion) {
        el.insertBefore(document.createTextNode(text), u);
        el.classList.add("typed");
      }
    });
    if (reducedMotion) return;
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

  function focusElement(el, scrollToStart = false) {
    if (!el || el.classList.contains("hidden")) return;
    if (scrollToStart) {
      el.focus({ preventScroll: true });
      el.scrollIntoView({ behavior: "auto", block: "start" });
      return;
    }
    el.focus();
  }

  function announceFileSelection(message) {
    if (fileSelectionStatus) fileSelectionStatus.textContent = message || "";
  }

  function setDownloadState(url = "", name = "") {
    downloadUrl = url;
    downloadName = name;
    if (btnDownload) btnDownload.disabled = !downloadUrl;
  }

  function setActiveTab(activeTab) {
    [tabUpload, tabPaste, tabConcat].forEach((tab) => {
      if (!tab) return;
      const isActive = tab === activeTab;
      tab.classList.toggle("active", isActive);
      tab.setAttribute("aria-pressed", isActive ? "true" : "false");
    });
  }

  function normalizeMode(mode) {
    return ["format", "text", "concat"].includes(mode) ? mode : "format";
  }

  function setCurrentMode(mode) {
    currentMode = normalizeMode(mode);
    setActiveTab(currentMode === "concat" ? tabConcat : currentMode === "text" ? tabPaste : tabUpload);
    return currentMode;
  }

  function getConcatFeatureDisabledMessage(reason) {
    const normalizedReason = typeof reason === "string" ? reason.trim() : "";
    const normalizedMinVersion = typeof docxcomposeSupportMinVersion === "string"
      && docxcomposeSupportMinVersion.trim()
      ? docxcomposeSupportMinVersion.trim()
      : DOCXCOMPOSE_MIN_VERSION;
    // Keep the reason lookup explicit so unknown runtime values safely use the
    // generic message instead of being interpolated into the UI.
    const mappedMessage = Object.prototype.hasOwnProperty.call(
      CONCAT_FEATURE_DISABLED_MESSAGE_BY_REASON,
      normalizedReason
    ) ? CONCAT_FEATURE_DISABLED_MESSAGE_BY_REASON[normalizedReason] : null;
    if (normalizedReason === "version_too_old") {
      return `当前部署的 docxcompose 版本低于最低支持版本（>=${normalizedMinVersion}），请先升级后重试。`;
    }
    return typeof mappedMessage === "string" ? mappedMessage : CONCAT_FEATURE_DISABLED_MESSAGE;
  }

  function getDocxcomposeMinVersionSourceHint(source, sourceLabel) {
    const label = getDocxcomposeMinVersionSourceLabel(source, sourceLabel);
    return label ? `（版本来源：${label}）` : "";
  }

  function getDocxcomposeMinVersionSourceLabel(source, sourceLabel) {
    const normalizedSourceLabel = typeof sourceLabel === "string" ? sourceLabel.trim() : "";
    if (normalizedSourceLabel) {
      const normalizedSourceLabelKey = normalizedSourceLabel.toLowerCase();
      if (Object.prototype.hasOwnProperty.call(
        DOCXCOMPOSE_MIN_VERSION_SOURCE_LABELS,
        normalizedSourceLabelKey,
      )) {
        return DOCXCOMPOSE_MIN_VERSION_SOURCE_LABELS[normalizedSourceLabelKey];
      }
      if (Object.values(DOCXCOMPOSE_MIN_VERSION_SOURCE_LABELS).includes(normalizedSourceLabel)) {
        return normalizedSourceLabel;
      }
    }
    const normalizedSource = typeof source === "string" ? source.trim().toLowerCase() : "";
    if (Object.prototype.hasOwnProperty.call(
      DOCXCOMPOSE_MIN_VERSION_SOURCE_LABELS,
      normalizedSource,
    )) {
      return DOCXCOMPOSE_MIN_VERSION_SOURCE_LABELS[normalizedSource];
    }
    return "未知来源";
  }

  function applyConcatAvailabilityState(featureState) {
  const state = asObject(featureState);
  const supported = state.supported !== false;
  const reason = typeof state.reason === "string" ? state.reason : null;
  const minVersion = typeof state.minVersion === "string" ? state.minVersion.trim() : "";
  const minVersionSource = typeof state.minVersionSource === "string" ? state.minVersionSource.trim() : "";
  const minVersionSourceLabel = typeof state.minVersionSourceLabel === "string"
    ? state.minVersionSourceLabel.trim()
    : "";
  if (minVersion) {
    docxcomposeSupportMinVersion = minVersion;
  }
  if (minVersionSource) {
    docxcomposeSupportMinVersionSource = minVersionSource;
  } else {
    docxcomposeSupportMinVersionSource = "unavailable";
  }
  docxcomposeSupportMinVersionSourceLabel = getDocxcomposeMinVersionSourceLabel(
    docxcomposeSupportMinVersionSource,
    minVersionSourceLabel,
  );
  docxcomposeSupportReason = reason;
  serverSupportsConcat = supported;
  const sourceHint = getDocxcomposeMinVersionSourceHint(
    docxcomposeSupportMinVersionSource,
    docxcomposeSupportMinVersionSourceLabel,
  );
  const disabledMessage = supported ? "" : getConcatFeatureDisabledMessage(reason);
  const disabledMessageWithSource = supported ? "" : `${disabledMessage} ${sourceHint}`.trim();
  // Legacy diagnostic expression retained as a stable source boundary:
  // disabledMessageWithSource = `${disabledMessage} ${sourceHint}`.trim();

    if (tabConcat) {
      tabConcat.disabled = isSubmitting || !supported;
      tabConcat.setAttribute("aria-disabled", tabConcat.disabled ? "true" : "false");
      tabConcat.title = supported ? "切换到拼接文档模式" : disabledMessageWithSource;
    }

    if (concatFeatureNotice) {
      concatFeatureNotice.textContent = supported ? "" : disabledMessageWithSource;
      concatFeatureNotice.title = disabledMessageWithSource;
      concatFeatureNotice.classList.toggle("hidden", supported);
    }

    syncBusyState();

    // A late capability response only changes whether a new concat may start.
    // Keep an existing/recoverable job visible, even if this deployment no
    // longer offers concat, so progress, cancellation and result mode survive.
    const preservesConcatContext = isSubmitting
      || (pendingResultUrl && pendingResultMode === "concat");
    if (!supported && currentMode === "concat" && !preservesConcatContext) {
      resetConcatSlots();
      showUpload();
      announceFileSelection("当前部署不支持文档拼接，已自动切换到上传文件模式。");
    }
  }

  function parseHealthConcatFeature(payload) {
  const features = asObject(payload && payload.features);
  const concatSupported = features.concat;
  const reason = asObject(features).docxcompose_support_reason;
  const minVersion = asObject(features).docxcompose_min_version;
  const minVersionSource = asObject(features).docxcompose_min_version_source;
  const minVersionSourceLabel = asObject(features).docxcompose_min_version_source_label;
  // Raw health field contract: reason: asObject(features).docxcompose_support_reason
  if (typeof features.concat === "boolean") {
    return {
      supported: concatSupported,
      reason: typeof reason === "string" ? reason.trim() : null,
      minVersion: typeof minVersion === "string" ? minVersion.trim() : null,
      minVersionSource: typeof minVersionSource === "string" ? minVersionSource.trim() : null,
      minVersionSourceLabel: typeof minVersionSourceLabel === "string"
        ? minVersionSourceLabel.trim()
        : null,
    };
  }
  return { supported: true };
}

  async function fetchServerCapabilities() {
    const controller = new AbortController();
    const timeoutTimer = window.setTimeout(() => controller.abort(), HEALTH_FEATURE_CHECK_TIMEOUT_MS);
    try {
      const response = await fetch("/api/health", {
        cache: "no-store",
        headers: { Accept: "application/json" },
        signal: controller.signal,
      });
      const data = await parseApiResponse(response, "获取服务能力失败，请稍后重试。");
      applyConcatAvailabilityState(parseHealthConcatFeature(data));
    } catch (error) {
      if (error && error.name === "AbortError") return;
      applyConcatAvailabilityState(true);
    } finally {
      window.clearTimeout(timeoutTimer);
    }
  }

  function placeCoverBuilderBefore(button) {
    if (!coverBuilder || !button || !button.parentNode) return;
    if (coverBuilder.nextElementSibling === button) return;
    button.parentNode.insertBefore(coverBuilder, button);
  }

  function showUpload() {
    setCurrentMode("format");
    updateRetryLabel();
    if (selectedFile) { showPreview(); return; }
    showSection(uploadSection);
  }

  function showPaste() {
    setCurrentMode("text");
    updateRetryLabel();
    placeCoverBuilderBefore(btnFormatText);
    showSection(pasteSection);
  }

  function showConcat() {
    if (!serverSupportsConcat) {
      showUpload();
      return;
    }
    setCurrentMode("concat");
    updateRetryLabel();
    showSection(concatSection);
  }

  function showPreview() { updateRetryLabel(); placeCoverBuilderBefore(btnFormat); showSection(filePreview); }

  function createEmptyPreviewMessage(text) {
    const empty = document.createElement("p");
    empty.className = "preview-empty";
    empty.textContent = text;
    return empty;
  }

  function asPreviewItem(item) {
    return item && typeof item === "object" ? item : {};
  }

  function asPreviewText(value, fallback = "") {
    return typeof value === "string" && value.trim() ? value : fallback;
  }

  function asObject(value) {
    return value && typeof value === "object" && !Array.isArray(value) ? value : {};
  }

  function asApiPayload(value, fallback) {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return { success: false, error: fallback };
    }

    const payload = Object.assign({}, value);
    payload.success = value.success === true;
    if (!payload.success) payload.error = normalizeErrorMessage(value.error, fallback);
    return payload;
  }

  function normalizeErrorMessage(value, fallback = "处理失败，请稍后重试。") {
    if (typeof value !== "string") return fallback;

    const message = value.replace(/\s+/g, " ").trim();
    if (!message) return fallback;
    if (message.length <= MAX_ERROR_MESSAGE_LENGTH) return message;

    return message.slice(0, MAX_ERROR_MESSAGE_LENGTH - 1).trimEnd() + "…";
  }

  function normalizeProgressText(value) {
    if (typeof value !== "string") return "";

    const message = value.replace(/\s+/g, " ").trim();
    if (message.length <= MAX_ERROR_MESSAGE_LENGTH) return message;

    return message.slice(0, MAX_ERROR_MESSAGE_LENGTH - 1).trimEnd() + "…";
  }

  function normalizeDownloadUrl(value) {
    if (typeof value !== "string") return "";
    const url = value.trim();
    return DOWNLOAD_URL_PATTERN.test(url) ? url : "";
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
        const previewItem = asPreviewItem(item);
        const card = document.createElement("article");
        card.className = "preview-item";

        const eyebrow = document.createElement("span");
        eyebrow.className = "preview-item-eyebrow";
        eyebrow.textContent = asPreviewText(previewItem.eyebrow, "排版动作");

        const title = document.createElement("h5");
        title.className = "preview-item-title";
        title.textContent = asPreviewText(previewItem.title, "已完成自动处理");

        const description = document.createElement("p");
        description.className = "preview-item-desc";
        description.textContent = asPreviewText(previewItem.description);

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
      const previewItem = asPreviewItem(item);
      const level = asPreviewText(previewItem.level);
      const row = document.createElement("div");
      row.className = "outline-item";
      if (level) {
        row.dataset.level = level;
      }

      const label = document.createElement("span");
      label.className = "outline-level";
      label.textContent = outlineLevelLabels[level] || "结构";

      const text = document.createElement("span");
      text.className = "outline-text";
      text.textContent = asPreviewText(previewItem.text);

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
    focusElement(processingSection, true);
  }

  function showResult(data, mode = currentMode) {
    const resultData = asObject(data);
    const rawDownloadUrl = normalizeDownloadUrl(resultData.download_url);
    if (!rawDownloadUrl) {
      // Keep the pending job so an incomplete response can be fetched again.
      showError("任务结果的下载地址缺失或异常，请重试获取结果。", "结果暂不可下载");
      return;
    }
    stopResultReveal(); completeSteps(); showSection(resultSection);
    const resultMode = setCurrentMode(mode);
    const isConcat = resultMode === "concat";
    const stats = asObject(resultData.stats);
    if (resultTitle) resultTitle.textContent = isConcat ? "拼接完成！" : "排版完成！";
    if (btnDownloadLabel) btnDownloadLabel.textContent = isConcat ? "下载拼接文档" : "下载排版文档";
    if (btnResetLabel) btnResetLabel.textContent = isConcat ? "继续拼接其他文档" : "继续排版其他文档";
    if (outlinePanel) outlinePanel.classList.toggle("hidden", isConcat);
    if (resultPreview) resultPreview.classList.toggle("is-single", isConcat);
    if (highlightsSubtitle) highlightsSubtitle.textContent = isConcat
      ? "保持两个文档各自的排版形态，仅把它们合并为一个文件"
      : "把真正做过的页面设置、页眉页码和结构识别结果直接展示出来";
    if (statInputSize) statInputSize.textContent = asPreviewText(stats.input_size, "-");
    if (statOutputSize) statOutputSize.textContent = asPreviewText(stats.output_size, "-");
    if (statElapsed) statElapsed.textContent = asPreviewText(stats.elapsed, "-");
    const nextDownloadName = asPreviewText(resultData.download_name, isConcat ? "拼接结果.docx" : "排版结果.docx");
    const nextDownloadUrl = rawDownloadUrl + "?name=" + encodeURIComponent(nextDownloadName);
    setDownloadState(nextDownloadUrl, nextDownloadName);
    renderPreview(resultData.preview || resultData.summary || {});
    focusElement(resultSection, true);
  }

  function showError(msg, title = "处理失败") {
    stopResultReveal(); stopStepAnimation(); closeProgressStream(); setDownloadState(); showSection(errorSection);
    if (errorTitle) errorTitle.textContent = title;
    if (errorMessage) errorMessage.textContent = normalizeErrorMessage(msg);
    focusElement(errorSection, true);
  }

  async function parseApiResponse(res, fb) {
    const ct = res.headers.get("content-type") || "";
    const fallback = normalizeErrorMessage(fb);
    try {
      const mimeType = ct.split(";", 1)[0].trim().toLowerCase();
      const isJson = mimeType === "application/json" || mimeType.endsWith("+json");
      // A reverse proxy or older deployment can omit Content-Type even when
      // Flask returned a JSON error.  Try JSON only for an absent header;
      // explicitly labelled HTML/text must still be discarded without
      // buffering a potentially large response body.
      if (!isJson && mimeType) {
        if (res.body && typeof res.body.cancel === "function") await res.body.cancel();
        return { success: false, error: fallback };
      }
      return asApiPayload(await res.json(), fallback);
    } catch (error) {
      if (error && error.name === "AbortError") throw error;
      return { success: false, error: fallback };
    }
  }

  function showPayloadLimitError(data) {
    data = asObject(data);
    showError(data.error || PAYLOAD_LIMIT_MESSAGE);
  }

  // ====== 动画 ======
  function stopResultReveal() { if (resultRevealTimer !== null) { clearTimeout(resultRevealTimer); resultRevealTimer = null; } }
  function stopStepAnimation() { if (stepAnimationTimer !== null) { clearInterval(stepAnimationTimer); stepAnimationTimer = null; } }
  function clearProgressIdleTimer() { if (currentProgressIdleTimer !== null) { window.clearTimeout(currentProgressIdleTimer); currentProgressIdleTimer = null; } }
  function clearJobWaitTimer() { if (currentJobWaitTimer !== null) { window.clearTimeout(currentJobWaitTimer); currentJobWaitTimer = null; } }
  function closeProgressStream() {
    clearProgressIdleTimer();
    if (currentEventSource) { currentEventSource.close(); currentEventSource = null; }
  }

  function syncBusyState() {
    if (processingSection) processingSection.setAttribute("aria-busy", isSubmitting ? "true" : "false");
    [tabUpload, tabPaste, tabConcat, pasteArea, btnFormatText, fileRemove, btnFormat, coverFileRemove, btnDiscardPending].forEach((el) => { if (el) el.disabled = isSubmitting; });
    [coverMetaEnabled, coverTitleInput, collegeInput, teacherInput, classNameInput, studentNameInput, studentIdInput].forEach((el) => { if (el) el.disabled = isSubmitting; });
    if (tabConcat) tabConcat.disabled = isSubmitting || !serverSupportsConcat;
    if (concatRestartPage) concatRestartPage.disabled = isSubmitting || !serverSupportsConcat;
    if (btnCancelRequest) btnCancelRequest.disabled = !isSubmitting || !canCancelRequest;
    if (concatFirstInput) concatFirstInput.disabled = isSubmitting || !serverSupportsConcat;
    if (concatSecondInput) concatSecondInput.disabled = isSubmitting || !serverSupportsConcat;
    if (concatFirstRemove) concatFirstRemove.disabled = isSubmitting || !serverSupportsConcat;
    if (concatSecondRemove) concatSecondRemove.disabled = isSubmitting || !serverSupportsConcat;
    [uploadZone, coverZone, concatZoneFirst, concatZoneSecond].forEach((zone) => {
      if (!zone) return;
      const isBusy = isSubmitting || (zone === concatZoneFirst || zone === concatZoneSecond) && !serverSupportsConcat;
      if (isBusy) zone.classList.remove("drag-over");
      zone.classList.toggle("is-busy", isBusy);
      zone.setAttribute("aria-disabled", isBusy ? "true" : "false");
      zone.tabIndex = isBusy ? -1 : 0;
    });
    updateConcatButton();
  }

  function updateRetryLabel() {
    const hasPendingResult = Boolean(pendingResultUrl);
    if (btnDiscardPending) btnDiscardPending.classList.toggle("hidden", !hasPendingResult);
    if (!retryLabel) return;
    if (hasPendingResult) { retryLabel.textContent = "重新获取任务结果"; return; }
    if (tabConcat && tabConcat.classList.contains("active")) { retryLabel.textContent = selectedConcatFirst || selectedConcatSecond ? "返回继续拼接" : "重新选择文档"; return; }
    if (tabPaste && tabPaste.classList.contains("active")) { retryLabel.textContent = "返回继续编辑"; return; }
    retryLabel.textContent = selectedFile ? "返回继续排版" : "重新上传";
  }

  function showPendingResultConflict() {
    showError(
      "已有尚未领取或确认放弃的任务结果。请先重新获取旧结果，或明确放弃后再提交新任务。",
      "先处理上个任务"
    );
    updateRetryLabel();
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

  function buildCoverMetaPayload() {
    if (!coverMetaEnabled || !coverMetaEnabled.checked) return null;
    const payload = { generate_cover: "1" };
    [
      ["cover_title", coverTitleInput],
      ["college", collegeInput],
      ["teacher", teacherInput],
      ["class_name", classNameInput],
      ["student_name", studentNameInput],
      ["student_id", studentIdInput],
    ].forEach(([key, input]) => {
      const value = input && input.value ? input.value.trim() : "";
      if (value) payload[key] = value;
    });
    return payload;
  }

  function collectCoverMeta(formData) {
    const payload = buildCoverMetaPayload();
    if (!payload || !formData) return;
    Object.entries(payload).forEach(([key, value]) => {
      formData.append(key, value);
    });
  }

  function buildTextFormatPayload(text) {
    return Object.assign({ text }, buildCoverMetaPayload() || {});
  }

  function beginRequest(options = {}) {
    if (isSubmitting) return null;
    if (pendingResultUrl && options.allowPending !== true) {
      showPendingResultConflict();
      return null;
    }
    stopResultReveal(); closeProgressStream(); setDownloadState();
    currentRequestController = new AbortController();
    isSubmitting = true;
    canCancelRequest = options.canCancel !== false;
    syncBusyState();
    return currentRequestController;
  }
  function setRequestCancellable(controller, enabled) {
    if (currentRequestController !== controller) return;
    canCancelRequest = enabled === true;
    syncBusyState();
  }
  function finishRequest(c) {
    if (currentRequestController !== c) return;
    clearJobWaitTimer();
    currentRequestController = null;
    isSubmitting = false;
    canCancelRequest = false;
    syncBusyState();
  }
  function abortActiveRequest() {
    if (currentRequestController) { currentRequestController.abort(); currentRequestController = null; }
    clearJobWaitTimer(); closeProgressStream();
    isSubmitting = false;
    canCancelRequest = false;
    syncBusyState();
  }

  function cancelActiveRequest() {
    if (!isSubmitting || !currentRequestController || !canCancelRequest) return;
    abortActiveRequest();
    showError(
      "已停止当前页面等待。服务器端任务可能仍在继续，请稍后再决定是否重新提交。",
      "已停止等待"
    );
  }

  function armJobWaitTimeout(controller) {
    clearJobWaitTimer();
    currentJobWaitTimer = window.setTimeout(() => {
      currentJobWaitTimer = null;
      if (currentRequestController !== controller) return;
      controller.abort();
      closeProgressStream();
      showError(
        "任务处理时间超过 10 分钟，已停止当前页面等待。服务器端任务可能仍会继续完成。",
        "等待超时"
      );
      finishRequest(controller);
    }, JOB_WAIT_TIMEOUT_MS);
  }

  function setStepState(step, state) {
    if (!step) return;
    const normalizedState = ["active", "done"].includes(state) ? state : "pending";
    step.classList.toggle("active", normalizedState === "active");
    step.classList.toggle("done", normalizedState === "done");
    if (normalizedState === "active") step.setAttribute("aria-current", "step");
    else step.removeAttribute("aria-current");

    const status = step.querySelector(".step-status");
    if (status) {
      status.textContent = normalizedState === "active"
        ? "进行中"
        : normalizedState === "done" ? "已完成" : "未开始";
    }
  }
  function resetSteps() {
    stopStepAnimation();
    steps.forEach((step) => setStepState(step, "pending"));
  }
  function completeSteps() {
    stopStepAnimation();
    steps.forEach((step) => setStepState(step, "done"));
  }
  function setStepProgress(step) {
    stopStepAnimation();
    const normalizedStep = Math.min(Math.max(Math.floor(Number(step)) || 1, 1), steps.length);
    steps.forEach((item, index) => {
      const position = index + 1;
      setStepState(
        item,
        position < normalizedStep
          ? "done"
          : position === normalizedStep ? "active" : "pending"
      );
    });
  }
  function animateSteps() {
    stopStepAnimation(); let cur = 1;
    stepAnimationTimer = setInterval(() => {
      if (cur > 0 && cur <= steps.length) setStepState(steps[cur - 1], "done");
      if (cur < steps.length) {
        setStepState(steps[cur], "active");
        cur++;
      } else {
        stopStepAnimation();
      }
    }, 600);
  }
  function scheduleResult(data, mode = currentMode) {
    stopResultReveal();
    showResult(data, mode);
  }

  function createAbortError() {
    try {
      return new DOMException("Aborted", "AbortError");
    } catch {
      const error = new Error("Aborted");
      error.name = "AbortError";
      return error;
    }
  }

  function createTimeoutError(message) {
    const error = new Error(normalizeErrorMessage(message, "请求等待超时，请稍后重试。"));
    error.name = "TimeoutError";
    return error;
  }

  function handleRequestFailure(error, fallbackMessage) {
    if (error && error.name === "AbortError") return false;
    const timedOut = error && error.name === "TimeoutError";
    showError(
      timedOut ? error.message : fallbackMessage,
      timedOut ? "等待超时" : "处理失败"
    );
    return true;
  }

  async function fetchApiResponse(url, options, fallbackMessage, timeoutMs, timeoutMessage) {
    const requestOptions = Object.assign({}, options || {});
    const parentSignal = requestOptions.signal;
    const timedController = new AbortController();
    let timedOut = false;
    const onParentAbort = () => timedController.abort();

    if (parentSignal) {
      if (parentSignal.aborted) throw createAbortError();
      parentSignal.addEventListener("abort", onParentAbort, { once: true });
    }

    const timeoutTimer = window.setTimeout(() => {
      timedOut = true;
      timedController.abort();
    }, timeoutMs);
    requestOptions.signal = timedController.signal;

    try {
      const res = await fetch(url, requestOptions);
      const data = await parseApiResponse(res, fallbackMessage);
      if (timedOut) throw createTimeoutError(timeoutMessage);
      throwIfAborted(parentSignal);
      return { res, data };
    } catch (error) {
      if (timedOut) throw createTimeoutError(timeoutMessage);
      if (parentSignal && parentSignal.aborted) throw createAbortError();
      throw error;
    } finally {
      window.clearTimeout(timeoutTimer);
      if (parentSignal) parentSignal.removeEventListener("abort", onParentAbort);
    }
  }

  function createResultError(res, data, fallback) {
    const payload = asObject(data);
    const status = Number(res && res.status) || 0;
    const error = new Error(payload.error || fallback);
    error.status = status;
    error.terminal = payload.terminal === true
      || payload.status === "failed"
      || status === 404
      || status === 410;
    // Preserve bounded server guidance when the SSE completion path has to
    // fetch a result and receives a transient HTTP error.  The polling path
    // already reads this header directly; keeping it on the error makes both
    // paths use the same retry contract.
    const retryDelay = retryAfterDelayMs(res, -1);
    if (retryDelay >= 0) error.retryAfterMs = retryDelay;
    return error;
  }

  function isTerminalResultError(error) {
    const status = Number(error && error.status) || 0;
    return Boolean(error && error.terminal === true) || status === 404 || status === 410;
  }

  function clearTerminalPendingResult(error, resultUrl) {
    if (isTerminalResultError(error)) clearPendingJobResult(resultUrl);
  }

  function shouldConfirmPendingJobResponse(res, data) {
    const status = Number(res && res.status) || 0;
    const payload = asObject(data);
    return status > 0 && status !== 404 && (
      // Generic HTTP errors (for example a proxy's 429) do not prove the
      // upload has created a job. Keep its provisional 404 grace period.
      (status === 202 && payload.status === "processing")
      || (status >= 200 && status < 300 && payload.success === true)
      || status === 410
      || payload.terminal === true
      || payload.status === "failed"
    );
  }

  function retryAfterDelayMs(res, fallbackMs) {
    const raw = res && res.headers && typeof res.headers.get === "function"
      ? res.headers.get("retry-after")
      : "";
    const value = typeof raw === "string" ? raw.trim() : "";
    // Retry-After from this API is delta-seconds.  Require the entire value
    // to be numeric: parseFloat("5oops") would otherwise silently turn a
    // malformed proxy header into a five-second client delay.
    if (!/^\d+(?:\.\d+)?$/.test(value)) return fallbackMs;
    const seconds = Number(value);
    if (!Number.isFinite(seconds)) return fallbackMs;
    // Keep a hostile or malformed proxy header from creating an unbounded
    // browser wait while still honoring the server's retry guidance.
    return Math.min(Math.max(seconds * 1000, 100), 10000);
  }

  function hasRetryAfterHeader(res) {
    if (!res || !res.headers || typeof res.headers.get !== "function") return false;
    const value = res.headers.get("retry-after");
    return typeof value === "string" && value.trim() !== "";
  }

  async function fetchAsyncJobResult(resultUrl, signal) {
    const fallback = "获取排版结果失败，请稍后重试。";
    for (let attempt = 0; attempt < RESULT_FETCH_ATTEMPTS; attempt++) {
      try {
        const { res, data } = await fetchApiResponse(
          resultUrl,
          { signal, cache: "no-store" },
          fallback,
          RESULT_FETCH_TIMEOUT_MS,
          "获取排版结果超时，请检查网络后重试。"
        );
        if (shouldConfirmPendingJobResponse(res, data)) markPendingResultConfirmed(resultUrl);
        if (res.status === 202 && data.status === "processing") {
          if (attempt === RESULT_FETCH_ATTEMPTS - 1) {
            throw createTimeoutError("任务仍在处理中，请稍后重试获取结果。");
          }
          await delay(retryAfterDelayMs(res, RESULT_FETCH_RETRY_DELAY_MS * (2 ** attempt)), signal);
          continue;
        }
        if (res.status === 429 && hasRetryAfterHeader(res) && attempt < RESULT_FETCH_ATTEMPTS - 1) {
          await delay(retryAfterDelayMs(res, RESULT_FETCH_RETRY_DELAY_MS * (2 ** attempt)), signal);
          continue;
        }
        if (!res.ok || !data.success) {
          throw createResultError(res, data, fallback);
        }
        return data;
      } catch (error) {
        if (error && error.name === "AbortError") throw error;
        const status = Number(error && error.status) || 0;
        const retryable = !isTerminalResultError(error) && (status === 0 || status >= 500);
        if (!retryable || attempt === RESULT_FETCH_ATTEMPTS - 1) throw error;
        const fallbackDelay = RESULT_FETCH_RETRY_DELAY_MS * (2 ** attempt);
        const retryDelay = Number.isFinite(error && error.retryAfterMs)
          ? error.retryAfterMs
          : fallbackDelay;
        await delay(retryDelay, signal);
      }
    }
    throw new Error(fallback);
  }

  function throwIfAborted(signal) {
    if (signal && signal.aborted) throw createAbortError();
  }

  function delay(ms, signal) {
    return new Promise((resolve, reject) => {
      if (signal && signal.aborted) {
        reject(createAbortError());
        return;
      }

      let timer = null;
      const cleanup = () => {
        if (signal) signal.removeEventListener("abort", onAbort);
      };
      const onAbort = () => {
        if (timer !== null) window.clearTimeout(timer);
        cleanup();
        reject(createAbortError());
      };

      timer = window.setTimeout(() => {
        cleanup();
        resolve();
      }, ms);

      if (signal) signal.addEventListener("abort", onAbort, { once: true });
    });
  }

  async function pollAsyncJobResult(
    resultUrl,
    signal,
    attempts = POLL_RESULT_ATTEMPTS,
    delayMs = POLL_RESULT_DELAY_MS
  ) {
    const fallback = "获取排版结果失败，请稍后重试。";
    let consecutiveTransientFailures = 0;
    for (let attempt = 0; attempt < attempts; attempt++) {
      throwIfAborted(signal);
      let response;
      try {
        response = await fetchApiResponse(
          resultUrl,
          { signal, cache: "no-store" },
          fallback,
          POLL_REQUEST_TIMEOUT_MS,
          "读取任务状态超时，请检查网络后重试。"
        );
      } catch (error) {
        if (error && error.name === "AbortError") throw error;
        consecutiveTransientFailures += 1;
        if (
          consecutiveTransientFailures >= POLL_TRANSIENT_FAILURE_LIMIT
          || attempt === attempts - 1
        ) throw error;
        await delay(delayMs, signal);
        continue;
      }
      const { res, data } = response;
      if (res.status === 404 && isPendingResultWithinCreationGrace(resultUrl)) {
        consecutiveTransientFailures = 0;
        setProcessingLive("正在确认任务是否已经创建...");
        if (attempt === attempts - 1) {
          throw createTimeoutError("等待任务创建超时，已停止当前页面等待。");
        }
        await delay(delayMs, signal);
        continue;
      }
      if (shouldConfirmPendingJobResponse(res, data)) markPendingResultConfirmed(resultUrl);
      if (!res.ok && (data.terminal === true || data.status === "failed")) {
        throw createResultError(res, data, fallback);
      }
      if (res.status === 202 && data.status === "processing") {
        consecutiveTransientFailures = 0;
        if (attempt === attempts - 1) {
          throw createTimeoutError("等待任务完成超时，已停止当前页面等待。");
        }
        // Honor the server's retry guidance when a proxy or a resumed worker
        // needs more time than the normal polling cadence.  Keep the delay
        // bounded so a hostile header cannot freeze the recovery UI.
        await delay(retryAfterDelayMs(res, delayMs), signal);
        continue;
      }
      if (res.status === 429 && hasRetryAfterHeader(res) && attempt < attempts - 1) {
        consecutiveTransientFailures = 0;
        await delay(retryAfterDelayMs(res, delayMs), signal);
        continue;
      }
      if (res.status >= 500) {
        consecutiveTransientFailures += 1;
      } else {
        consecutiveTransientFailures = 0;
      }
      if (
        res.status >= 500
        && consecutiveTransientFailures < POLL_TRANSIENT_FAILURE_LIMIT
        && attempt < attempts - 1
      ) {
        // A draining worker can deliberately return 503 with Retry-After.
        // Respect that bounded server guidance instead of immediately
        // hammering the recovering instance with the normal poll cadence.
        await delay(retryAfterDelayMs(res, delayMs), signal);
        continue;
      }
      if (!res.ok || !data.success) {
        throw createResultError(res, data, fallback);
      }
      return data;
    }
    throw new Error(fallback);
  }

  function handleProgressEvent(payload) {
    if (payload && typeof payload.step === "number") setStepProgress(payload.step);
    const message = normalizeProgressText(payload && payload.message);
    const detail = normalizeProgressText(payload && payload.detail);
    const liveMessage = message && detail ? `${message} · ${detail}` : message || detail;
    setProcessingLive(liveMessage || "服务器正在处理文档...");
  }

  function acknowledgeFailedJob(resultUrl) {
    // Failure acknowledgement is best effort, but must use the same bounded
    // request/body handling as result recovery so a stalled proxy cannot leave
    // an unbounded request alive after the task UI has already finished.
    void fetchApiResponse(
      resultUrl,
      { cache: "no-store", headers: { Accept: "application/json" } },
      "确认失败任务状态失败。",
      RESULT_FETCH_TIMEOUT_MS,
      "确认失败任务状态超时。"
    ).catch(() => {});
  }

  function ownsActiveRequest(controller, source = null) {
    if (!controller || controller.signal.aborted || currentRequestController !== controller) return false;
    return !source || currentEventSource === source;
  }

  function pollProgressResult(resultUrl, controller, fallbackMessage) {
    pollAsyncJobResult(resultUrl, controller.signal)
      .then((result) => {
        if (!ownsActiveRequest(controller)) return;
        completeSteps();
        scheduleResult(result, currentMode);
      })
      .catch((err) => {
        if ((err && err.name === "AbortError") || !ownsActiveRequest(controller)) return;
        clearTerminalPendingResult(err, resultUrl);
        showError(
          (err && err.message) || fallbackMessage,
          err && err.name === "TimeoutError" ? "等待超时" : "处理失败"
        );
      })
      .finally(() => {
        finishRequest(controller);
      });
  }

  function openProgressStream(eventsUrl, resultUrl, controller) {
    closeProgressStream();
    if (typeof EventSource === "undefined") {
      setProcessingLive("实时进度不可用，正在轮询结果...");
      pollProgressResult(resultUrl, controller, "进度连接不可用，请检查网络后重试。");
      return;
    }
    let isTerminal = false;
    let source = null;
    try {
      source = new EventSource(eventsUrl);
    } catch {
      setProcessingLive("实时进度不可用，正在轮询结果...");
      pollProgressResult(resultUrl, controller, "进度连接不可用，请检查网络后重试。");
      return;
    }
    currentEventSource = source;

    const resetProgressIdleTimeout = () => {
      if (isTerminal || !ownsActiveRequest(controller, source)) return;
      clearProgressIdleTimer();
      currentProgressIdleTimer = window.setTimeout(() => {
        currentProgressIdleTimer = null;
        if (isTerminal || !ownsActiveRequest(controller, source)) return;
        isTerminal = true;
        closeProgressStream();
        setProcessingLive("长时间未收到新进度，正在轮询任务结果...");
        pollProgressResult(resultUrl, controller, "长时间未收到任务进度，请检查网络后重试。");
      }, PROGRESS_IDLE_TIMEOUT_MS);
    };
    resetProgressIdleTimeout();
    source.addEventListener("open", () => {
      if (!isTerminal && ownsActiveRequest(controller, source)) resetProgressIdleTimeout();
    });

    source.addEventListener("progress", (event) => {
      if (isTerminal || !ownsActiveRequest(controller, source)) return;
      resetProgressIdleTimeout();
      try {
        handleProgressEvent(JSON.parse(event.data || "{}"));
      } catch {
        setProcessingLive("服务器正在处理文档...");
      }
    });

    source.addEventListener("complete", async () => {
      if (isTerminal || !ownsActiveRequest(controller, source)) return;
      isTerminal = true;
      closeProgressStream();
      setProcessingLive("排版完成，正在整理结果...");
      try {
        const result = await fetchAsyncJobResult(resultUrl, controller.signal);
        if (!ownsActiveRequest(controller)) return;
        completeSteps();
        scheduleResult(result, currentMode);
      } catch (err) {
        if ((err && err.name === "AbortError") || !ownsActiveRequest(controller)) return;
        clearTerminalPendingResult(err, resultUrl);
        showError(
          (err && err.message) || "获取排版结果失败，请稍后重试。",
          err && err.name === "TimeoutError" ? "等待超时" : "处理失败"
        );
      } finally {
        finishRequest(controller);
      }
    });

    source.addEventListener("failed", (event) => {
      if (isTerminal || !ownsActiveRequest(controller, source)) return;
      isTerminal = true;
      closeProgressStream();
      let message = "排版处理失败，请稍后重试。";
      try {
        const payload = JSON.parse(event.data || "{}");
        message = normalizeErrorMessage(payload.message, message);
      } catch {}
      showError(message);
      acknowledgeFailedJob(resultUrl);
      clearPendingJobResult(resultUrl);
      finishRequest(controller);
    });

    source.onerror = () => {
      if (isTerminal || !ownsActiveRequest(controller, source)) return;
      isTerminal = true;
      closeProgressStream();
      setProcessingLive("进度连接中断，正在尝试获取结果...");
      pollProgressResult(resultUrl, controller, "进度连接中断，请检查网络后重试。");
    };
  }

  function normalizeAsyncJobUrl(value, expectedSuffix) {
    if (typeof value !== "string") return "";
    const url = value.trim();
    if (expectedSuffix !== "events" && expectedSuffix !== "result") return "";
    const pattern = new RegExp(`^/api/jobs/${JOB_ID_PATTERN}/${expectedSuffix}$`);
    return pattern.test(url) ? url : "";
  }

  function createClientJobId() {
    const cryptoApi = window.crypto;
    if (!cryptoApi || typeof cryptoApi.getRandomValues !== "function") return "";
    const bytes = new Uint8Array(16);
    cryptoApi.getRandomValues(bytes);
    return Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
  }

  function persistPendingJobResult() {
    try {
      window.sessionStorage.setItem(
        PENDING_RESULT_STORAGE_KEY,
        JSON.stringify({
          result_url: pendingResultUrl,
          mode: pendingResultMode,
          created_at: pendingResultCreatedAt,
          provisional: pendingResultProvisional,
        })
      );
    } catch {}
  }

  function setPendingJobResult(resultUrl, mode = currentMode, metadata = {}) {
    const normalizedUrl = normalizeAsyncJobUrl(resultUrl, "result");
    if (!normalizedUrl) return false;
    const normalizedMode = normalizeMode(mode);
    const requestedCreatedAt = Number(metadata.createdAt);
    pendingResultUrl = normalizedUrl;
    pendingResultMode = normalizedMode;
    pendingResultCreatedAt = Number.isFinite(requestedCreatedAt) && requestedCreatedAt > 0
      ? requestedCreatedAt
      : Date.now();
    pendingResultProvisional = metadata.provisional === true;
    persistPendingJobResult();
    updateRetryLabel();
    return true;
  }

  function isPendingResultWithinCreationGrace(resultUrl) {
    if (!pendingResultProvisional || resultUrl !== pendingResultUrl) return false;
    const ageMs = Date.now() - pendingResultCreatedAt;
    return ageMs >= 0 && ageMs <= PROVISIONAL_RESULT_GRACE_MS;
  }

  function markPendingResultConfirmed(resultUrl) {
    if (resultUrl !== pendingResultUrl || !pendingResultProvisional) return;
    pendingResultProvisional = false;
    persistPendingJobResult();
  }

  function clearPendingJobResult(expectedResultUrl = "") {
    if (expectedResultUrl && pendingResultUrl && expectedResultUrl !== pendingResultUrl) return false;
    pendingResultUrl = "";
    pendingResultMode = "format";
    pendingResultCreatedAt = 0;
    pendingResultProvisional = false;
    try { window.sessionStorage.removeItem(PENDING_RESULT_STORAGE_KEY); } catch {}
    updateRetryLabel();
    return true;
  }

  function restorePendingJobResult() {
    let stored = null;
    try { stored = JSON.parse(window.sessionStorage.getItem(PENDING_RESULT_STORAGE_KEY) || "null"); } catch {}
    const data = asObject(stored);
    // A provisional id is only a short-lived recovery hint while the upload
    // request may still be creating its server-side job.  Once that grace
    // window has elapsed, retaining it would permanently block new uploads
    // after a browser restart even though the client can no longer reconcile
    // the fabricated id with a real task.
    if (data.provisional === true) {
      const createdAt = Number(data.created_at);
      const ageMs = Date.now() - createdAt;
      if (!Number.isFinite(createdAt) || ageMs < 0 || ageMs > PROVISIONAL_RESULT_GRACE_MS) {
        try { window.sessionStorage.removeItem(PENDING_RESULT_STORAGE_KEY); } catch {}
        return false;
      }
    }
    const restored = setPendingJobResult(data.result_url, data.mode, {
      createdAt: data.created_at,
      provisional: data.provisional === true,
    });
    if (restored) setCurrentMode(pendingResultMode);
    return restored;
  }

  async function startAsyncJob(endpoint, fetchOptions, controller, fallbackMessage) {
    const requestOptions = Object.assign({}, fetchOptions || {});
    if (controller && controller.signal && !requestOptions.signal) {
      requestOptions.signal = controller.signal;
    }
    if (!requestOptions.cache) requestOptions.cache = "no-store";

    const clientJobId = createClientJobId();
    const provisionalResultUrl = clientJobId ? `/api/jobs/${clientJobId}/result` : "";
    if (clientJobId) {
      requestOptions.headers = Object.assign({}, requestOptions.headers || {}, {
        [CLIENT_JOB_ID_HEADER]: clientJobId,
      });
      setPendingJobResult(provisionalResultUrl, currentMode, { provisional: true });
    }

    const { res, data } = await fetchApiResponse(
      endpoint,
      requestOptions,
      fallbackMessage,
      JOB_CREATION_TIMEOUT_MS,
      "上传或创建任务超时，请检查网络后重试。"
    );
    if (res.status === 413) {
      if (provisionalResultUrl) clearPendingJobResult(provisionalResultUrl);
      showPayloadLimitError(data);
      return false;
    }
    if (!res.ok || !data.success) {
      const responseResultUrl = normalizeAsyncJobUrl(data.result_url, "result");
      const hasRecoverableJob = Boolean(
        provisionalResultUrl
        && data.job_created === true
        && responseResultUrl === provisionalResultUrl
      );
      if (hasRecoverableJob) {
        setPendingJobResult(responseResultUrl, currentMode, { provisional: false });
      } else if (
        provisionalResultUrl
        && (
          (data.job_created === false)
          || (res.status >= 400 && res.status < 500)
        )
      ) {
        clearPendingJobResult(provisionalResultUrl);
      }
      showError(data.error || fallbackMessage);
      return false;
    }

    const eventsUrl = normalizeAsyncJobUrl(data.events_url, "events");
    const resultUrl = normalizeAsyncJobUrl(data.result_url, "result");
    if (!eventsUrl || !resultUrl) {
      // The provisional client id is only recoverable while the server has
      // returned a valid task URL.  Leaving it persisted here would make the
      // retry action poll a fabricated job for the full creation grace window
      // and block the user from submitting a new document.
      if (provisionalResultUrl) clearPendingJobResult(provisionalResultUrl);
      showError("任务已创建，但进度地址异常，请刷新后重试。");
      return false;
    }

    setPendingJobResult(resultUrl, currentMode, { provisional: false });
    setRequestCancellable(controller, true);
    setProcessingLive("任务已创建，正在连接实时进度流...");
    armJobWaitTimeout(controller);
    openProgressStream(eventsUrl, resultUrl, controller);
    return true;
  }

  // ====== 工具 ======
  // 上传上限：与后端 app.py 的 MAX_CONTENT_LENGTH 保持一致（常驻服务器无 4.5MB 平台限制），
  // 留少量余量给 multipart 开销；同时保留对 413 的兜底处理，兼容更小限制的平台。
  const MAX_UPLOAD_BYTES = 48 * 1024 * 1024;
  const PAYLOAD_LIMIT_MESSAGE = "文档体积超过上传限制（建议不超过 48MB，服务器上限约 50MB）。请压缩正文中的图片后再传。";
  const MAX_TEXT_INPUT_BYTES = 2 * 1024 * 1024;
  const TEXT_INPUT_TOO_LARGE_MESSAGE = "文本内容超过服务器单次处理上限（约 2MB），请拆分后再试。";
  const MAX_TEXT_PARAGRAPHS = 10000;
  const TEXT_PARAGRAPH_LIMIT_MESSAGE = "文本段落数量超过服务器单次处理上限（10000 段），请删除多余空行或拆分后再试。";
  function formatSize(b) { if (b < 1024) return b + " B"; if (b < 1048576) return (b / 1024).toFixed(1) + " KB"; return (b / 1048576).toFixed(2) + " MB"; }
  function utf8ByteLength(text) {
    return new Blob([text]).size;
  }
  function textExceedsParagraphLimit(text) {
    let paragraphCount = 1;
    let pendingBreaks = 0;

    for (let index = 0; index < text.length; index += 1) {
      const char = text[index];
      if (char === "\r") {
        if (text[index + 1] === "\n") index += 1;
        pendingBreaks = Math.min(MAX_TEXT_PARAGRAPHS + 1, pendingBreaks + 1);
      } else if (char === "\n") {
        pendingBreaks = Math.min(MAX_TEXT_PARAGRAPHS + 1, pendingBreaks + 1);
      } else if (pendingBreaks) {
        paragraphCount += pendingBreaks;
        if (paragraphCount > MAX_TEXT_PARAGRAPHS) return true;
        pendingBreaks = 0;
      }
    }

    return false;
  }
  function exceedsUploadLimit(totalBytes) {
    if (totalBytes > MAX_UPLOAD_BYTES) {
      showError(`两个文档合计约 ${formatSize(totalBytes)}，${PAYLOAD_LIMIT_MESSAGE}`);
      return true;
    }
    return false;
  }
  function exceedsTextLimit(text) {
    if (utf8ByteLength(text) > MAX_TEXT_INPUT_BYTES) {
      showError(TEXT_INPUT_TOO_LARGE_MESSAGE);
      return true;
    }
    if (textExceedsParagraphLimit(text)) {
      showError(TEXT_PARAGRAPH_LIMIT_MESSAGE);
      return true;
    }
    return false;
  }
  function validateDocx(f) {
    if (!f.name.toLowerCase().endsWith(".docx")) { showError("仅支持 .docx 格式的 Word 文档"); return false; }
    if (f.size <= 0) { showError("该文档内容为空，请重新选择有效的 .docx 文档。"); return false; }
    if (f.size > MAX_UPLOAD_BYTES) { showError(`该文档约 ${formatSize(f.size)}，${PAYLOAD_LIMIT_MESSAGE}`); return false; }
    return true;
  }

  // ====== 文件选择 ======
  function selectFile(file) {
    if (!file || !validateDocx(file)) return;
    selectedFile = file;
    if (fileName) fileName.textContent = file.name;
    if (fileSize) fileSize.textContent = formatSize(file.size);
    updateRetryLabel(); updateFormatButton(); showPreview();
    focusElement(fileCard);
    announceFileSelection(`已选择待排版文档：${file.name}，大小 ${formatSize(file.size)}。`);
  }

  // ====== 封面选择 ======
  function selectCover(file) {
    if (!file || !validateDocx(file)) return;
    selectedCover = file;
    if (coverFileName) coverFileName.textContent = file.name;
    if (coverFileSize) coverFileSize.textContent = formatSize(file.size);
    if (coverZone) coverZone.classList.add("hidden");
    if (coverFileCard) coverFileCard.classList.remove("hidden");
    updateFormatButton();
    focusElement(coverFileCard);
    announceFileSelection(`已选择封面文档：${file.name}，大小 ${formatSize(file.size)}。`);
  }

  function removeCover() {
    const removedName = selectedCover && selectedCover.name;
    selectedCover = null;
    if (coverInput) coverInput.value = "";
    if (coverZone) coverZone.classList.remove("hidden");
    if (coverFileCard) coverFileCard.classList.add("hidden");
    updateFormatButton();
    focusElement(coverZone);
    announceFileSelection(removedName ? `已移除封面文档：${removedName}。` : "已移除封面文档。");
  }

  // ====== 拼接文档：文件选择 ======
  function updateConcatButton() {
    if (!btnConcat) return;
    btnConcat.disabled = isSubmitting || !serverSupportsConcat || !selectedConcatFirst || !selectedConcatSecond;
  }

  function selectConcatFile(slot, file) {
    if (!file || !validateDocx(file)) return;
    const isFirst = slot === "first";
    if (isFirst) selectedConcatFirst = file; else selectedConcatSecond = file;
    const nameEl = isFirst ? concatFirstName : concatSecondName;
    const sizeEl = isFirst ? concatFirstSize : concatSecondSize;
    const zoneEl = isFirst ? concatZoneFirst : concatZoneSecond;
    const cardEl = isFirst ? concatFirstCard : concatSecondCard;
    if (nameEl) nameEl.textContent = file.name;
    if (sizeEl) sizeEl.textContent = formatSize(file.size);
    if (zoneEl) zoneEl.classList.add("hidden");
    if (cardEl) cardEl.classList.remove("hidden");
    updateConcatButton();
    focusElement(cardEl);
    announceFileSelection(`已选择${isFirst ? "第一个" : "第二个"}拼接文档：${file.name}，大小 ${formatSize(file.size)}。`);
  }

  function removeConcatFile(slot) {
    const isFirst = slot === "first";
    const removedFile = isFirst ? selectedConcatFirst : selectedConcatSecond;
    if (isFirst) { selectedConcatFirst = null; if (concatFirstInput) concatFirstInput.value = ""; }
    else { selectedConcatSecond = null; if (concatSecondInput) concatSecondInput.value = ""; }
    const zoneEl = isFirst ? concatZoneFirst : concatZoneSecond;
    const cardEl = isFirst ? concatFirstCard : concatSecondCard;
    if (zoneEl) zoneEl.classList.remove("hidden");
    if (cardEl) cardEl.classList.add("hidden");
    updateConcatButton();
    focusElement(zoneEl);
    announceFileSelection(
      removedFile && removedFile.name
        ? `已移除${isFirst ? "第一个" : "第二个"}拼接文档：${removedFile.name}。`
        : `已移除${isFirst ? "第一个" : "第二个"}拼接文档。`
    );
  }

  function resetConcatSlots() {
    selectedConcatFirst = null;
    selectedConcatSecond = null;
    if (concatFirstInput) concatFirstInput.value = "";
    if (concatSecondInput) concatSecondInput.value = "";
    if (concatFirstName) concatFirstName.textContent = "";
    if (concatSecondName) concatSecondName.textContent = "";
    if (concatFirstSize) concatFirstSize.textContent = "";
    if (concatSecondSize) concatSecondSize.textContent = "";
    if (concatRestartPage) concatRestartPage.checked = true;
    [concatZoneFirst, concatZoneSecond].forEach((z) => { if (z) z.classList.remove("hidden"); });
    [concatFirstCard, concatSecondCard].forEach((c) => { if (c) c.classList.add("hidden"); });
    updateConcatButton();
  }

  // ====== 重置 ======
  function resetCoverMetadata() {
    if (coverMetaEnabled) coverMetaEnabled.checked = false;
    [coverTitleInput, collegeInput, teacherInput, classNameInput, studentNameInput, studentIdInput]
      .forEach((input) => { if (input) input.value = ""; });
    toggleCoverMetaPanel();
  }

  function resetAll() {
    abortActiveRequest(); stopResultReveal(); stopStepAnimation();
    clearPendingJobResult();
    selectedFile = null; selectedCover = null;
    setDownloadState();
    if (pasteArea) pasteArea.value = "";
    if (fileInput) fileInput.value = "";
    if (coverInput) coverInput.value = "";
    if (fileName) fileName.textContent = "";
    if (fileSize) fileSize.textContent = "";
    if (coverFileName) coverFileName.textContent = "";
    if (coverFileSize) coverFileSize.textContent = "";
    if (coverZone) coverZone.classList.remove("hidden");
    if (coverFileCard) coverFileCard.classList.add("hidden");
    resetCoverMetadata();
    resetConcatSlots();
    updateRetryLabel(); updateFormatButton();
    if (currentMode === "concat") showConcat();
    else if (currentMode === "text") showPaste();
    else showUpload();
    focusCurrentInput();
  }

  function focusCurrentInput() {
    if (currentMode === "text") {
      focusElement(pasteArea);
      return;
    }
    if (currentMode === "concat") {
      if (!selectedConcatFirst) focusElement(concatZoneFirst);
      else if (!selectedConcatSecond) focusElement(concatZoneSecond);
      else focusElement(btnConcat);
      return;
    }
    focusElement(selectedFile ? btnFormat : uploadZone);
  }

  function startNewTask() {
    resetAll();
  }

  function removeMainFile() {
    if (pendingResultUrl) {
      showPendingResultConflict();
      return;
    }
    const removedName = selectedFile && selectedFile.name;
    selectedFile = null;
    setDownloadState();
    if (fileInput) fileInput.value = "";
    if (fileName) fileName.textContent = "";
    if (fileSize) fileSize.textContent = "";
    updateRetryLabel(); updateFormatButton();
    showUpload();
    focusElement(uploadZone);
    announceFileSelection(removedName ? `已移除待排版文档：${removedName}。` : "已移除待排版文档。");
  }

  function discardPendingResult() {
    if (!pendingResultUrl) {
      returnToCurrentInputs();
      return;
    }
    clearPendingJobResult();
    setDownloadState();
    updateFormatButton();
    if (currentMode === "concat") showConcat();
    else if (currentMode === "text") showPaste();
    else showUpload();
    focusCurrentInput();
  }

  function returnToCurrentInputs() {
    if (pendingResultUrl) {
      retryPendingJobResult();
      return;
    }
    abortActiveRequest(); stopResultReveal(); stopStepAnimation();
    setDownloadState();
    updateRetryLabel(); updateFormatButton();
    if (currentMode === "concat") showConcat();
    else if (currentMode === "text") showPaste();
    else showUpload();
    focusCurrentInput();
  }

  async function retryPendingJobResult() {
    if (!pendingResultUrl || isSubmitting) return;
    const resultUrl = pendingResultUrl;
    setCurrentMode(pendingResultMode);
    const controller = beginRequest({ allowPending: true });
    if (!controller) return;
    showProcessing();
    setProcessingLive("正在重新获取已完成任务的结果...");
    armJobWaitTimeout(controller);
    try {
      const result = await pollAsyncJobResult(resultUrl, controller.signal);
      completeSteps();
      scheduleResult(result, currentMode);
    } catch (error) {
      if (error && error.name === "AbortError") return;
      clearTerminalPendingResult(error, resultUrl);
      showError(
        (error && error.message) || "获取排版结果失败，请稍后重试。",
        error && error.name === "TimeoutError" ? "等待超时" : "处理失败"
      );
    } finally {
      finishRequest(controller);
    }
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
      const { res, data } = await fetchApiResponse(
        endpoint,
        { method: "POST", body: formData, signal: controller.signal, cache: "no-store" },
        "排版处理失败，请稍后重试。",
        LEGACY_PROCESSING_TIMEOUT_MS,
        "上传和排版等待超时，请检查网络后重试。"
      );
      if (res.status === 413) { showPayloadLimitError(data); return; }
      if (!res.ok || !data.success) { showError(data.error || "排版处理失败"); return; }
      completeSteps(); scheduleResult(data);
    } catch (err) {
      handleRequestFailure(err, "网络连接失败，请检查网络后重试。");
    } finally { finishRequest(controller); }
  }

  async function uploadAndFormat() {
    if (!selectedFile || isSubmitting) return;
    if (selectedCover && exceedsUploadLimit(selectedFile.size + selectedCover.size)) return;

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
      handleRequestFailure(err, "网络连接失败，请检查网络后重试。");
    } finally {
      if (!handedOff) finishRequest(controller);
    }
  }

  // ====== 文字排版 ======
  async function submitPastedTextLegacy() {
    const text = pasteArea ? pasteArea.value.trim() : "";
    if (!text) { showError("文本内容不能为空"); return; }
    if (exceedsTextLimit(text)) return;
    if (isSubmitting) return;
    const controller = beginRequest();
    if (!controller) return;
    showProcessing(true);
    try {
      const { res, data } = await fetchApiResponse(
        "/api/format_text",
        { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(buildTextFormatPayload(text)), signal: controller.signal, cache: "no-store" },
        "排版失败",
        LEGACY_PROCESSING_TIMEOUT_MS,
        "排版等待超时，请检查网络后重试。"
      );
      if (res.status === 413) { showPayloadLimitError(data); return; }
      if (!res.ok || !data.success) { showError(data.error || "排版失败"); return; }
      completeSteps(); scheduleResult(data);
    } catch (err) {
      handleRequestFailure(err, "网络连接失败，请检查网络后重试。");
    } finally { finishRequest(controller); }
  }

  async function submitPastedText() {
    const text = pasteArea ? pasteArea.value.trim() : "";
    if (!text) { showError("文本内容不能为空"); return; }
    if (exceedsTextLimit(text)) return;
    if (isSubmitting) return;

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
          body: JSON.stringify(buildTextFormatPayload(text)),
          signal: controller.signal,
        },
        controller,
        "排版任务创建失败，请稍后重试。"
      );
    } catch (err) {
      handleRequestFailure(err, "网络连接失败，请检查网络后重试。");
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
      const { res, data } = await fetchApiResponse(
        "/api/concat",
        { method: "POST", body: formData, signal: controller.signal, cache: "no-store" },
        "文档拼接失败，请稍后重试。",
        LEGACY_PROCESSING_TIMEOUT_MS,
        "上传和拼接等待超时，请检查网络后重试。"
      );
      if (res.status === 413) { showPayloadLimitError(data); return; }
      if (!res.ok || !data.success) { showError(data.error || "文档拼接失败"); return; }
      completeSteps(); scheduleResult(data);
    } catch (err) {
      handleRequestFailure(err, "网络连接失败，请检查网络后重试。");
    } finally { finishRequest(controller); }
  }

  async function submitConcat() {
    if (!serverSupportsConcat) {
      showError(getConcatFeatureDisabledMessage(docxcomposeSupportReason), "功能不可用");
      return;
    }
    if (!selectedConcatFirst || !selectedConcatSecond) { showError("请先分别选择两个 .docx 文档"); return; }
    if (isSubmitting) return;
    if (exceedsUploadLimit(selectedConcatFirst.size + selectedConcatSecond.size)) return;
    currentMode = "concat";

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
      handleRequestFailure(err, "网络连接失败，请检查网络后重试。");
    } finally {
      if (!handedOff) finishRequest(controller);
    }
  }

  // ====== 事件绑定 ======
  // Legacy marker retained for consumers that locate this stable binding by
  // its original two-argument signature: function bindFilePicker(zone, input)
  function bindFilePicker(zone, input, featureRequired = false) {
    if (!zone || !input) return;
    // The hidden input lives inside the role=button zone. A programmatic
    // input.click() dispatches a bubbling click event; stop it here or the
    // zone handler re-enters openPicker recursively (notably on mobile and
    // keyboard activation).
    on(input, "click", (event) => event.stopPropagation());
    const openPicker = () => {
      if (isSubmitting || (featureRequired && !serverSupportsConcat)) return;
      // 清空原生值，使校验失败后重选同一文件仍会触发 change。
      input.value = "";
      input.click();
    };
    on(zone, "click", openPicker);
    on(zone, "keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      openPicker();
    });
  }

  on(tabUpload, "click", showUpload);
  on(tabPaste, "click", showPaste);
  on(tabConcat, "click", showConcat);
  on(btnFormatText, "click", submitPastedText);
  bindFilePicker(uploadZone, fileInput);
  on(fileInput, "change", (e) => {
    if (uploadZone) uploadZone.classList.remove("drag-over");
    const files = e.target && e.target.files;
    if (files && files.length > 0) selectFile(files[0]);
  });

  on(uploadZone, "dragover", (e) => { e.preventDefault(); if (!isSubmitting) uploadZone.classList.add("drag-over"); });
  on(uploadZone, "dragleave", (e) => { e.preventDefault(); uploadZone.classList.remove("drag-over"); });
  on(uploadZone, "drop", (e) => {
    e.preventDefault();
    uploadZone.classList.remove("drag-over");
    const files = e.dataTransfer && e.dataTransfer.files;
    if (!isSubmitting && files && files.length > 0) selectFile(files[0]);
  });

  on(fileRemove, "click", removeMainFile);
  on(btnFormat, "click", uploadAndFormat);

  // 封面事件
  bindFilePicker(coverZone, coverInput);
  on(coverInput, "change", (e) => {
    const files = e.target && e.target.files;
    if (files && files.length > 0) selectCover(files[0]);
  });
  on(coverFileRemove, "click", removeCover);
  on(coverMetaEnabled, "change", toggleCoverMetaPanel);
  [coverTitleInput, collegeInput, teacherInput, classNameInput, studentNameInput, studentIdInput].forEach((input) => {
    if (input) input.addEventListener("input", updateFormatButton);
  });

  // ====== 拼接文档事件 ======
  on(btnConcat, "click", submitConcat);
  [["first", concatZoneFirst, concatFirstInput, concatFirstRemove], ["second", concatZoneSecond, concatSecondInput, concatSecondRemove]].forEach(([slot, zone, input, removeBtn]) => {
    if (zone && input) {
      bindFilePicker(zone, input, true);
      on(zone, "dragover", (e) => { e.preventDefault(); if (!isSubmitting && serverSupportsConcat) zone.classList.add("drag-over"); });
      on(zone, "dragleave", (e) => { e.preventDefault(); zone.classList.remove("drag-over"); });
      on(zone, "drop", (e) => {
        e.preventDefault();
        zone.classList.remove("drag-over");
        const files = e.dataTransfer && e.dataTransfer.files;
        if (!isSubmitting && serverSupportsConcat && files && files.length > 0) selectConcatFile(slot, files[0]);
      });
      on(input, "change", (e) => {
        const files = e.target && e.target.files;
        if (serverSupportsConcat && files && files.length > 0) selectConcatFile(slot, files[0]);
      });
    }
    on(removeBtn, "click", () => removeConcatFile(slot));
  });

  on(btnDownload, "click", () => {
    if (downloadUrl) { const a = document.createElement("a"); a.href = downloadUrl; a.download = downloadName; document.body.appendChild(a); a.click(); a.remove(); }
  });
  on(btnCancelRequest, "click", cancelActiveRequest);
  on(btnReset, "click", startNewTask);
  on(btnRetry, "click", returnToCurrentInputs);
  on(btnDiscardPending, "click", discardPendingResult);

  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("drop", (e) => e.preventDefault());
  document.addEventListener("paste", (e) => {
    if (!isSubmitting && uploadSection && !uploadSection.classList.contains("hidden") && e.clipboardData && e.clipboardData.files.length > 0) {
      e.preventDefault(); selectFile(e.clipboardData.files[0]);
    }
  });

  const hasPendingResult = restorePendingJobResult();
  void fetchServerCapabilities();
  setDownloadState(); syncBusyState(); updateRetryLabel(); updateFormatButton(); updateConcatButton(); toggleCoverMetaPanel(); initParticles(); initSignatureTyping();
  if (hasPendingResult) {
    showError("检测到上次尚未领取的任务结果，可点击下方按钮继续获取。", "可恢复任务");
    updateRetryLabel();
  }
})();
