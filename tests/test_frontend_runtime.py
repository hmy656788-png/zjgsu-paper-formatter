"""Exercise frontend request state with Node's runtime and a small DOM fixture."""

import shutil
import subprocess
import unittest
from pathlib import Path


NODE = shutil.which("node")
SCRIPT_PATH = Path(__file__).resolve().parents[1] / "static" / "script.js"


FRONTEND_FIXTURE = r"""
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(process.argv[1], 'utf8');
const storageKey = 'zjgsu-paper-formatter.pending-result.v1';
const jobId = 'deadbeef'.repeat(4);
const resultUrl = `/api/jobs/${jobId}/result`;

function deferred() {
  let resolve;
  const promise = new Promise((accept) => { resolve = accept; });
  return { promise, resolve };
}

class Element {
  constructor() {
    const classes = new Set();
    this.classList = {
      add: (...values) => values.forEach((value) => classes.add(value)),
      remove: (...values) => values.forEach((value) => classes.delete(value)),
      contains: (value) => classes.has(value),
      toggle: (value, enabled) => {
        const next = enabled === undefined ? !classes.has(value) : enabled;
        if (next) classes.add(value); else classes.delete(value);
        return next;
      },
    };
    this.listeners = new Map();
    this.attributes = new Map();
    this.children = new Map();
    this.disabled = false;
    this.checked = false;
    this.textContent = '';
    this.value = '';
    this.dataset = {};
  }
  addEventListener(name, callback) {
    if (!this.listeners.has(name)) this.listeners.set(name, []);
    this.listeners.get(name).push(callback);
  }
  dispatch(name, event = {}) {
    for (const callback of this.listeners.get(name) || []) callback(event);
  }
  click() { if (!this.disabled) this.dispatch('click'); }
  setAttribute(name, value) { this.attributes.set(name, value); }
  removeAttribute(name) { this.attributes.delete(name); }
  querySelector(selector) {
    if (!this.children.has(selector)) this.children.set(selector, new Element());
    return this.children.get(selector);
  }
  appendChild() {}
  focus() {}
  scrollIntoView() {}
}

function response(payload, status = 200, extraHeaders = {}) {
  return {
    status, ok: status >= 200 && status < 300,
    headers: { get: (name) => name.toLowerCase() === 'content-type' ? 'application/json' : (extraHeaders[name.toLowerCase()] || '') },
    json: async () => payload,
  };
}

function createFrontend({ pending = false, provisional = false, createdAt = Date.now(), eventSourceClass, resultResponses = [] } = {}) {
  const elements = {};
  for (const match of source.matchAll(/\$\("#([^\"]+)"\)/g)) {
    elements[match[1]] = new Element();
  }
  for (const id of [
    'pasteSection', 'concatSection', 'filePreview', 'processingSection',
    'resultSection', 'errorSection',
  ]) elements[id].classList.add('hidden');
  elements.tabUpload.classList.add('active');
  const storage = new Map();
  if (pending) storage.set(storageKey, JSON.stringify({
    result_url: resultUrl, mode: 'concat', created_at: createdAt, provisional,
  }));
  const health = deferred();
  const result = deferred();
  const creation = deferred();
  const requests = [];
  const timers = new Map();
  const timerDelays = new Map();
  let nextTimer = 0;
  const setTimer = (callback, delay) => {
    timers.set(++nextTimer, callback);
    timerDelays.set(nextTimer, delay);
    return nextTimer;
  };
  const clearTimer = (id) => { timers.delete(id); timerDelays.delete(id); };
  const context = vm.createContext({
    document: {
      querySelector: (selector) => elements[selector.slice(1)] || null,
      querySelectorAll: () => [],
      addEventListener: () => {},
      createElement: () => new Element(),
      body: new Element(),
    },
    window: {
      matchMedia: () => ({ matches: true }),
      setTimeout: setTimer, clearTimeout: clearTimer,
      sessionStorage: {
        getItem: (key) => storage.get(key) || null,
        setItem: (key, value) => storage.set(key, value),
        removeItem: (key) => storage.delete(key),
      },
    },
    setTimeout: setTimer, clearTimeout: clearTimer,
    setInterval: setTimer, clearInterval: clearTimer,
    AbortController, DOMException, Blob,
    ...(eventSourceClass ? { EventSource: eventSourceClass } : {}),
    FormData: class { append() {} },
    fetch: (url, options) => {
      requests.push({ url, options });
      if (url === '/api/health') return health.promise;
      if (url === resultUrl) return resultResponses.length ? Promise.resolve(resultResponses.shift()) : result.promise;
      if (url === '/api/concat_async') return creation.promise;
      throw new Error(`Unexpected request: ${url}`);
    },
  });
  vm.runInContext(source, context, { filename: 'static/script.js' });
  return { elements, storage, health, result, creation, requests, timers, timerDelays };
}

const settle = () => new Promise((resolve) => setImmediate(resolve));
const disableConcat = async (frontend) => {
  frontend.health.resolve(response({ success: true, features: { concat: false } }));
  await settle();
};
const visible = (element) => !element.classList.contains('hidden');
"""


@unittest.skipUnless(NODE, "Node.js is required for frontend runtime tests")
class FrontendRuntimeTests(unittest.TestCase):
    def test_complete_event_retries_processing_result_using_retry_after(self):
        self.run_frontend(r"""
(async () => {
  const sources = [];
  class EventSourceFixture extends Element {
    constructor() { super(); sources.push(this); }
    close() { this.closed = true; }
  }
  const frontend = createFrontend({
    eventSourceClass: EventSourceFixture,
    resultResponses: [
      response({ success: false, status: 'processing' }, 202, { 'retry-after': '0.1' }),
      response({ success: true, download_url: `/api/download/${jobId}_output.docx` }),
    ],
  });
  const el = frontend.elements;
  el.tabConcat.click();
  for (const id of ['concatFirstInput', 'concatSecondInput']) {
    el[id].dispatch('change', { target: { files: [{ name: 'paper.docx', size: 200 }] } });
  }
  el.btnConcat.click();
  frontend.creation.resolve(response({
    success: true, result_url: resultUrl, events_url: `/api/jobs/${jobId}/events`,
  }, 202));
  await settle();
  sources[0].dispatch('complete');
  await settle();
  const retryTimer = [...frontend.timerDelays].find(([, delay]) => delay === 100);
  assert.ok(retryTimer, 'Retry-After should drive a bounded retry delay');
  frontend.timers.get(retryTimer[0])();
  await settle();
  assert.equal(visible(el.resultSection), true);
  assert.equal(el.btnDownload.disabled, false);
})();
""")

    def test_complete_event_honors_retry_after_for_transient_result_error(self):
        self.run_frontend(r"""
(async () => {
  const sources = [];
  class EventSourceFixture extends Element {
    constructor() { super(); sources.push(this); }
    close() { this.closed = true; }
  }
  const frontend = createFrontend({
    eventSourceClass: EventSourceFixture,
    resultResponses: [
      response({ success: false, error: '服务排空中' }, 503, { 'retry-after': '0.15' }),
      response({ success: true, download_url: `/api/download/${jobId}_output.docx` }),
    ],
  });
  const el = frontend.elements;
  el.tabConcat.click();
  for (const id of ['concatFirstInput', 'concatSecondInput']) {
    el[id].dispatch('change', { target: { files: [{ name: 'paper.docx', size: 200 }] } });
  }
  el.btnConcat.click();
  frontend.creation.resolve(response({
    success: true, result_url: resultUrl, events_url: `/api/jobs/${jobId}/events`,
  }, 202));
  await settle();
  sources[0].dispatch('complete');
  await settle();
  const retryTimer = [...frontend.timerDelays].find(([, delay]) => delay === 150);
  assert.ok(retryTimer, 'SSE result fetch should honor 503 Retry-After');
  frontend.timers.get(retryTimer[0])();
  await settle();
  assert.equal(visible(el.resultSection), true);
  assert.equal(el.btnDownload.disabled, false);
})();
""")

    def test_polling_fallback_honors_retry_after_header(self):
        self.run_frontend(r"""
(async () => {
  const frontend = createFrontend({
    resultResponses: [
      response({ success: false, status: 'processing' }, 202, { 'retry-after': '0.2' }),
      response({ success: true, download_url: `/api/download/${jobId}_output.docx` }),
    ],
  });
  const el = frontend.elements;
  el.tabConcat.click();
  for (const id of ['concatFirstInput', 'concatSecondInput']) {
    el[id].dispatch('change', { target: { files: [{ name: 'paper.docx', size: 200 }] } });
  }
  el.btnConcat.click();
  frontend.creation.resolve(response({
    success: true, result_url: resultUrl, events_url: `/api/jobs/${jobId}/events`,
  }, 202));
  await settle();
  const retryTimer = [...frontend.timerDelays].find(([, delay]) => delay === 200);
  assert.ok(retryTimer, 'polling fallback should honor Retry-After');
  frontend.timers.get(retryTimer[0])();
  await settle();
  assert.equal(visible(el.resultSection), true);
  assert.equal(el.btnDownload.disabled, false);
})();
""")

    def test_failed_stream_acknowledgement_has_a_bounded_request(self):
        self.run_frontend(r"""
(async () => {
  const sources = [];
  class EventSourceFixture extends Element {
    constructor() { super(); sources.push(this); }
    close() { this.closed = true; }
  }
  const frontend = createFrontend({ eventSourceClass: EventSourceFixture });
  const el = frontend.elements;
  el.tabConcat.click();
  for (const id of ['concatFirstInput', 'concatSecondInput']) {
    el[id].dispatch('change', { target: { files: [{ name: 'paper.docx', size: 200 }] } });
  }
  el.btnConcat.click();
  frontend.creation.resolve(response({
    success: true, result_url: resultUrl, events_url: `/api/jobs/${jobId}/events`,
  }, 202));
  await settle();
  sources[0].dispatch('failed', { data: JSON.stringify({ message: '文档处理失败' }) });
  const acknowledgement = frontend.requests.find((request) => request.url === resultUrl);
  assert.ok(acknowledgement.options.signal, 'acknowledgement must have an abort signal');
  assert.equal(acknowledgement.options.signal.aborted, false);
  assert.equal(frontend.storage.has(storageKey), false);
  assert.equal(visible(el.errorSection), true);
  assert.equal(el.errorMessage.textContent, '文档处理失败');
  const timeout = [...frontend.timerDelays].find(([, delay]) => delay === 30000);
  assert.ok(timeout, 'acknowledgement must have a result request timeout');
  frontend.timers.get(timeout[0])();
  assert.equal(acknowledgement.options.signal.aborted, true);
  // Even a late response cannot replace the original failure message.
  frontend.result.resolve(response({ success: false, error: '服务器错误' }, 500));
  await settle();
  assert.equal(el.errorMessage.textContent, '文档处理失败');
  assert.equal(frontend.timers.has(timeout[0]), false);
})();
""")

    def test_file_picker_stops_bubbling_hidden_input_click(self):
        source = SCRIPT_PATH.read_text()
        self.assertIn('on(input, "click", (event) => event.stopPropagation());', source)

    def run_frontend(self, test_script):
        result = subprocess.run(
            [NODE, "-e", FRONTEND_FIXTURE + "\n" + test_script, str(SCRIPT_PATH)],
            text=True,
            capture_output=True,
            timeout=10,
            check=False,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_rate_limit_does_not_confirm_provisional_job(self):
        self.run_frontend(r"""
(async () => {
  for (const status of [400, 401, 403, 408, 429]) {
    const frontend = createFrontend({ pending: true, provisional: true });
    frontend.elements.btnRetry.click();
    frontend.result.resolve(response({ success: false, error: 'Request rejected' }, status));
    await settle();
    assert.equal(visible(frontend.elements.errorSection), true);
    // A proxy/rate-limit response says nothing about whether upload processing
    // has created the job yet; its subsequent 404 still needs creation grace.
            assert.equal(JSON.parse(frontend.storage.get(storageKey)).provisional, true);
  }
})();
""")

    def test_result_retry_honors_retry_after_for_rate_limit(self):
        self.run_frontend(r"""
(async () => {
  const frontend = createFrontend({
    pending: true,
    resultResponses: [
      response({ success: false, error: '稍后再试' }, 429, { 'retry-after': '0.1' }),
      response({ success: true, download_url: `/api/download/${jobId}_output.docx` }),
    ],
  });
  const el = frontend.elements;
  el.btnRetry.click();
  await settle();
  const retryTimer = [...frontend.timerDelays].find(([, delay]) => delay === 100);
  assert.ok(retryTimer, '429 Retry-After should schedule a bounded retry');
  frontend.timers.get(retryTimer[0])();
  await settle();
    assert.equal(visible(el.resultSection), true);
  assert.equal(el.btnDownload.disabled, false);
})();
""")

    def test_result_retry_ignores_malformed_retry_after(self):
        self.run_frontend(r"""
(async () => {
  const frontend = createFrontend({
    pending: true,
    resultResponses: [
      response({ success: false, error: '稍后再试' }, 429, { 'retry-after': '5oops' }),
      response({ success: true, download_url: `/api/download/${jobId}_output.docx` }),
    ],
  });
  frontend.elements.btnRetry.click();
  await settle();
  const retryTimer = [...frontend.timerDelays].find(([, delay]) => delay === 4000);
  assert.ok(retryTimer, 'malformed Retry-After should use the bounded fallback delay');
  assert.equal([...frontend.timerDelays].some(([, delay]) => delay === 5000), false);
})();
""")

    def test_polling_fallback_honors_retry_after_for_service_unavailable(self):
        self.run_frontend(r"""
(async () => {
  const frontend = createFrontend({
    resultResponses: [
      response({ success: false, error: '服务排空中' }, 503, { 'retry-after': '0.15' }),
      response({ success: true, download_url: `/api/download/${jobId}_output.docx` }),
    ],
  });
  const el = frontend.elements;
  el.tabConcat.click();
  for (const id of ['concatFirstInput', 'concatSecondInput']) {
    el[id].dispatch('change', { target: { files: [{ name: 'paper.docx', size: 200 }] } });
  }
  el.btnConcat.click();
  frontend.creation.resolve(response({
    success: true, result_url: resultUrl, events_url: `/api/jobs/${jobId}/events`,
  }, 202));
  await settle();
  const retryTimer = [...frontend.timerDelays].find(([, delay]) => delay === 150);
  assert.ok(retryTimer, '503 Retry-After should drive a bounded polling delay');
  frontend.timers.get(retryTimer[0])();
  await settle();
  assert.equal(visible(el.resultSection), true);
  assert.equal(el.btnDownload.disabled, false);
})();
""")

    def test_stale_provisional_job_is_discarded_after_reload(self):
        self.run_frontend(r"""
(() => {
  const frontend = createFrontend({
    pending: true,
    provisional: true,
    createdAt: Date.now() - (12 * 60 * 1000 + 1),
  });
  // A fabricated client id cannot be reconciled after its creation grace
  // window; a stale browser session must not block a fresh upload forever.
  assert.equal(frontend.storage.get(storageKey), undefined);
  assert.equal(frontend.elements.btnDiscardPending.classList.contains('hidden'), true);
})();
""")

    def test_real_job_responses_still_confirm_provisional_job(self):
        self.run_frontend(r"""
(async () => {
  for (const [status, payload] of [
    [202, { success: false, status: 'processing' }],
    [200, { success: true, download_url: `/api/download/${jobId}_output.docx` }],
  ]) {
    const frontend = createFrontend({ pending: true, provisional: true });
    frontend.elements.btnRetry.click();
    frontend.result.resolve(response(payload, status));
    await settle();
    assert.equal(JSON.parse(frontend.storage.get(storageKey)).provisional, false);
    frontend.elements.btnCancelRequest.click();
  }
})();
""")

    def test_invalid_result_download_preserves_recovery_instead_of_false_success(self):
        self.run_frontend(r"""
(async () => {
  for (const downloadUrl of [
    undefined, '', 'https://example.com/output.docx',
    `https://formatter.example/api/download/${jobId}_output.docx`,
    `/api/download/../${jobId}_output.docx`, '/api/download/output.docx',
  ]) {
    const frontend = createFrontend({ pending: true });
    const el = frontend.elements;
    el.btnRetry.click();
    const payload = { success: true, download_url: downloadUrl };
    frontend.result.resolve(response(payload));
    await settle();
    assert.equal(visible(el.resultSection), false);
    assert.equal(visible(el.errorSection), true);
    assert.equal(el.errorTitle.textContent, '结果暂不可下载');
    assert.match(el.errorMessage.textContent, /下载地址/);
    assert.equal(el.btnDownload.disabled, true);
    assert.equal(el.retryLabel.textContent, '重新获取任务结果');
    assert.equal(JSON.parse(frontend.storage.get(storageKey)).result_url, resultUrl);

    // The same job remains recoverable without uploading the document again.
    payload.download_url = `/api/download/${jobId}_output.docx`;
    el.btnRetry.click();
    await settle();
    assert.equal(visible(el.resultSection), true);
    assert.equal(el.btnDownload.disabled, false);
    assert.equal(frontend.requests.filter((request) => request.url === resultUrl).length, 2);
  }
})();
""")

    def test_json_result_without_content_type_remains_compatible(self):
        self.run_frontend(r"""
(async () => {
  const frontend = createFrontend({ pending: true });
  frontend.elements.btnRetry.click();
  // Some reverse proxies strip Content-Type while preserving the JSON body.
  frontend.result.resolve({
    status: 200,
    ok: true,
    headers: { get: () => '' },
    json: async () => ({
      success: true,
      download_url: `/api/download/${jobId}_output.docx`,
    }),
  });
  await settle();
  assert.equal(visible(frontend.elements.resultSection), true);
  assert.equal(frontend.elements.btnDownload.disabled, false);
})();
""")

    def test_late_health_response_preserves_recoverable_concat_task(self):
        self.run_frontend(r"""
(async () => {
  const frontend = createFrontend({ pending: true });
  const el = frontend.elements;
  assert.equal(el.errorTitle.textContent, '可恢复任务');
  await disableConcat(frontend);
  assert.equal(visible(el.errorSection), true);
  assert.equal(visible(el.uploadSection), false);
  assert.equal(el.tabConcat.classList.contains('active'), true);
  assert.equal(el.tabConcat.disabled, true);
  assert.equal(el.retryLabel.textContent, '重新获取任务结果');
  assert.equal(JSON.parse(frontend.storage.get(storageKey)).mode, 'concat');
})();
""")

    def test_late_health_response_preserves_concat_recovery_progress_and_result(self):
        self.run_frontend(r"""
(async () => {
  const frontend = createFrontend({ pending: true });
  const el = frontend.elements;
  el.btnRetry.click();
  assert.equal(visible(el.processingSection), true);
  await disableConcat(frontend);
  assert.equal(visible(el.processingSection), true);
  assert.equal(el.btnCancelRequest.disabled, false);
  frontend.result.resolve(response({
    success: true, download_url: `/api/download/${jobId}_output.docx`,
  }));
  await settle();
  assert.equal(visible(el.resultSection), true);
  assert.equal(el.resultTitle.textContent, '拼接完成！');
  assert.equal(el.btnDownloadLabel.textContent, '下载拼接文档');
  assert.equal(el.outlinePanel.classList.contains('hidden'), true);
  assert.equal(el.btnDownload.disabled, false);
  // A deliberate new-task action still leaves the disabled concat mode.
  el.btnReset.click();
  assert.equal(visible(el.uploadSection), true);
  assert.equal(el.tabUpload.classList.contains('active'), true);
  assert.equal(frontend.storage.has(storageKey), false);
})();
""")

    def test_late_health_response_preserves_concat_creation_without_client_id(self):
        self.run_frontend(r"""
(async () => {
  const frontend = createFrontend();
  const el = frontend.elements;
  el.tabConcat.click();
  for (const id of ['concatFirstInput', 'concatSecondInput']) {
    el[id].dispatch('change', { target: { files: [{ name: 'paper.docx', size: 200 }] } });
  }
  el.btnConcat.click();
  assert.equal(frontend.storage.has(storageKey), false);
  assert.equal(visible(el.processingSection), true);
  await disableConcat(frontend);
  assert.equal(visible(el.processingSection), true);
  assert.equal(el.tabConcat.classList.contains('active'), true);
  assert.equal(el.btnCancelRequest.disabled, false);
  assert.equal(el.concatFirstName.textContent, 'paper.docx');
  el.btnCancelRequest.click();
  assert.equal(el.errorTitle.textContent, '已停止等待');
})();
""")

    def test_health_response_still_disables_new_concat_tasks(self):
        self.run_frontend(r"""
(async () => {
  const frontend = createFrontend();
  const el = frontend.elements;
  el.tabConcat.click();
  assert.equal(visible(el.concatSection), true);
  await disableConcat(frontend);
  assert.equal(visible(el.uploadSection), true);
  assert.equal(el.tabUpload.classList.contains('active'), true);
  assert.equal(el.tabConcat.disabled, true);
  assert.equal(el.btnConcat.disabled, true);
})();
""")
