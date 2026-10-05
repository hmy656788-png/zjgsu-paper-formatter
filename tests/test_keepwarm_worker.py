"""Runtime contract checks for the Cloudflare keepwarm Worker module."""

import shutil
import subprocess
import unittest
from pathlib import Path


NODE = shutil.which("node")
WORKER_PATH = Path(__file__).resolve().parents[1] / "keepwarm-worker" / "index.js"


WORKER_FIXTURE = r"""
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
let source = fs.readFileSync(process.argv[1], 'utf8');
// Load the module in Node without requiring a package-level ESM configuration.
source = source.replace(/export default\s*/, 'const worker = ');

class ResponseStub {
  constructor(body, options = {}) {
    this.body = body;
    this.status = options.status || 200;
    this.headers = options.headers || {};
  }
}

const requests = [];
const context = vm.createContext({
  URL,
  Response: ResponseStub,
  AbortSignal: { timeout: (ms) => ({ timeoutMs: ms }) },
  console: { warn: () => {} },
  fetch: async (target, options) => {
    requests.push({ target, options });
    return {
      ok: true,
      status: 200,
      headers: { get: (name) => name === 'content-type' ? 'application/json' : null },
      json: async () => ({ status: 'ok', storage: { storage_ready: true } }),
    };
  },
  Date,
});
vm.runInContext(source + '\nthis.worker = worker;', context, { filename: 'keepwarm-worker/index.js' });

(async () => {
  const env = {
    KEEPWARM_TARGET_URL: 'https://custom.example/health',
    KEEPWARM_REQUEST_TIMEOUT_MS: '7000',
  };
  const response = await context.worker.fetch({}, env);
  assert.equal(response.status, 200);
  assert.equal(requests.length, 1);
  assert.equal(requests[0].target, env.KEEPWARM_TARGET_URL);
  assert.equal(requests[0].options.signal.timeoutMs, 7000);
  assert.match(response.body, /target_source: env/);
})();
"""

# A successful HTTP response with an incomplete health payload must still be
# surfaced as degraded so the cron/monitor can alert instead of masking an
# unavailable storage budget.
WORKER_DEGRADED_FIXTURE = (
    WORKER_FIXTURE
    .replace("storage: { storage_ready: true }", "storage: { storage_ready: null }")
    .replace("assert.equal(response.status, 200);", "assert.equal(response.status, 502);")
)


@unittest.skipUnless(NODE, "Node.js is required for keepwarm Worker runtime tests")
class KeepwarmWorkerRuntimeTests(unittest.TestCase):
    def test_manual_fetch_uses_module_handler_env(self):
        result = subprocess.run(
            [NODE, "-e", WORKER_FIXTURE, str(WORKER_PATH)],
            text=True,
            capture_output=True,
            timeout=10,
            check=False,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_incomplete_health_payload_is_degraded(self):
        result = subprocess.run(
            [NODE, "-e", WORKER_DEGRADED_FIXTURE, str(WORKER_PATH)],
            text=True,
            capture_output=True,
            timeout=10,
            check=False,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_target_with_credentials_or_missing_host_falls_back_to_default(self):
        for invalid_target in (
            "https://user:password@custom.example/health",
            "https://:443/api/health",
        ):
            with self.subTest(invalid_target=invalid_target):
                fixture = WORKER_FIXTURE.replace(
                    "KEEPWARM_TARGET_URL: 'https://custom.example/health',",
                    f"KEEPWARM_TARGET_URL: '{invalid_target}',",
                ).replace(
                    "assert.equal(requests[0].target, env.KEEPWARM_TARGET_URL);",
                    "assert.equal(requests[0].target, 'https://zjgsu-paper-formatter.onrender.com/api/health');",
                ).replace("assert.match(response.body, /target_source: env/);", "assert.match(response.body, /target_source: default/);")
                result = subprocess.run(
                    [NODE, "-e", fixture, str(WORKER_PATH)],
                    text=True,
                    capture_output=True,
                    timeout=10,
                    check=False,
                )
                self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_manual_report_redacts_target_query_and_fragment(self):
        fixture = WORKER_FIXTURE.replace(
            "KEEPWARM_TARGET_URL: 'https://custom.example/health',",
            "KEEPWARM_TARGET_URL: 'https://custom.example/health?token=secret-value#fragment',",
        ).replace(
            "assert.match(response.body, /target_source: env/);",
            "assert.match(response.body, /target_source: env/);\n"
            "  assert.match(response.body, /https:\\/\\/custom.example\\/health/);\n"
            "  assert.doesNotMatch(response.body, /secret-value|fragment/);",
        )
        result = subprocess.run(
            [NODE, "-e", fixture, str(WORKER_PATH)],
            text=True,
            capture_output=True,
            timeout=10,
            check=False,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_fetch_error_redacts_url_credentials_and_query(self):
        fixture = WORKER_FIXTURE.replace(
            "requests.push({ target, options });\n    return {",
            "requests.push({ target, options });\n    throw new Error('fetch failed for https://custom.example/health?token=secret-value#fragment');\n    return {",
        ).replace(
            "assert.equal(response.status, 200);",
            "assert.equal(response.status, 502);\n"
            "  assert.doesNotMatch(response.body, /secret-value|fragment/);\n"
            "  assert.match(response.body, /custom.example\\/health/);",
        ).replace(
            "assert.match(response.body, /target_source: env/);",
            "assert.match(response.body, /custom.example\\/health/);",
        )
        result = subprocess.run(
            [NODE, "-e", fixture, str(WORKER_PATH)],
            text=True,
            capture_output=True,
            timeout=10,
            check=False,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
