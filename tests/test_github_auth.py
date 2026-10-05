import os
import unittest

from app import _github_api_request_headers


GITHUB_RAW_ACCEPT = "application/vnd.github.raw+json"


class GithubApiRequestHeadersTestCase(unittest.TestCase):
    def setUp(self):
        self.original_environment = os.environ.copy()

    def tearDown(self):
        os.environ.clear()
        os.environ.update(self.original_environment)

    def test_prefers_gh_token(self):
        os.environ["GH_TOKEN"] = "gh-token"
        os.environ["GITHUB_TOKEN"] = "github-token"

        headers = _github_api_request_headers(GITHUB_RAW_ACCEPT)

        self.assertEqual(headers["Authorization"], "Bearer gh-token")


    def test_falls_back_from_empty_gh_token(self):
        os.environ["GH_TOKEN"] = ""
        os.environ["GITHUB_TOKEN"] = "github-token"

        headers = _github_api_request_headers(GITHUB_RAW_ACCEPT)

        self.assertEqual(headers["Authorization"], "Bearer github-token")


    def test_rejects_unsafe_gh_token_and_falls_back(self):
        os.environ["GH_TOKEN"] = "unsafe\r\ntoken"
        os.environ["GITHUB_TOKEN"] = "github-token"

        headers = _github_api_request_headers(GITHUB_RAW_ACCEPT)

        self.assertEqual(headers["Authorization"], "Bearer github-token")
