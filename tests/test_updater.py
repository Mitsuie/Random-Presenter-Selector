"""
自動更新モジュールおよびリソースパス解決の単体テスト
オフライン・モック環境で安全に動作検証を行います。
"""
import os
import tempfile
import unittest
from unittest.mock import patch, MagicMock
import json
import urllib.error
from pathlib import Path

from core.app_updater import (
    parse_version_tuple,
    compare_versions,
    is_newer_version,
    fetch_latest_release_info,
    download_installer,
    UpdateError
)
from core.utils import get_resource_path


class TestAppUpdater(unittest.TestCase):

    def test_version_comparisons(self):
        """セマンティックバージョン比較と新旧判定の検証"""
        self.assertEqual(compare_versions("1.1.0", "1.0.1"), 1)
        self.assertEqual(compare_versions("v2.0.0", "2.0.0"), 0)
        self.assertEqual(compare_versions("1.0.0", "1.0.1"), -1)
        self.assertEqual(compare_versions("1.0", "1.0.0"), 0)
        self.assertEqual(compare_versions("2.1.0-beta", "2.0.9"), 1)

        self.assertTrue(is_newer_version("1.0.0", "1.1.0"))
        self.assertFalse(is_newer_version("1.1.0", "1.1.0"))
        self.assertFalse(is_newer_version("1.2.0", "1.1.0"))

    @patch("urllib.request.urlopen")
    def test_fetch_latest_release_mock_success(self, mock_urlopen):
        """GitHub Releases API レスポンス解析の検証（モック）"""
        mock_data = {
            "tag_name": "v1.1.0",
            "name": "v1.1.0 Release",
            "body": "## 変更履歴\n- 自動更新機能の追加\n- インストーラー対応",
            "html_url": "https://github.com/Mitsuie/Random-Presenter-Selector/releases/tag/v1.1.0",
            "published_at": "2026-09-09T00:00:00Z",
            "assets": [
                {
                    "name": "Random-Presenter-Selector_Setup_v1.1.0.exe",
                    "browser_download_url": "https://github.com/Mitsuie/Random-Presenter-Selector/releases/download/v1.1.0/Random-Presenter-Selector_Setup_v1.1.0.exe",
                    "size": 15728640
                }
            ]
        }
        mock_cm = MagicMock()
        mock_cm.status = 200
        mock_cm.read.return_value = json.dumps(mock_data).encode("utf-8")
        mock_cm.__enter__.return_value = mock_cm
        mock_urlopen.return_value = mock_cm

        info = fetch_latest_release_info("Mitsuie", "Random-Presenter-Selector", current_version="1.0.0")
        self.assertIsNotNone(info)
        self.assertTrue(info.is_update_available)
        self.assertEqual(info.version, "1.1.0")
        self.assertEqual(info.installer_name, "Random-Presenter-Selector_Setup_v1.1.0.exe")
        self.assertEqual(info.installer_size, 15728640)
        self.assertIn("自動更新機能の追加", info.release_notes)

    @patch("urllib.request.urlopen")
    def test_fetch_latest_release_offline(self, mock_urlopen):
        """オフライン時またはネットワークエラー時に例外を出さず None を返す耐障害性の検証"""
        mock_urlopen.side_effect = urllib.error.URLError("No route to host")
        info = fetch_latest_release_info("Mitsuie", "Random-Presenter-Selector", current_version="1.0.0")
        self.assertIsNone(info)

    @patch("urllib.request.urlopen")
    def test_fetch_uses_exact_installer_name(self, mock_urlopen):
        """インストーラー名のテンプレートを指定した場合、完全一致するアセットだけを採用すること"""
        mock_data = {
            "tag_name": "v1.1.0",
            "assets": [
                {"name": "other-tool.exe", "browser_download_url": "https://github.com/x/other-tool.exe", "size": 1},
                {"name": "MyApp_Setup_v1.1.0.exe", "browser_download_url": "https://github.com/x/MyApp_Setup_v1.1.0.exe", "size": 2},
            ]
        }
        mock_cm = MagicMock()
        mock_cm.status = 200
        mock_cm.read.return_value = json.dumps(mock_data).encode("utf-8")
        mock_cm.__enter__.return_value = mock_cm
        mock_urlopen.return_value = mock_cm

        info = fetch_latest_release_info("o", "r", "1.0.0", installer_name_template="MyApp_Setup_v{version}.exe")
        self.assertEqual(info.installer_name, "MyApp_Setup_v1.1.0.exe")
        self.assertEqual(info.installer_size, 2)

        mock_data["assets"] = mock_data["assets"][:1]
        mock_cm.read.return_value = json.dumps(mock_data).encode("utf-8")
        info = fetch_latest_release_info("o", "r", "1.0.0", installer_name_template="MyApp_Setup_v{version}.exe")
        self.assertIsNone(info.installer_download_url)

    def test_get_resource_path(self):
        """リソースパス解決が正しい絶対パスを返すことの検証"""
        resolved = get_resource_path("assets/test.txt")
        self.assertTrue(isinstance(resolved, Path))
        self.assertTrue(str(resolved).endswith("assets\\test.txt") or str(resolved).endswith("assets/test.txt"))


class TestDownloadInstaller(unittest.TestCase):
    """インストーラーのダウンロード検証（モック）"""

    URL = "https://github.com/o/r/releases/download/v1.1.0/MyApp_Setup_v1.1.0.exe"
    FINAL_URL = "https://objects.githubusercontent.com/github-production-release-asset/abc"

    def _mock_response(self, body: bytes, content_length=None, final_url=FINAL_URL):
        resp = MagicMock()
        resp.geturl.return_value = final_url
        resp.headers = {"Content-Length": str(len(body) if content_length is None else content_length)}
        chunks = [body, b""]
        resp.read.side_effect = lambda n: chunks.pop(0)
        resp.__enter__.return_value = resp
        return resp

    def _cleanup(self, path):
        os.remove(path)
        os.rmdir(os.path.dirname(path))

    @patch("urllib.request.urlopen")
    def test_download_success(self, mock_urlopen):
        mock_urlopen.return_value = self._mock_response(b"x" * 10)
        path = download_installer(self.URL, "MyApp_Setup_v1.1.0.exe", expected_size=10)
        try:
            self.assertEqual(os.path.basename(path), "MyApp_Setup_v1.1.0.exe")
            with open(path, "rb") as f:
                self.assertEqual(f.read(), b"x" * 10)
        finally:
            self._cleanup(path)

    def test_rejects_unexpected_host(self):
        for url in ("https://example.com/MyApp.exe", "http://github.com/o/r/MyApp.exe"):
            with self.assertRaises(UpdateError):
                download_installer(url)

    @patch("urllib.request.urlopen")
    def test_rejects_unexpected_redirect(self, mock_urlopen):
        mock_urlopen.return_value = self._mock_response(b"x", final_url="https://evil.example.com/a.exe")
        with self.assertRaises(UpdateError):
            download_installer(self.URL)

    @patch("urllib.request.urlopen")
    def test_rejects_truncated_download(self, mock_urlopen):
        """Content-Length やリリース情報のサイズと一致しない場合は破棄すること"""
        for body, content_length, expected_size in ((b"x" * 5, 10, 0), (b"x" * 5, None, 10)):
            mock_urlopen.return_value = self._mock_response(body, content_length=content_length)
            save_dir = tempfile.mkdtemp()
            with patch("tempfile.mkdtemp", return_value=save_dir):
                with self.assertRaises(UpdateError):
                    download_installer(self.URL, "a.exe", expected_size=expected_size)
            # 不完全なファイルと一時ディレクトリが残らないこと
            self.assertFalse(os.path.exists(save_dir))

    @patch("urllib.request.urlopen")
    def test_target_filename_is_sanitized(self, mock_urlopen):
        mock_urlopen.return_value = self._mock_response(b"x")
        path = download_installer(self.URL, "..\\..\\evil.exe")
        try:
            self.assertEqual(os.path.basename(path), "evil.exe")
            self.assertTrue(os.path.basename(os.path.dirname(path)).startswith("app_update_"))
        finally:
            self._cleanup(path)


if __name__ == "__main__":
    unittest.main()
