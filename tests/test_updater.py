"""
自動更新モジュールおよびリソースパス解決の単体テスト
オフライン・モック環境で安全に動作検証を行います。
"""
import unittest
from unittest.mock import patch, MagicMock
import json
import urllib.error
from pathlib import Path

from core.app_updater import (
    parse_version_tuple,
    compare_versions,
    is_newer_version,
    fetch_latest_release_info
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

    def test_get_resource_path(self):
        """リソースパス解決が正しい絶対パスを返すことの検証"""
        resolved = get_resource_path("assets/test.txt")
        self.assertTrue(isinstance(resolved, Path))
        self.assertTrue(str(resolved).endswith("assets\\test.txt") or str(resolved).endswith("assets/test.txt"))


if __name__ == "__main__":
    unittest.main()
