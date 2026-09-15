"""path_utils.ensure_environment の起動失敗通知のユニットテスト。"""
import unittest
from unittest.mock import MagicMock, patch

from app.utils import path_utils


class EnsureEnvironmentTests(unittest.TestCase):
    def test_permission_error_shows_native_messagebox_and_exits(self) -> None:
        fake_dir = MagicMock()
        fake_dir.exists.return_value = False
        fake_dir.mkdir.side_effect = PermissionError("denied")
        fake_root = MagicMock()
        fake_root.__truediv__.return_value = fake_dir

        with patch("app.utils.path_utils.get_app_root", return_value=fake_root), \
             patch("ctypes.windll.user32.MessageBoxW", create=True) as mock_msgbox:
            with self.assertRaises(SystemExit):
                path_utils.ensure_environment()

        mock_msgbox.assert_called_once()
        args = mock_msgbox.call_args[0]
        self.assertIn("logs", args[1])


if __name__ == "__main__":
    unittest.main()
