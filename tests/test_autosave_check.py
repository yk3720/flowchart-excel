"""autosave_reminder / autosave_reminder_for_watch のユニットテスト（構想設計§7）。"""
import unittest
from unittest.mock import MagicMock, PropertyMock, patch

import pywintypes

from app.core.autosave_check import (
    AUTOSAVE_REMINDER,
    autosave_reminder,
    autosave_reminder_for_watch,
)

_WATCH = {"workbookName": "Book1.xlsx", "sheetName": "Sheet1"}


class AutosaveReminderTests(unittest.TestCase):
    def test_autosave_on_returns_none(self) -> None:
        workbook = MagicMock()
        workbook.AutoSaveOn = True
        self.assertIsNone(autosave_reminder(workbook))

    def test_autosave_off_returns_reminder(self) -> None:
        workbook = MagicMock()
        workbook.AutoSaveOn = False
        self.assertEqual(autosave_reminder(workbook), AUTOSAVE_REMINDER)

    def test_property_unavailable_returns_none(self) -> None:
        workbook = MagicMock()
        type(workbook).AutoSaveOn = PropertyMock(
            side_effect=pywintypes.com_error(-1, "no prop", None, None)
        )
        self.assertIsNone(autosave_reminder(workbook))


class AutosaveReminderForWatchTests(unittest.TestCase):
    def test_excel_not_running_returns_none(self) -> None:
        with patch("app.core.autosave_check.get_excel_app", return_value=None):
            self.assertIsNone(autosave_reminder_for_watch(_WATCH))

    def test_workbook_not_found_returns_none(self) -> None:
        app = MagicMock()
        other_wb = MagicMock()
        other_wb.Name = "Other.xlsx"
        app.Workbooks = [other_wb]
        with patch("app.core.autosave_check.get_excel_app", return_value=app):
            self.assertIsNone(autosave_reminder_for_watch(_WATCH))

    def test_workbook_found_with_autosave_off_returns_reminder(self) -> None:
        app = MagicMock()
        workbook = MagicMock()
        workbook.Name = _WATCH["workbookName"]
        workbook.AutoSaveOn = False
        app.Workbooks = [workbook]
        with patch("app.core.autosave_check.get_excel_app", return_value=app):
            self.assertEqual(autosave_reminder_for_watch(_WATCH), AUTOSAVE_REMINDER)


if __name__ == "__main__":
    unittest.main()
