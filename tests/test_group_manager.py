"""group_manager のユニットテスト（Excel COM は MagicMock で代替）。"""
import unittest
from unittest.mock import MagicMock

import pywintypes

from app.constants import ExcelConstants
from app.core.flow_colors import DEFAULT_FILL_HEX, hex_to_vba_rgb
from app.core.group_manager import add_frame_and_title, create_final_groups


class AddFrameAndTitleTests(unittest.TestCase):
    def test_frame_is_filled_white_and_sent_to_back(self) -> None:
        sheet = MagicMock()
        frame = MagicMock()
        title = MagicMock()
        # AddShape はタイトルの後に外枠を作る（この順に呼ばれる想定）
        sheet.Shapes.AddTextbox.return_value = title
        sheet.Shapes.AddShape.return_value = frame

        add_frame_and_title(sheet, (0.0, 0.0, 100.0, 100.0), "タイトル")

        self.assertTrue(frame.Fill.Visible)
        self.assertEqual(frame.Fill.ForeColor.RGB, hex_to_vba_rgb(DEFAULT_FILL_HEX))
        frame.ZOrder.assert_called_once_with(ExcelConstants.MSO_SEND_TO_BACK)


class CreateFinalGroupsTests(unittest.TestCase):
    def test_returns_final_group_name_on_success(self) -> None:
        sheet = MagicMock()
        final_group = MagicMock()
        final_group.Name = "final_group"
        sheet.Shapes.Range.return_value.Group.return_value = final_group

        result = create_final_groups(sheet, ["shape_1", "conn_1"], [])

        self.assertEqual(result, "final_group")

    def test_final_grouping_failure_ungroups_composites_and_raises(self) -> None:
        sheet = MagicMock()
        bg, tx = MagicMock(), MagicMock()
        bg.Name, tx.Name = "bg_1", "tx_1"
        composite_group = MagicMock()
        composite_group.Name = "composite_group_1"

        range_mock = sheet.Shapes.Range
        # 1回目の Range(...).Group() は複合グループ化（成功）、2回目は最終グループ化（失敗）
        range_mock.return_value.Group.side_effect = [
            composite_group,
            pywintypes.com_error(-1, "grouping failed", None, None),
        ]

        with self.assertRaises(RuntimeError):
            create_final_groups(sheet, ["shape_1"], [(bg, tx)])

        sheet.Shapes.assert_any_call("composite_group_1")
        sheet.Shapes("composite_group_1").Ungroup.assert_called_once()


if __name__ == "__main__":
    unittest.main()
