"""group_manager のユニットテスト（Excel COM は MagicMock で代替）。"""
import unittest
from unittest.mock import MagicMock

import pywintypes

from app.core.group_manager import create_final_groups


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
