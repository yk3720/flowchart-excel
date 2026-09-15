"""smart_palette のユニットテスト（Excel COM は MagicMock で代替）。"""
import unittest
from unittest.mock import MagicMock

from app.constants import ExcelConstants
from app.core.smart_palette import create_smart_shape


def _theme():
    return {"shape_line": 0x000000, "connector": 0x00C0FF}


class CreateSmartShapeTests(unittest.TestCase):
    def _make_sheet(self):
        sheet = MagicMock()
        shp = MagicMock(name="shp")
        shp.Name = "shape_1"
        sheet.Shapes.AddShape.return_value = shp

        connectors = []

        def _add_connector(*args, **kwargs):
            conn = MagicMock()
            conn.Name = f"conn_{len(connectors)}"
            connectors.append(conn)
            return conn

        sheet.Shapes.AddConnector.side_effect = _add_connector

        group = MagicMock()
        group.Name = "group_1"
        sheet.Shapes.Range.return_value.Group.return_value = group
        return sheet, shp, connectors, group

    def test_process_shape_creates_no_anchor_and_groups_with_connector(self) -> None:
        sheet, shp, connectors, group = self._make_sheet()

        result = create_smart_shape(sheet, "処理", 10.0, 20.0, 160.0, 60.0, _theme())

        # 主図形1個のみ（アンカー図形は作らない）
        self.assertEqual(sheet.Shapes.AddShape.call_count, 1)
        # 下側コネクタのみ生成（右側・判断ラベルは無し）
        self.assertEqual(sheet.Shapes.AddConnector.call_count, 1)
        conn = connectors[0]
        conn.ConnectorFormat.BeginConnect.assert_called_once_with(
            shp, ExcelConstants.CONNECTOR_SITE_BOTTOM
        )
        # 終点にはConnectしない（アンカー不要の設計）
        conn.ConnectorFormat.EndConnect.assert_not_called()
        # 図形＋コネクタがグループ化される
        sheet.Shapes.Range.assert_called_once_with((shp.Name, conn.Name))
        self.assertEqual(result, group.Name)

    def test_decision_shape_creates_both_connectors_and_labels_grouped(self) -> None:
        sheet, shp, connectors, group = self._make_sheet()
        labels = iter(["lbl_yes", "lbl_no"])
        with unittest.mock.patch(
            "app.core.smart_palette.add_decision_label",
            side_effect=lambda *_a, **_k: next(labels),
        ):
            result = create_smart_shape(sheet, "判断", 0.0, 0.0, 160.0, 60.0, _theme())

        self.assertEqual(sheet.Shapes.AddShape.call_count, 1)
        self.assertEqual(sheet.Shapes.AddConnector.call_count, 2)
        for conn in connectors:
            conn.ConnectorFormat.EndConnect.assert_not_called()
        expected_members = (shp.Name, connectors[0].Name, "lbl_yes", connectors[1].Name, "lbl_no")
        sheet.Shapes.Range.assert_called_once_with(expected_members)
        self.assertEqual(result, group.Name)

    def test_grouping_failure_returns_none(self) -> None:
        sheet, _shp, _connectors, _group = self._make_sheet()
        sheet.Shapes.Range.return_value.Group.side_effect = AttributeError("boom")

        result = create_smart_shape(sheet, "処理", 0.0, 0.0, 160.0, 60.0, _theme())

        self.assertIsNone(result)


if __name__ == "__main__":
    unittest.main()
