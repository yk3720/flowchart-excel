"""connector_manager.connect_nodes のユニットテスト（Excel COM は MagicMock で代替）。"""
import unittest
from unittest.mock import MagicMock

from app.constants import ExcelConstants
from app.core.connector_manager import connect_nodes


def _theme():
    return {"connector": 0x00C0FF}


def _stop_event():
    ev = MagicMock()
    ev.is_set.return_value = False
    return ev


def _shape(left=0.0):
    shp = MagicMock()
    shp.Left = left
    return shp


class ConnectNodesTests(unittest.TestCase):
    def _run(self, nodes, shape_map):
        conn = MagicMock()
        sheet = MagicMock()
        sheet.Shapes.AddConnector.return_value = conn
        connect_nodes(sheet, nodes, shape_map, _theme(), _stop_event())
        return sheet, conn

    def test_straight_for_same_column_forward_adjacent_tier(self) -> None:
        nodes = [
            {"id": "10", "type": "処理", "dests_down": ["20"], "dests_right": [], "level": 0, "tier": 0, "ridx": 0},
            {"id": "20", "type": "処理", "dests_down": [], "dests_right": [], "level": 0, "tier": 1, "ridx": 1},
        ]
        shape_map = {"10": _shape(0.0), "20": _shape(0.0)}
        sheet, conn = self._run(nodes, shape_map)

        sheet.Shapes.AddConnector.assert_called_once_with(
            ExcelConstants.MSOCONNECTOR_STRAIGHT, 0, 0, 10, 10
        )
        conn.ConnectorFormat.BeginConnect.assert_called_once_with(
            shape_map["10"], ExcelConstants.CONNECTOR_SITE_BOTTOM
        )
        conn.ConnectorFormat.EndConnect.assert_called_once_with(
            shape_map["20"], ExcelConstants.CONNECTOR_SITE_TOP
        )

    def test_elbow_when_branching_to_a_lower_level(self) -> None:
        # level_diff > 0: 分岐先が右の列へ広がるケース
        nodes = [
            {"id": "10", "type": "判断", "dests_down": ["20"], "dests_right": [], "level": 0, "tier": 0, "ridx": 0},
            {"id": "20", "type": "処理", "dests_down": [], "dests_right": [], "level": 1, "tier": 1, "ridx": 1},
        ]
        shape_map = {"10": _shape(0.0), "20": _shape(200.0)}
        sheet, conn = self._run(nodes, shape_map)

        sheet.Shapes.AddConnector.assert_called_once_with(
            ExcelConstants.MSOCONNECTOR_ELBOW, 0, 0, 10, 10
        )
        conn.ConnectorFormat.BeginConnect.assert_called_once_with(
            shape_map["10"], ExcelConstants.CONNECTOR_SITE_RIGHT
        )
        conn.ConnectorFormat.EndConnect.assert_called_once_with(
            shape_map["20"], ExcelConstants.CONNECTOR_SITE_TOP
        )

    def test_elbow_when_target_level_is_lower(self) -> None:
        nodes = [
            {"id": "10", "type": "処理", "dests_down": ["20"], "dests_right": [], "level": 1, "tier": 0, "ridx": 0},
            {"id": "20", "type": "処理", "dests_down": [], "dests_right": [], "level": 0, "tier": 1, "ridx": 1},
        ]
        shape_map = {"10": _shape(200.0), "20": _shape(0.0)}
        _sheet, conn = self._run(nodes, shape_map)

        conn.ConnectorFormat.EndConnect.assert_called_once_with(
            shape_map["20"], ExcelConstants.CONNECTOR_SITE_LEFT
        )

    def test_elbow_with_default_sites_for_loop_back_reference(self) -> None:
        # target が source より前段（ループ）: elbow だが sites は既定(BOTTOM/TOP)のまま
        nodes = [
            {"id": "10", "type": "処理", "dests_down": ["5"], "dests_right": [], "level": 0, "tier": 3, "ridx": 3},
            {"id": "5", "type": "処理", "dests_down": [], "dests_right": [], "level": 0, "tier": 0, "ridx": 0},
        ]
        shape_map = {"10": _shape(0.0), "5": _shape(0.0)}
        sheet, conn = self._run(nodes, shape_map)

        sheet.Shapes.AddConnector.assert_called_once_with(
            ExcelConstants.MSOCONNECTOR_ELBOW, 0, 0, 10, 10
        )
        conn.ConnectorFormat.BeginConnect.assert_called_once_with(
            shape_map["10"], ExcelConstants.CONNECTOR_SITE_BOTTOM
        )
        conn.ConnectorFormat.EndConnect.assert_called_once_with(
            shape_map["5"], ExcelConstants.CONNECTOR_SITE_TOP
        )

    def test_right_direction_uses_elbow_and_right_site(self) -> None:
        nodes = [
            {"id": "10", "type": "判断", "dests_down": [], "dests_right": ["30"], "level": 0, "tier": 0, "ridx": 0},
            {"id": "30", "type": "処理", "dests_down": [], "dests_right": [], "level": 1, "tier": 0, "ridx": 1},
        ]
        shape_map = {"10": _shape(0.0), "30": _shape(200.0)}
        sheet, conn = self._run(nodes, shape_map)

        sheet.Shapes.AddConnector.assert_called_once_with(
            ExcelConstants.MSOCONNECTOR_ELBOW, 0, 0, 10, 10
        )
        conn.ConnectorFormat.BeginConnect.assert_called_once_with(
            shape_map["10"], ExcelConstants.CONNECTOR_SITE_RIGHT
        )
        conn.ConnectorFormat.EndConnect.assert_called_once_with(
            shape_map["30"], ExcelConstants.CONNECTOR_SITE_TOP
        )

    def test_decision_label_added_for_judgment_source(self) -> None:
        nodes = [
            {"id": "10", "type": "判断", "dests_down": ["20"], "dests_right": [], "level": 0, "tier": 0, "ridx": 0},
            {"id": "20", "type": "処理", "dests_down": [], "dests_right": [], "level": 0, "tier": 1, "ridx": 1},
        ]
        shape_map = {"10": _shape(0.0), "20": _shape(0.0)}
        sheet = MagicMock()
        conn = MagicMock()
        conn.Name = "conn_1"
        sheet.Shapes.AddConnector.return_value = conn
        label = MagicMock()
        label.Name = "label_1"
        sheet.Shapes.AddTextbox.return_value = label

        names = connect_nodes(sheet, nodes, shape_map, _theme(), _stop_event())

        self.assertIn("conn_1", names)
        self.assertIn("label_1", names)

    def test_missing_target_logs_warning_and_is_skipped(self) -> None:
        nodes = [
            {"id": "10", "type": "処理", "dests_down": ["999"], "dests_right": [], "level": 0, "tier": 0, "ridx": 0},
        ]
        shape_map = {"10": _shape(0.0)}
        sheet = MagicMock()

        with self.assertLogs("flowchart-excel", level="WARNING") as cm:
            names = connect_nodes(sheet, nodes, shape_map, _theme(), _stop_event())

        self.assertEqual(names, [])
        sheet.Shapes.AddConnector.assert_not_called()
        self.assertTrue(any("connector_target_missing" in m for m in cm.output))


if __name__ == "__main__":
    unittest.main()
