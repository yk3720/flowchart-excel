"""flowchart-studio 同等プレビュー（WebView2 · 別プロセス）。"""
from __future__ import annotations

import json
import logging
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, Optional, Union

PreviewResult = Optional[Union[bool, Dict[str, Any]]]

logger = logging.getLogger("flowchart-excel")


def resolve_preview_dist() -> Optional[Path]:
    """preview-web/dist を探す（開発・exe 両対応）。"""
    if getattr(sys, "frozen", False):
        candidates = [
            Path(sys._MEIPASS) / "preview-web" / "dist",  # type: ignore[attr-defined]
            Path(sys.executable).resolve().parent / "preview-web" / "dist",
        ]
    else:
        root = Path(__file__).resolve().parents[2]
        candidates = [root / "preview-web" / "dist"]

    for path in candidates:
        if (path / "index.html").is_file():
            return path
    return None


def run_studio_preview(payload: Dict[str, Any]) -> PreviewResult:
    """studio 品質プレビューを開く。

    確定時は result dict（payload/fingerprint 含む）。
    キャンセル時 False。dist 無し時 None。
    """
    dist = resolve_preview_dist()
    if dist is None:
        logger.warning("studio_preview_dist_missing")
        return None

    with tempfile.TemporaryDirectory(prefix="fc-excel-preview-") as tmp:
        tmp_path = Path(tmp)
        payload_path = tmp_path / "payload.json"
        result_path = tmp_path / "result.json"
        payload_path.write_text(
            json.dumps(payload, ensure_ascii=False),
            encoding="utf-8",
        )

        cmd = [
            sys.executable,
            "--flowchart-preview",
            str(payload_path),
            str(result_path),
            str(dist),
        ]
        logger.info("studio_preview_start | dist=%s", dist)
        completed = subprocess.run(cmd, check=False)
        if completed.returncode not in (0, 1):
            logger.error(
                "studio_preview_process_failed | code=%s", completed.returncode
            )
            return False

        if not result_path.is_file():
            logger.info("studio_preview_no_result")
            return False

        try:
            result = json.loads(result_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            logger.exception("studio_preview_result_invalid")
            return False

        if result.get("action") == "confirm":
            return result
        return False
