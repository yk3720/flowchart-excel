"""プレビュー WebView への payload 注入 JS（pywebview / tkwebview2 共通）。"""
from __future__ import annotations

import json
from typing import Any, Dict


def build_payload_inject_js(payload: Dict[str, Any]) -> str:
    payload_json = json.dumps(payload, ensure_ascii=False)
    return f"""
    window.__PREVIEW_PAYLOAD__ = {payload_json};
    (function tryInject(n) {{
      if (window.setPreviewPayload) {{
        window.setPreviewPayload(window.__PREVIEW_PAYLOAD__);
      }} else if (n < 40) {{
        setTimeout(function() {{ tryInject(n + 1); }}, 50);
      }}
    }})(0);
    """
