"""プレビュー WebView への payload 注入 JS（pywebview / tkwebview2 共通）。"""
from __future__ import annotations

import json
import re
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


_JS_IDENTIFIER = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def build_set_global_js(fn_name: str, payload: Any) -> str:
    """window[fn_name](payload) を呼ぶ注入 JS（build_payload_inject_js と同じ retry パターン）。

    C-2 の `window.setProposalResult` / `window.setProposalUpdateResult` など、
    React 側が登録するコールバック関数へ結果を push するのに使う汎用版。
    fn_name はコード内で決め打ちした識別子のみを渡す想定（信頼できない文字列を
    そのまま JS へ埋め込まないよう、単純な識別子でなければ ValueError にする）。
    """
    if not _JS_IDENTIFIER.match(fn_name):
        raise ValueError(f"fn_name は単純な識別子である必要があります: {fn_name!r}")
    payload_json = json.dumps(payload, ensure_ascii=False)
    return f"""
    (function tryCall(n) {{
      if (window.{fn_name}) {{
        window.{fn_name}({payload_json});
      }} else if (n < 40) {{
        setTimeout(function() {{ tryCall(n + 1); }}, 50);
      }}
    }})(0);
    """
