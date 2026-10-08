# -*- coding: utf-8 -*-
"""Flask 外壳：同一份 WebApi 以 HTTP 方式提供给浏览器。

通用分发：POST /api/<method>，body 为 JSON 参数数组；返回 {"result": ...}
或 {"error": ...}。前端通过注入的 Proxy 脚本以 bridge.api.method(...) 调用，
与 pywebview 外壳的前端代码完全一致。
"""
from __future__ import annotations

import re
import sys
import time
import traceback
from pathlib import Path

from flask import Flask, jsonify, request, Response

ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) \
    else Path(__file__).resolve().parent
# 冻结态：优先 exe 旁的 core\（发行包布局），否则用 exe 自身目录
# （模板等运行时数据落在 exe 旁，随 exe 走）
if getattr(sys, "frozen", False):
    CORE = ROOT / "core" if (ROOT / "core").is_dir() else ROOT
else:
    CORE = ROOT / "core" if (ROOT / "core" / "src").is_dir() else ROOT.parent / "core"
SRC = CORE / "src"
if not getattr(sys, "frozen", False) and str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from scan2excel.web_app import WebApi  # noqa: E402

API = WebApi(CORE)


def _index_html() -> Path:
    """开发态读 core/frontend/web/；打包态读 PyInstaller 内嵌资源。"""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", ".")) / "web" / "index.html"
    return CORE / "frontend" / "web" / "index.html"

# 让前端以 bridge.api.method(...) 调用后端（与 pywebview js_api 形态一致）
_BRIDGE_JS = """
<script>
(function(){
  var proxy = new Proxy({}, {
    get: function(_, method){
      return function(){
        var args = Array.prototype.slice.call(arguments);
        return fetch('/api/' + method, {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify(args)
        }).then(function(resp){
          return resp.json().then(function(data){
            if (resp.ok) return data.result;
            throw new Error(data.error || ('调用失败: ' + method));
          });
        });
      };
    }
  });
  window.bridge = {api: proxy};
  window.dispatchEvent(new Event('bridgeready'));
})();
</script>
"""


def create_app() -> Flask:
    app = Flask(__name__)

    @app.get("/")
    def index() -> Response:
        html = _index_html().read_text(encoding="utf-8")
        return html.replace("</head>", _BRIDGE_JS + "</head>", 1)

    @app.post("/api/<method>")
    def call(method: str):
        handler = getattr(API, method, None)
        if not callable(handler) or method.startswith("_"):
            return jsonify({"error": f"未知接口：{method}"}), 404
        args = request.get_json(silent=True) or []
        try:
            return jsonify({"result": handler(*args)})
        except (ValueError, RuntimeError) as exc:
            return jsonify({"error": str(exc)}), 400
        except Exception as exc:  # noqa: BLE001
            traceback.print_exc()
            return jsonify({"error": f"{type(exc).__name__}: {exc}"}), 500

    @app.post("/api/upload_images")
    def upload_images():
        """浏览器外壳的拖拽导入：保存上传文件后按路径入列。

        浏览器出于安全不暴露文件真实路径，拖入的 File 对象以
        multipart 形式上传，落盘后走同一套 add_image_paths 逻辑。
        """
        files = request.files.getlist("files")
        if not files:
            return jsonify({"error": "没有收到文件"}), 400
        upload_dir = CORE / "data" / "uploads"
        upload_dir.mkdir(parents=True, exist_ok=True)
        paths = []
        for f in files:
            safe = re.sub(r"[^\w.\-\u4e00-\u9fff]+", "_", f.filename or "file")
            out = upload_dir / f"{int(time.time() * 1000)}_{safe}"
            f.save(out)
            paths.append(str(out))
        try:
            added = API.add_image_paths(paths)
            return jsonify({"result": added})
        except (ValueError, RuntimeError) as exc:
            return jsonify({"error": str(exc)}), 400

    return app


app = create_app()

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8750, debug=False)
