# -*- coding: utf-8 -*-
"""Linux-specific path and browser-picker behavior, exercised with mocks."""
from __future__ import annotations

import sys
import tempfile
import tarfile
import io
from pathlib import Path
from unittest.mock import patch

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "core" / "src"))
sys.path.insert(0, str(HERE.parent / "shell-flask"))

import build_deb
from scan2excel.web_app import WebApi


def check_deb_archive(root: Path) -> None:
    stage = root / "stage"
    control = stage / "DEBIAN" / "control"
    control.parent.mkdir(parents=True)
    control.write_text("Package: test\nVersion: 1.0\n", encoding="utf-8")
    payload_file = stage / "opt" / "ocr-guoku" / "app.bin"
    payload_file.parent.mkdir(parents=True)
    payload_file.write_bytes(b"fixture")
    (stage / "usr" / "bin").mkdir(parents=True)
    output = root / "fixture.deb"
    build_deb._make_deb(stage, output)

    raw = output.read_bytes()
    assert raw.startswith(b"!<arch>\n")
    offset = 8
    members = {}
    while offset < len(raw):
        header = raw[offset:offset + 60]
        name = header[:16].decode("ascii").strip().rstrip("/")
        size = int(header[48:58].decode("ascii").strip())
        start = offset + 60
        members[name] = raw[start:start + size]
        offset = start + size + size % 2
    assert list(members) == ["debian-binary", "control.tar.gz", "data.tar.gz"]
    assert members["debian-binary"] == b"2.0\n"
    with tarfile.open(fileobj=io.BytesIO(members["control.tar.gz"]), mode="r:gz") as archive:
        assert archive.extractfile("control").read().startswith(b"Package: test")
    with tarfile.open(fileobj=io.BytesIO(members["data.tar.gz"]), mode="r:gz") as archive:
        assert archive.extractfile("opt/ocr-guoku/app.bin").read() == b"fixture"


def main() -> int:
    with tempfile.TemporaryDirectory() as td:
        check_deb_archive(Path(td))
        api = WebApi(Path(td))
        with patch("scan2excel.web_app.sys.platform", "linux"):
            api.state["file_picker_mode"] = "自动"
            assert api._picker_mode() == "浏览器内置"
            api.set_file_picker_mode("自动")
            assert api.state["file_picker_mode"] == "浏览器内置"
            api.set_file_picker_mode("Tk 对话框")
            assert api.state["file_picker_mode"] == "浏览器内置"
            with patch("scan2excel.web_app._tk_open_images",
                       side_effect=AssertionError("Linux browser mode must not use Tk")):
                assert api._dialog_open_images() == []
            api.state["dp_mode"] = "国库数据校验归集"
            with patch("scan2excel.web_app._tk_open_data",
                       side_effect=AssertionError("Linux browser mode must not use Tk")):
                api.dp_pick_file("timeseries")
                api.tool_pick_files("excel")
            try:
                api.tool_convert("excel")
            except RuntimeError as exc:
                assert "Linux 版不包含 COM 引擎" in str(exc)
            else:
                raise AssertionError("Linux must explain that COM conversion is unavailable")

            file_path = Path(td) / "report.xlsx"
            file_path.write_bytes(b"fixture")
            with patch("scan2excel.web_app.subprocess.Popen") as popen:
                assert api.open_path(str(file_path)) is True
                assert popen.call_args.args[0] == ["xdg-open", str(file_path)]

            with patch("scan2excel.web_app._tk_save_file",
                       side_effect=AssertionError("Linux browser mode must not use Tk")):
                saved = api._dialog_save_file("export.xlsx")
            assert saved == str(Path(td) / "data" / "exports" / "export.xlsx")

    print("Linux 文件选择、xdg-open、用户导出路径与 .deb 结构检查通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
