"""Prepare public Linux assets without publishing local workbook data.

Only the three seed workbooks are changed; application binaries stay identical.
Usage: python shell-flask/prepare_public_release.py --source DIR --output DIR
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import io
import json
import re
import tarfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import range_boundaries
from openpyxl.workbook.properties import CalcProperties

from build_deb import _ar_member

SEEDS = ("config/config.xlsx", "templates/Template.xlsx", "transforms/国库报表.xlsx")


def clean_workbook(payload: bytes, kind: str) -> bytes:
    wb = load_workbook(io.BytesIO(payload))
    if kind == "templates/Template.xlsx":
        definition = wb.defined_names.get("数字区域")
        if definition is None:
            raise ValueError("Template has no numeric region definition")
        for sheet, area in definition.destinations:
            ws = wb[sheet]
            left, top, right, bottom = range_boundaries(area)
            for row in ws.iter_rows(min_row=top, max_row=min(bottom, ws.max_row),
                                    min_col=left, max_col=min(right, ws.max_column)):
                for cell in row:
                    if not isinstance(cell, MergedCell) and cell.data_type != "f":
                        cell.value = None
    elif kind == "transforms/国库报表.xlsx":
        # A fresh data-only target keeps the exact required column headers.
        headers = [c.value for c in wb["源数据时序"][1]]
        wb.close()
        wb = Workbook()
        wb.active.title = "源数据时序"
        wb.active.append(headers)
    elif kind == "config/config.xlsx":
        for ws in wb:
            for row in ws:
                for cell in row:
                    if isinstance(cell.value, str) and re.search(r"[A-Za-z]:[\\/]", cell.value):
                        if cell.value.replace("\\", "/").endswith("/core/data/transforms/国库报表.xlsx"):
                            cell.value = r"C:\OCR_Tools\core\data\transforms\国库报表.xlsx"
                        else:
                            raise ValueError(f"Unrecognized local path in {kind}: {ws.title}!{cell.coordinate}")
    for ws in wb:
        for row in ws:
            for cell in row:
                if cell.comment is not None:
                    cell.comment = None
    wb.properties.creator = "OCR Tools"
    wb.properties.lastModifiedBy = "OCR Tools"
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    output = io.BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue()


def rewrite_tar(source, destination) -> dict[str, str]:
    changed = {}
    with tarfile.open(fileobj=source, mode="r|gz") as old:
        with tarfile.open(fileobj=destination, mode="w|gz", format=tarfile.GNU_FORMAT, compresslevel=6) as new:
            for member in old:
                info = copy.copy(member)
                info.pax_headers = {}
                info.mtime = int(info.mtime)
                stream = old.extractfile(member) if member.isfile() else None
                kind = next((k for k in SEEDS if member.name.endswith("/core/data/" + k)), None)
                if kind:
                    payload = clean_workbook(stream.read(), kind)
                    info.size = len(payload)
                    stream = io.BytesIO(payload)
                    changed[kind] = hashlib.sha256(payload).hexdigest()
                if "/core/data/" in member.name and member.isfile() and kind is None:
                    raise ValueError(f"Unexpected runtime data in release: {member.name}")
                new.addfile(info, stream)
    if set(changed) != set(SEEDS):
        raise ValueError(f"Incomplete release seeds: {list(changed)}")
    return changed


def rewrite_deb(source: Path, output: Path) -> dict[str, str]:
    changed = {}
    with source.open("rb") as old, output.open("wb") as new:
        if old.read(8) != b"!<arch>\n":
            raise ValueError("Invalid Debian archive")
        new.write(b"!<arch>\n")
        while header := old.read(60):
            name = header[:16].decode("ascii").strip().rstrip("/")
            size = int(header[48:58])
            payload = old.read(size)
            if size % 2:
                old.read(1)
            if name == "data.tar.gz":
                buffer = io.BytesIO()
                changed = rewrite_tar(io.BytesIO(payload), buffer)
                payload = buffer.getvalue()
            new.write(_ar_member(name, payload))
    return changed


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.source.resolve() == args.output.resolve():
        parser.error("Public output must be separate from the private package directory")
    args.output.mkdir(parents=True, exist_ok=True)
    assets = list(args.source.glob("*.deb")) + list(args.source.glob("*.tar.gz"))
    if len(assets) != 2:
        parser.error("Expected exactly one .deb and one .tar.gz")
    records = {}
    for source in assets:
        output = args.output / source.name
        if output.exists():
            raise FileExistsError(f"Refusing to overwrite prepared asset: {output}")
        if source.suffix == ".deb":
            records[source.name] = rewrite_deb(source, output)
        else:
            with source.open("rb") as old, output.open("wb") as new:
                records[source.name] = rewrite_tar(old, new)
        print(f"Prepared public asset: {source.name}", flush=True)
    checksums = []
    for source in assets:
        output = args.output / source.name
        digest = hashlib.sha256()
        with output.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
        checksums.append(f"{digest.hexdigest()}  {output.name}\n")
    (args.output / "SHA256SUMS").write_text("".join(checksums), encoding="ascii")
    (args.output / "PUBLIC-SEEDS.json").write_text(json.dumps({
        "policy": "No historical target rows or template amounts; only public rules and structure",
        "workbooks": records,
    }, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
