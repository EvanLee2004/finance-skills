#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""在生成端完成 Excel 计算，并派生无历史外链的便携副本。"""

from __future__ import annotations

import os
import re
import json
import subprocess
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Mapping, Optional, Sequence, Set, Tuple


_CELL_RE = re.compile(r'<c\b[^>]*/>|<c\b[^>]*>.*?</c>', re.S)
_FORMULA_RE = re.compile(r'<f\b[^>]*/>|<f\b[^>]*>.*?</f>', re.S)
_EXTERNAL_FORMULA_RE = re.compile(r'^\s*\[\d+\]')


class WorkbookFinalizeError(RuntimeError):
    pass


@dataclass(frozen=True)
class FinalizeResult:
    mode: str
    formula_cells: int
    calc_chain_present: bool
    full_calc_on_load: str
    force_full_calc: str


@dataclass(frozen=True)
class PortableAudit:
    frozen_external_formulas: int
    external_link_parts: int
    external_references: int
    external_formula_cells: int
    display_value_mismatches: int
    formula_cells: int
    calc_chain_present: bool
    full_calc_on_load: str
    force_full_calc: str


def _entry_text(payload: Mapping[str, bytes], name: str) -> str:
    return payload[name].decode("utf-8", "replace")


def _read_package(path: Path):
    with zipfile.ZipFile(path) as zin:
        infos = zin.infolist()
        payload = {info.filename: zin.read(info.filename) for info in infos}
    return infos, payload


def _write_package(path: Path, infos, payload: Mapping[str, bytes]) -> None:
    temp = path.with_name(f".{path.stem}_package_{os.getpid()}{path.suffix}")
    temp.unlink(missing_ok=True)
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
            written = set()
            for info in infos:
                if info.filename in payload:
                    zout.writestr(info, payload[info.filename])
                    written.add(info.filename)
            for name, data in payload.items():
                if name not in written:
                    zout.writestr(name, data)
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


def _set_tag_attr(tag: str, name: str, value: str) -> str:
    pattern = rf'(\b{re.escape(name)}=")[^"]*(")'
    if re.search(pattern, tag):
        return re.sub(pattern, rf'\g<1>{value}\2', tag, count=1)
    return tag[:-2] + f' {name}="{value}"/>' if tag.endswith("/>") else tag


def _set_safe_calc_flags(workbook_xml: str) -> str:
    """保留 Excel 刚生成的 calcId，只关闭转嫁到收件电脑的全量重算标记。"""
    match = re.search(r'<calcPr\b[^>]*/>', workbook_xml)
    if not match:
        return workbook_xml.replace(
            "</workbook>",
            '<calcPr calcMode="auto" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>',
        )
    tag = match.group(0)
    tag = _set_tag_attr(tag, "calcMode", "auto")
    tag = _set_tag_attr(tag, "fullCalcOnLoad", "0")
    tag = _set_tag_attr(tag, "forceFullCalc", "0")
    return workbook_xml[: match.start()] + tag + workbook_xml[match.end() :]


def _calc_attrs(workbook_xml: str) -> Tuple[str, str]:
    tag_match = re.search(r'<calcPr\b[^>]*/>', workbook_xml)
    tag = tag_match.group(0) if tag_match else ""

    def attr(name: str) -> str:
        match = re.search(rf'\b{re.escape(name)}="([^"]*)"', tag)
        return match.group(1) if match else ""

    return attr("fullCalcOnLoad"), attr("forceFullCalc")


def _formula_cell_count(payload: Mapping[str, bytes]) -> int:
    return sum(
        len(_FORMULA_RE.findall(_entry_text(payload, name)))
        for name in payload
        if re.match(r'^xl/worksheets/sheet\d+\.xml$', name)
    )


def _formula_cells_missing_cache(payload: Mapping[str, bytes]) -> int:
    missing = 0
    for name in payload:
        if not re.match(r'^xl/worksheets/sheet\d+\.xml$', name):
            continue
        for cell in _CELL_RE.findall(_entry_text(payload, name)):
            if _FORMULA_RE.search(cell) and not re.search(r'<v\b[^>]*>.*?</v>|<v\b[^>]*/>', cell, re.S):
                missing += 1
    return missing


def _sheet_parts_in_order(payload: Mapping[str, bytes]) -> Sequence[Tuple[int, str]]:
    workbook_xml = _entry_text(payload, "xl/workbook.xml")
    rels_xml = _entry_text(payload, "xl/_rels/workbook.xml.rels")
    targets: Dict[str, str] = {}
    for match in re.finditer(r'<Relationship\b([^>]*)/?>', rels_xml):
        attrs = match.group(1)
        rid = re.search(r'\bId="([^"]+)"', attrs)
        target = re.search(r'\bTarget="([^"]+)"', attrs)
        if rid and target:
            part = target.group(1).lstrip("/")
            targets[rid.group(1)] = part if part.startswith("xl/") else f"xl/{part}"
    result = []
    for index, match in enumerate(re.finditer(r'<sheet\b([^>]*)/?>', workbook_xml), 1):
        rid = re.search(r'\br:id="([^"]+)"', match.group(1))
        if rid and rid.group(1) in targets:
            result.append((index, targets[rid.group(1)]))
    return result


def _ensure_calc_chain_registration(payload: Dict[str, bytes]) -> None:
    rels_name = "xl/_rels/workbook.xml.rels"
    rels_xml = _entry_text(payload, rels_name)
    if not re.search(r'\bType="[^"]*/calcChain"', rels_xml):
        used = {int(x) for x in re.findall(r'\bId="rId(\d+)"', rels_xml)}
        next_id = 1
        while next_id in used:
            next_id += 1
        relationship = (
            f'<Relationship Id="rId{next_id}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" '
            'Target="calcChain.xml"/>'
        )
        rels_xml = rels_xml.replace("</Relationships>", relationship + "</Relationships>")
        payload[rels_name] = rels_xml.encode("utf-8")

    content_name = "[Content_Types].xml"
    content_xml = _entry_text(payload, content_name)
    if 'PartName="/xl/calcChain.xml"' not in content_xml:
        override = (
            '<Override PartName="/xl/calcChain.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
        )
        content_xml = content_xml.replace("</Types>", override + "</Types>")
        payload[content_name] = content_xml.encode("utf-8")


def _rebuild_calc_chain(payload: Dict[str, bytes]) -> int:
    entries = []
    for sheet_index, part in _sheet_parts_in_order(payload):
        if part not in payload:
            continue
        for cell in _CELL_RE.findall(_entry_text(payload, part)):
            if not _FORMULA_RE.search(cell):
                continue
            ref = re.search(r'\br="([A-Z]{1,3}\d+)"', cell)
            if ref:
                entries.append(f'<c r="{ref.group(1)}" i="{sheet_index}"/>')
    chain = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + "".join(entries)
        + "</calcChain>"
    )
    payload["xl/calcChain.xml"] = chain.encode("utf-8")
    _ensure_calc_chain_registration(payload)
    return len(entries)


def _external_formula_count(payload: Mapping[str, bytes]) -> int:
    count = 0
    for name in payload:
        if not re.match(r'^xl/worksheets/sheet\d+\.xml$', name):
            continue
        xml = _entry_text(payload, name)
        external_shared: Set[str] = set()
        for formula in _FORMULA_RE.findall(xml):
            body = re.search(r'<f\b[^>]*>(.*?)</f>', formula, re.S)
            if body and _EXTERNAL_FORMULA_RE.match(body.group(1)):
                si = re.search(r'\bsi="([^"]+)"', formula)
                if si:
                    external_shared.add(si.group(1))
        for cell in _CELL_RE.findall(xml):
            formula_match = _FORMULA_RE.search(cell)
            if not formula_match:
                continue
            formula = formula_match.group(0)
            body = re.search(r'<f\b[^>]*>(.*?)</f>', formula, re.S)
            si = re.search(r'\bsi="([^"]+)"', formula)
            if (
                body
                and _EXTERNAL_FORMULA_RE.match(body.group(1))
            ) or (si and si.group(1) in external_shared):
                count += 1
    return count


def inspect_calculation(path: Path) -> FinalizeResult:
    _, payload = _read_package(Path(path))
    workbook_xml = _entry_text(payload, "xl/workbook.xml")
    full, force = _calc_attrs(workbook_xml)
    return FinalizeResult(
        mode="inspection",
        formula_cells=_formula_cell_count(payload),
        calc_chain_present="xl/calcChain.xml" in payload,
        full_calc_on_load=full,
        force_full_calc=force,
    )


def external_link_count(path: Path) -> int:
    _, payload = _read_package(Path(path))
    return sum(
        1
        for name in payload
        if re.match(r'^xl/externalLinks/externalLink\d+\.xml$', name)
    )


def _record_excel_pid(excel, pid_file: Optional[Path]) -> None:
    if pid_file is None:
        return
    try:
        import ctypes

        process_id = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(
            int(excel.Hwnd), ctypes.byref(process_id)
        )
        pid_file.write_text(str(process_id.value), encoding="ascii")
    except Exception:
        pass


def _write_status(status_file: Optional[Path], value: str) -> None:
    if status_file is not None:
        status_file.write_text(value, encoding="utf-8")


def _native_excel_recalculate_inline(
    path: Path,
    *,
    full_rebuild: bool,
    dirty_cells_by_sheet: Mapping[str, Sequence[str]],
    timeout_seconds: int,
    pid_file: Optional[Path] = None,
    status_file: Optional[Path] = None,
) -> str:
    if os.name != "nt":
        raise WorkbookFinalizeError("生成端不是 Windows，无法调用本机 Excel 完成计算")
    try:
        import pythoncom
        import win32com.client
    except Exception as exc:
        raise WorkbookFinalizeError("缺少 pywin32，无法调用本机 Excel 完成计算") from exc

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    save_copy = Path(path).with_name(
        f".{Path(path).stem}_excel_save_{os.getpid()}{Path(path).suffix}"
    )
    save_copy.unlink(missing_ok=True)
    mode = "依赖链重建" if full_rebuild else "定点增量计算"
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        _record_excel_pid(excel, pid_file)
        _write_status(status_file, "Excel 已启动，准备打开工作簿")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        try:
            excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:
            pass
        try:
            excel.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass
        try:
            excel.CalculateBeforeSave = False
        except Exception:
            pass
        workbook = excel.Workbooks.Open(
            Filename=str(Path(path).resolve()),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )
        workbook.Activate()
        try:
            workbook.ForceFullCalculation = False
            workbook.FullCalculationOnLoad = False
        except Exception:
            pass
        _write_status(status_file, "工作簿已打开，正在定点计算")

        # 写入脚本已为新公式保存了精确缓存值。这里让 Excel 只重算本次写入格及其
        # 直接依赖，不调用会把 5 万多个公式和历史外链全部卷入的全量计算。
        for sheet_name, cells in dirty_cells_by_sheet.items():
            sheet = workbook.Worksheets(sheet_name)
            for address in cells:
                cell = sheet.Range(address)
                if bool(cell.HasFormula):
                    cell.Calculate()
                    continue
                value = cell.Value2
                cell.Value2 = value
                try:
                    cell.Dependents.Calculate()
                except Exception:
                    # 没有同表直接依赖时 Excel 会抛 COM 异常，这是正常情况。
                    pass
        _write_status(status_file, "定点计算完成，正在保存")

        deadline = time.monotonic() + max(int(timeout_seconds), 1)
        while int(excel.CalculationState) != 0:
            if time.monotonic() >= deadline:
                raise WorkbookFinalizeError(
                    f"Excel 计算超过 {timeout_seconds} 秒，未交付未完成计算的文件"
                )
            time.sleep(0.2)
        try:
            workbook.CheckCompatibility = False
        except Exception:
            pass
        workbook.SaveCopyAs(str(save_copy))
        _write_status(status_file, "副本保存完成，正在关闭")
        workbook.Close(SaveChanges=False)
        workbook = None
        save_copy.replace(Path(path))
    except WorkbookFinalizeError:
        raise
    except Exception as exc:
        raise WorkbookFinalizeError(
            f"本机 Excel 计算保存失败：{type(exc).__name__}"
        ) from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        save_copy.unlink(missing_ok=True)
        pythoncom.CoUninitialize()
    return mode


def _kill_owned_process(pid: int) -> None:
    if pid <= 0:
        return
    subprocess.run(
        ["taskkill", "/PID", str(pid), "/T", "/F"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def _native_excel_recalculate(
    path: Path,
    *,
    full_rebuild: bool,
    dirty_cells_by_sheet: Mapping[str, Sequence[str]],
    timeout_seconds: int,
) -> str:
    """在隔离进程调用 Excel，使 COM 阻塞时仍能执行硬超时。"""
    path = Path(path).resolve()
    temp_root = Path(tempfile.mkdtemp(prefix="ar_hexiao_excel_"))
    request_file = temp_root / "request.json"
    result_file = temp_root / "result.json"
    pid_file = temp_root / "excel.pid"
    status_file = temp_root / "status.txt"
    request_file.write_text(
        json.dumps(
            {
                "path": str(path),
                "full_rebuild": bool(full_rebuild),
                "dirty_cells_by_sheet": {
                    name: list(cells)
                    for name, cells in dirty_cells_by_sheet.items()
                },
                "timeout_seconds": max(int(timeout_seconds), 1),
                "pid_file": str(pid_file),
                "status_file": str(status_file),
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    process = subprocess.Popen(
        [sys.executable, str(Path(__file__).resolve()), "--native-worker", str(request_file), str(result_file)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    try:
        stdout, stderr = process.communicate(timeout=max(int(timeout_seconds), 1))
    except subprocess.TimeoutExpired as exc:
        excel_pid = 0
        try:
            excel_pid = int(pid_file.read_text(encoding="ascii").strip())
        except Exception:
            pass
        _kill_owned_process(excel_pid)
        process.kill()
        process.communicate()
        try:
            status = status_file.read_text(encoding="utf-8").strip()
        except Exception:
            status = "未知阶段"
        for child in (request_file, result_file, pid_file, status_file):
            child.unlink(missing_ok=True)
        try:
            temp_root.rmdir()
        except OSError:
            pass
        raise WorkbookFinalizeError(
            f"Excel 处理超过 {timeout_seconds} 秒（{status}），已终止本次隐藏计算进程"
        ) from exc
    finally:
        if process.poll() is None:
            process.kill()
            process.communicate()

    try:
        result = json.loads(result_file.read_text(encoding="utf-8"))
    except Exception as exc:
        detail = (stderr or stdout or "无返回信息").strip()[-500:]
        raise WorkbookFinalizeError(f"Excel 隔离计算没有有效结果：{detail}") from exc
    finally:
        for child in (request_file, result_file, pid_file, status_file):
            child.unlink(missing_ok=True)
        try:
            temp_root.rmdir()
        except OSError:
            pass

    if process.returncode != 0 or not result.get("ok"):
        raise WorkbookFinalizeError(str(result.get("error") or "Excel 隔离计算失败"))
    return str(result["mode"])


def finalize_workbook(
    path: Path,
    patch_results_by_sheet: Mapping[str, object],
    *,
    timeout_seconds: int = 300,
) -> FinalizeResult:
    """保留公式缓存并补齐计算链，避免把全量重算转嫁给打开文件的电脑。"""
    path = Path(path)
    infos, payload = _read_package(path)
    missing_cache = _formula_cells_missing_cache(payload)
    if missing_cache:
        raise WorkbookFinalizeError(
            f"有 {missing_cache} 个公式没有缓存值，拒绝生成可能显示空白的文件"
        )
    needs_chain_rebuild = "xl/calcChain.xml" not in payload or any(
        bool(getattr(result, "requires_full_rebuild", True))
        for result in patch_results_by_sheet.values()
    )
    mode = "保留原计算链"
    if needs_chain_rebuild:
        rebuilt = _rebuild_calc_chain(payload)
        if rebuilt != _formula_cell_count(payload):
            raise WorkbookFinalizeError("计算链条目数与公式单元格数不一致")
        mode = "按公式缓存重建计算链"
    workbook_name = "xl/workbook.xml"
    payload[workbook_name] = _set_safe_calc_flags(
        _entry_text(payload, workbook_name)
    ).encode("utf-8")
    _write_package(path, infos, payload)

    result = inspect_calculation(path)
    if result.formula_cells and not result.calc_chain_present:
        raise WorkbookFinalizeError("Excel 保存后仍没有计算链，拒绝交付可能再次全量重算的文件")
    if result.full_calc_on_load not in ("", "0") or result.force_full_calc not in ("", "0"):
        raise WorkbookFinalizeError("写后全量重算标记没有关闭")
    return FinalizeResult(
        mode=mode,
        formula_cells=result.formula_cells,
        calc_chain_present=result.calc_chain_present,
        full_calc_on_load=result.full_calc_on_load,
        force_full_calc=result.force_full_calc,
    )


def _sheet_part_to_id(payload: Mapping[str, bytes]) -> Dict[str, str]:
    workbook_xml = _entry_text(payload, "xl/workbook.xml")
    rels_xml = _entry_text(payload, "xl/_rels/workbook.xml.rels")
    rid_to_target: Dict[str, str] = {}
    for rel in re.finditer(r'<Relationship\b([^>]*)/?>', rels_xml):
        attrs = rel.group(1)
        rid = re.search(r'\bId="([^"]+)"', attrs)
        target = re.search(r'\bTarget="([^"]+)"', attrs)
        if rid and target:
            part = target.group(1).lstrip("/")
            if not part.startswith("xl/"):
                part = "xl/" + part
            rid_to_target[rid.group(1)] = part

    result: Dict[str, str] = {}
    for sheet in re.finditer(r'<sheet\b([^>]*)/?>', workbook_xml):
        attrs = sheet.group(1)
        sheet_id = re.search(r'\bsheetId="([^"]+)"', attrs)
        rid = re.search(r'\br:id="([^"]+)"', attrs)
        if sheet_id and rid and rid.group(1) in rid_to_target:
            result[rid_to_target[rid.group(1)]] = sheet_id.group(1)
    return result


def _freeze_external_formulas(sheet_xml: str) -> Tuple[str, Set[str]]:
    external_shared: Set[str] = set()
    for formula in _FORMULA_RE.findall(sheet_xml):
        body = re.search(r'<f\b[^>]*>(.*?)</f>', formula, re.S)
        if body and _EXTERNAL_FORMULA_RE.match(body.group(1)):
            si = re.search(r'\bsi="([^"]+)"', formula)
            if si:
                external_shared.add(si.group(1))

    frozen: Set[str] = set()

    def replace_cell(match) -> str:
        cell = match.group(0)
        ref_match = re.search(r'\br="([A-Z]{1,3}\d+)"', cell)
        formula_match = _FORMULA_RE.search(cell)
        if not ref_match or not formula_match:
            return cell
        formula = formula_match.group(0)
        body = re.search(r'<f\b[^>]*>(.*?)</f>', formula, re.S)
        si = re.search(r'\bsi="([^"]+)"', formula)
        is_external = bool(
            (body and _EXTERNAL_FORMULA_RE.match(body.group(1)))
            or (si and si.group(1) in external_shared)
        )
        if not is_external:
            return cell
        if not re.search(r'<v(?:\s[^>]*)?(?:/>|>.*?</v>)', cell, re.S):
            raise WorkbookFinalizeError(
                f"外链公式 {ref_match.group(1)} 没有缓存显示值，不能安全固化"
            )
        frozen.add(ref_match.group(1))
        return cell[: formula_match.start()] + cell[formula_match.end() :]

    return _CELL_RE.sub(replace_cell, sheet_xml), frozen


def _filter_calc_chain(calc_xml: str, frozen: Set[Tuple[str, str]]) -> str:
    pattern = re.compile(r'<c\b[^>]*/>')
    current_sheet: Optional[str] = None
    emitted_sheet: Optional[str] = None
    pieces = []
    cursor = 0
    for match in pattern.finditer(calc_xml):
        pieces.append(calc_xml[cursor : match.start()])
        tag = match.group(0)
        sheet_match = re.search(r'\bi="([^"]+)"', tag)
        if sheet_match:
            current_sheet = sheet_match.group(1)
        ref_match = re.search(r'\br="([^"]+)"', tag)
        if ref_match and current_sheet and (current_sheet, ref_match.group(1)) in frozen:
            cursor = match.end()
            continue
        if current_sheet and emitted_sheet != current_sheet and not sheet_match:
            tag = tag.replace("<c", f'<c i="{current_sheet}"', 1)
        pieces.append(tag)
        emitted_sheet = current_sheet
        cursor = match.end()
    pieces.append(calc_xml[cursor:])
    return "".join(pieces)


def _display_signatures(payload: Mapping[str, bytes]):
    signatures = {}
    for name in payload:
        if not re.match(r'^xl/worksheets/sheet\d+\.xml$', name):
            continue
        xml = _entry_text(payload, name)
        for cell in _CELL_RE.findall(xml):
            ref = re.search(r'\br="([A-Z]{1,3}\d+)"', cell)
            if not ref:
                continue
            style = re.search(r'\bs="([^"]+)"', cell)
            cell_type = re.search(r'\bt="([^"]+)"', cell)
            value = re.search(r'<v(?:\s[^>]*)?>(.*?)</v>', cell, re.S)
            inline = re.search(r'<is\b[^>]*>(.*?)</is>', cell, re.S)
            signatures[(name, ref.group(1))] = (
                style.group(1) if style else "",
                cell_type.group(1) if cell_type else "",
                value.group(1) if value else "",
                inline.group(1) if inline else "",
            )
    return signatures


def _portable_audit(
    source_payload: Mapping[str, bytes],
    portable_payload: Mapping[str, bytes],
    frozen_count: int,
) -> PortableAudit:
    workbook_xml = _entry_text(portable_payload, "xl/workbook.xml")
    full, force = _calc_attrs(workbook_xml)
    source_values = _display_signatures(source_payload)
    portable_values = _display_signatures(portable_payload)
    mismatches = sum(
        1
        for key in set(source_values) | set(portable_values)
        if source_values.get(key) != portable_values.get(key)
    )
    return PortableAudit(
        frozen_external_formulas=frozen_count,
        external_link_parts=sum(
            1
            for name in portable_payload
            if re.match(r'^xl/externalLinks/externalLink\d+\.xml$', name)
        ),
        external_references=len(re.findall(r'<externalReference\b', workbook_xml)),
        external_formula_cells=_external_formula_count(portable_payload),
        display_value_mismatches=mismatches,
        formula_cells=_formula_cell_count(portable_payload),
        calc_chain_present="xl/calcChain.xml" in portable_payload,
        full_calc_on_load=full,
        force_full_calc=force,
    )


def create_portable_copy(source: Path, out: Path) -> PortableAudit:
    """固化历史外链缓存值；保留工作底稿，便携副本不再访问旧外部文件。"""
    source, out = Path(source), Path(out)
    infos, source_payload = _read_package(source)
    payload = dict(source_payload)
    part_to_id = _sheet_part_to_id(payload)
    frozen: Set[Tuple[str, str]] = set()

    for part, sheet_id in part_to_id.items():
        if part not in payload:
            continue
        patched, refs = _freeze_external_formulas(_entry_text(payload, part))
        if refs:
            payload[part] = patched.encode("utf-8")
            frozen.update((sheet_id, ref) for ref in refs)

    if not frozen:
        raise WorkbookFinalizeError("工作簿没有可固化的历史外链公式")

    for name in list(payload):
        if name.startswith("xl/externalLinks/"):
            payload.pop(name, None)

    workbook_name = "xl/workbook.xml"
    workbook_xml = _entry_text(payload, workbook_name)
    workbook_xml = re.sub(
        r'<externalReferences\b[^>]*>.*?</externalReferences>',
        "",
        workbook_xml,
        flags=re.S,
    )
    payload[workbook_name] = _set_safe_calc_flags(workbook_xml).encode("utf-8")

    rels_name = "xl/_rels/workbook.xml.rels"
    rels_xml = _entry_text(payload, rels_name)
    rels_xml = re.sub(
        r'<Relationship\b(?=[^>]*\bType="[^"]*/externalLink")[^>]*/>',
        "",
        rels_xml,
    )
    payload[rels_name] = rels_xml.encode("utf-8")

    content_name = "[Content_Types].xml"
    content_xml = _entry_text(payload, content_name)
    content_xml = re.sub(
        r'<Override\b(?=[^>]*\bPartName="/xl/externalLinks/)[^>]*/>',
        "",
        content_xml,
    )
    payload[content_name] = content_xml.encode("utf-8")

    if "xl/calcChain.xml" in payload:
        payload["xl/calcChain.xml"] = _filter_calc_chain(
            _entry_text(payload, "xl/calcChain.xml"), frozen
        ).encode("utf-8")
    elif _formula_cell_count(payload):
        raise WorkbookFinalizeError("工作底稿没有计算链，不能生成免重算便携副本")

    out.parent.mkdir(parents=True, exist_ok=True)
    if out.exists():
        out.unlink()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        written = set()
        for info in infos:
            if info.filename in payload:
                zout.writestr(info, payload[info.filename])
                written.add(info.filename)
        for name, data in payload.items():
            if name not in written:
                zout.writestr(name, data)

    _, portable_payload = _read_package(out)
    audit = _portable_audit(source_payload, portable_payload, len(frozen))
    problems = []
    if audit.external_link_parts or audit.external_references or audit.external_formula_cells:
        problems.append("外部链接没有清理干净")
    if audit.display_value_mismatches:
        problems.append(f"显示值变化 {audit.display_value_mismatches} 格")
    if audit.formula_cells and not audit.calc_chain_present:
        problems.append("剩余公式没有计算链")
    if audit.full_calc_on_load not in ("", "0") or audit.force_full_calc not in ("", "0"):
        problems.append("便携副本仍要求打开时全量重算")
    if problems:
        out.unlink(missing_ok=True)
        raise WorkbookFinalizeError("；".join(problems))
    return audit


def portable_path_for(path: Path) -> Path:
    path = Path(path)
    return path.with_name(f"{path.stem}_便携版{path.suffix}")


def _native_worker(request_path: Path, result_path: Path) -> int:
    try:
        request = json.loads(Path(request_path).read_text(encoding="utf-8"))
        mode = _native_excel_recalculate_inline(
            Path(request["path"]),
            full_rebuild=bool(request["full_rebuild"]),
            dirty_cells_by_sheet=request.get("dirty_cells_by_sheet", {}),
            timeout_seconds=int(request.get("timeout_seconds", 300)),
            pid_file=Path(request["pid_file"]),
            status_file=Path(request["status_file"]),
        )
        result = {"ok": True, "mode": mode}
        code = 0
    except Exception as exc:
        result = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
        code = 1
    Path(result_path).write_text(
        json.dumps(result, ensure_ascii=False), encoding="utf-8"
    )
    return code


if __name__ == "__main__" and len(sys.argv) == 4 and sys.argv[1] == "--native-worker":
    raise SystemExit(_native_worker(Path(sys.argv[2]), Path(sys.argv[3])))
