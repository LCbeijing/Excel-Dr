from __future__ import annotations

import argparse
import queue
import tempfile
import threading
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "docrel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

for prefix, uri in NS.items():
    ET.register_namespace("" if prefix == "main" else prefix, uri)


ProgressCallback = Callable[[int, int, str], None]


def report_progress(progress: Optional[ProgressCallback], current: int, total: int, message: str) -> None:
    if progress:
        progress(max(0, current), max(1, total), message)


@dataclass
class AnchorInfo:
    start_col: Optional[int]
    start_row: Optional[int]
    hidden: bool
    embed_id: Optional[str]
    image_path: Optional[str]
    descr: str

    @property
    def start_key(self) -> Tuple[Optional[int], Optional[int]]:
        return (self.start_col, self.start_row)

    @property
    def signature(self) -> Tuple[Optional[int], Optional[int], Optional[str], Optional[str], str]:
        return (self.start_col, self.start_row, self.embed_id, self.image_path, self.descr)


@dataclass
class DrawingPlan:
    rel_id: str
    drawing_path: str
    anchor_total: int
    hidden_total: int
    suspicious_signatures: Counter = field(default_factory=Counter)
    suspicious_total: int = 0
    remove_entire_drawing: bool = False
    summary: List[str] = field(default_factory=list)


@dataclass
class SheetPlan:
    sheet_name: str
    sheet_path: str
    drawing_plans: List[DrawingPlan] = field(default_factory=list)
    broken_validations: List[str] = field(default_factory=list)


@dataclass
class WorkbookReport:
    source: Path
    output: Optional[Path] = None
    sheet_plans: List[SheetPlan] = field(default_factory=list)
    findings: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    removed_anchors: int = 0
    removed_drawings: int = 0
    removed_validations: int = 0

    @property
    def suspicious_total(self) -> int:
        return sum(
            drawing.suspicious_total
            for sheet in self.sheet_plans
            for drawing in sheet.drawing_plans
        )

    @property
    def broken_validation_total(self) -> int:
        return sum(len(sheet.broken_validations) for sheet in self.sheet_plans)

    @property
    def needs_cleanup(self) -> bool:
        return self.suspicious_total > 0 or self.broken_validation_total > 0

    def render(self) -> str:
        lines = [f"文件: {self.source}"]
        if self.output:
            lines.append(f"输出: {self.output}")
        lines.append("")
        lines.append("检测结果:")
        if self.findings:
            lines.extend(f"- {line}" for line in self.findings)
        else:
            lines.append("- 未发现需要清理的问题")

        if self.sheet_plans:
            lines.append("")
            lines.append("明细:")
            for sheet in self.sheet_plans:
                lines.append(f"- 工作表 {sheet.sheet_name}")
                for drawing in sheet.drawing_plans:
                    lines.append(
                        f"  绘图关系 {drawing.rel_id}: 总对象 {drawing.anchor_total}，隐藏 {drawing.hidden_total}，计划清理 {drawing.suspicious_total}"
                    )
                    for summary in drawing.summary:
                        lines.append(f"    {summary}")
                for item in sheet.broken_validations:
                    lines.append(f"  损坏数据有效性: {item}")

        if self.warnings:
            lines.append("")
            lines.append("提示:")
            lines.extend(f"- {line}" for line in self.warnings)

        if self.output:
            lines.append("")
            lines.append("清理结果:")
            lines.append(f"- 删除异常隐藏对象: {self.removed_anchors}")
            lines.append(f"- 移除空绘图关系: {self.removed_drawings}")
            lines.append(f"- 移除损坏数据有效性: {self.removed_validations}")

        return "\n".join(lines)


@dataclass
class BatchResult:
    reports: List[WorkbookReport] = field(default_factory=list)
    skipped: List[Tuple[Path, str]] = field(default_factory=list)
    failed: List[Tuple[Path, str]] = field(default_factory=list)

    @property
    def file_count(self) -> int:
        return len(self.reports)

    @property
    def actionable_reports(self) -> List[WorkbookReport]:
        return [report for report in self.reports if report.needs_cleanup]

    @property
    def actionable_count(self) -> int:
        return len(self.actionable_reports)

    @property
    def suspicious_total(self) -> int:
        return sum(report.suspicious_total for report in self.reports)

    @property
    def broken_validation_total(self) -> int:
        return sum(report.broken_validation_total for report in self.reports)

    def render(self) -> str:
        lines = [
            f"扫描文件数: {self.file_count}",
            f"需要处理的文件: {self.actionable_count}",
            f"异常隐藏对象: {self.suspicious_total}",
            f"损坏数据有效性: {self.broken_validation_total}",
        ]
        if self.skipped:
            lines.append("")
            lines.append("跳过:")
            lines.extend(f"- {path.name}: {reason}" for path, reason in self.skipped)
        if self.failed:
            lines.append("")
            lines.append("失败:")
            lines.extend(f"- {path.name}: {reason}" for path, reason in self.failed)
        if self.reports:
            lines.append("")
            lines.append("文件明细:")
            for report in self.reports:
                status = "需要清理" if report.needs_cleanup else "正常"
                lines.append(
                    f"- {report.source.name}: {status}，异常对象 {report.suspicious_total}，损坏有效性 {report.broken_validation_total}"
                )
        return "\n".join(lines)


def localname(tag: str) -> str:
    return tag.split("}", 1)[-1]


def source_part_from_rels(rel_path: str) -> str:
    rel_path = rel_path.replace("\\", "/")
    if "/_rels/" not in rel_path or not rel_path.endswith(".rels"):
        return rel_path
    left, right = rel_path.split("/_rels/", 1)
    owner_name = right[:-5]
    return f"{left}/{owner_name}" if left else owner_name


def resolve_zip_target(base_part: str, target: str) -> str:
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    base_parts = Path(base_part).parent.parts
    target_path = Path(*base_parts, target)
    normalized: List[str] = []
    for part in target_path.parts:
        if part in ("", "."):
            continue
        if part == "..":
            if normalized:
                normalized.pop()
        else:
            normalized.append(part)
    return "/".join(normalized)


def read_relationships(zipf: zipfile.ZipFile, rel_path: str) -> Dict[str, str]:
    if rel_path not in zipf.namelist():
        return {}
    owner_part = source_part_from_rels(rel_path)
    root = ET.fromstring(zipf.read(rel_path))
    mapping: Dict[str, str] = {}
    for rel in root.findall("rel:Relationship", NS):
        mapping[rel.attrib["Id"]] = resolve_zip_target(owner_part, rel.attrib["Target"])
    return mapping


def parse_anchor(anchor: ET.Element, drawing_rels: Dict[str, str]) -> AnchorInfo:
    from_node = anchor.find("xdr:from", NS)
    start_col = None
    start_row = None
    if from_node is not None:
        col_text = from_node.findtext("xdr:col", default=None, namespaces=NS)
        row_text = from_node.findtext("xdr:row", default=None, namespaces=NS)
        start_col = int(col_text) if col_text is not None else None
        start_row = int(row_text) if row_text is not None else None

    c_nv_pr = anchor.find(".//xdr:cNvPr", NS)
    hidden = c_nv_pr is not None and c_nv_pr.attrib.get("hidden") == "1"
    descr = c_nv_pr.attrib.get("descr", "") if c_nv_pr is not None else ""

    blip = anchor.find(".//a:blip", NS)
    embed_id = None
    image_path = None
    if blip is not None:
        embed_id = blip.attrib.get(f"{{{NS['docrel']}}}embed")
        if embed_id:
            image_path = drawing_rels.get(embed_id)

    return AnchorInfo(
        start_col=start_col,
        start_row=start_row,
        hidden=hidden,
        embed_id=embed_id,
        image_path=image_path,
        descr=descr,
    )


def detect_suspicious_signatures(anchors: List[AnchorInfo]) -> Counter:
    hidden = [anchor for anchor in anchors if anchor.hidden]
    if len(hidden) < 100:
        return Counter()

    by_position = Counter(anchor.start_key for anchor in hidden)
    by_signature = Counter(anchor.signature for anchor in hidden)
    suspicious = Counter()

    for anchor in hidden:
        if by_position[anchor.start_key] >= 50 and by_signature[anchor.signature] >= 100:
            suspicious[anchor.signature] += 1

    if suspicious:
        return suspicious

    if len(hidden) >= 5000:
        dominant_signature, dominant_count = by_signature.most_common(1)[0]
        if dominant_count / len(hidden) >= 0.9:
            suspicious[dominant_signature] = dominant_count

    return suspicious


def summarize_signature_counts(signature_counts: Counter) -> List[str]:
    result: List[str] = []
    if not signature_counts:
        return result

    by_position = Counter()
    by_image = Counter()
    for sig, count in signature_counts.items():
        by_position[(sig[0], sig[1])] += count
        by_image[sig[3] or "未知"] += count

    top_positions = by_position.most_common(3)
    top_images = by_image.most_common(3)
    if top_positions:
        result.append("重复锚点 " + "、".join(f"(col={pos[0]}, row={pos[1]}) x {count}" for pos, count in top_positions))
    if top_images:
        result.append("重复图片资源 " + "、".join(f"{img} x {count}" for img, count in top_images))
    return result


def analyze_workbook(path: Path, progress: Optional[ProgressCallback] = None) -> WorkbookReport:
    report = WorkbookReport(source=path)
    with zipfile.ZipFile(path) as zipf:
        names = zipf.namelist()
        worksheet_count = sum(1 for name in names if name.startswith("xl/worksheets/") and name.endswith(".xml"))
        drawing_count = sum(1 for name in names if name.startswith("xl/drawings/") and name.endswith(".xml"))
        total_units = max(1, 1 + worksheet_count + drawing_count)
        completed_units = 0
        report_progress(progress, completed_units, total_units, "正在读取工作簿结构")

        workbook_root = ET.fromstring(zipf.read("xl/workbook.xml"))
        wb_rels = read_relationships(zipf, "xl/_rels/workbook.xml.rels")
        completed_units += 1
        report_progress(progress, completed_units, total_units, "已读取工作簿结构")

        for sheet in workbook_root.findall("main:sheets/main:sheet", NS):
            sheet_name = sheet.attrib.get("name", "")
            sheet_rid = sheet.attrib.get(f"{{{NS['docrel']}}}id", "")
            sheet_path = wb_rels.get(sheet_rid, "")
            if not sheet_path:
                continue

            report_progress(progress, completed_units, total_units, f"正在检测工作表：{sheet_name}")
            sheet_plan = SheetPlan(sheet_name=sheet_name, sheet_path=sheet_path)
            sheet_root = ET.fromstring(zipf.read(sheet_path))
            sheet_rel_path = f"{Path(sheet_path).parent.as_posix()}/_rels/{Path(sheet_path).name}.rels"
            sheet_rels = read_relationships(zipf, sheet_rel_path)
            completed_units += 1
            report_progress(progress, completed_units, total_units, f"已读取工作表：{sheet_name}")

            for drawing_node in sheet_root.findall("main:drawing", NS):
                rel_id = drawing_node.attrib.get(f"{{{NS['docrel']}}}id", "")
                drawing_path = sheet_rels.get(rel_id, "")
                if not drawing_path or drawing_path not in zipf.namelist():
                    continue

                report_progress(progress, completed_units, total_units, f"正在检测绘图对象：{sheet_name}")
                drawing_root = ET.fromstring(zipf.read(drawing_path))
                drawing_rel_path = f"{Path(drawing_path).parent.as_posix()}/_rels/{Path(drawing_path).name}.rels"
                drawing_rels = read_relationships(zipf, drawing_rel_path)
                anchors = [
                    parse_anchor(anchor, drawing_rels)
                    for anchor in list(drawing_root)
                    if localname(anchor.tag).endswith("Anchor")
                ]
                if not anchors:
                    continue

                suspicious = detect_suspicious_signatures(anchors)
                drawing_plan = DrawingPlan(
                    rel_id=rel_id,
                    drawing_path=drawing_path,
                    anchor_total=len(anchors),
                    hidden_total=sum(1 for anchor in anchors if anchor.hidden),
                    suspicious_signatures=suspicious,
                    suspicious_total=sum(suspicious.values()),
                    remove_entire_drawing=sum(suspicious.values()) == len(anchors),
                    summary=summarize_signature_counts(suspicious),
                )
                if drawing_plan.suspicious_total:
                    report.findings.append(
                        f"{sheet_name} 存在异常隐藏绘图对象 {drawing_plan.suspicious_total} 个，来源关系 {rel_id}"
                    )
                sheet_plan.drawing_plans.append(drawing_plan)
                completed_units += 1
                report_progress(progress, completed_units, total_units, f"已检测绘图对象：{sheet_name}")

            for dv in sheet_root.findall(".//main:dataValidation", NS):
                formula1 = dv.attrib.get("formula1") or dv.findtext("main:formula1", default="", namespaces=NS)
                if "#REF!" in formula1:
                    sqref = dv.attrib.get("sqref", "")
                    sheet_plan.broken_validations.append(f"{sqref} -> {formula1}")

            if sheet_plan.broken_validations:
                report.findings.append(f"{sheet_name} 存在损坏的数据有效性 {len(sheet_plan.broken_validations)} 处")

            if sheet_plan.drawing_plans or sheet_plan.broken_validations:
                report.sheet_plans.append(sheet_plan)

    if not report.findings:
        report.warnings.append("未命中异常阈值，默认不会做任何删除")
    report.warnings.append("本工具不会修改单元格公式、共享字符串、样式或 cellImages 图片公式")
    report.warnings.append("默认仅清理被判定为异常膨胀的隐藏 drawing 对象和损坏数据有效性")
    report_progress(progress, total_units, total_units, "检测完成")
    return report


def remove_suspicious_anchors_from_drawing(
    drawing_bytes: bytes,
    drawing_rels: Dict[str, str],
    suspicious_signatures: Counter,
) -> Tuple[bytes, Counter]:
    root = ET.fromstring(drawing_bytes)
    remaining_by_rel = Counter()
    remaining_suspicious = suspicious_signatures.copy()

    for anchor in list(root):
        if not localname(anchor.tag).endswith("Anchor"):
            continue
        info = parse_anchor(anchor, drawing_rels)
        if remaining_suspicious.get(info.signature, 0) > 0:
            remaining_suspicious[info.signature] -= 1
            root.remove(anchor)
            continue
        if info.embed_id:
            remaining_by_rel[info.embed_id] += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), remaining_by_rel


def remove_broken_validations(sheet_root: ET.Element) -> int:
    removed = 0
    data_validations = sheet_root.find("main:dataValidations", NS)
    if data_validations is None:
        return 0

    for dv in list(data_validations.findall("main:dataValidation", NS)):
        formula1 = dv.attrib.get("formula1") or dv.findtext("main:formula1", default="", namespaces=NS)
        if "#REF!" in formula1:
            data_validations.remove(dv)
            removed += 1

    if removed:
        remaining = data_validations.findall("main:dataValidation", NS)
        if remaining:
            data_validations.attrib["count"] = str(len(remaining))
        else:
            sheet_root.remove(data_validations)

    return removed


def clean_workbook_from_report(
    report: WorkbookReport,
    output: Path,
    progress: Optional[ProgressCallback] = None,
) -> WorkbookReport:
    source = report.source
    output = unique_output_path(output)
    report.output = output

    with zipfile.ZipFile(source) as zipf, tempfile.TemporaryDirectory() as temp_dir:
        names = zipf.namelist()
        clean_units = max(
            1,
            len(report.sheet_plans)
            + sum(1 for sheet in report.sheet_plans for drawing in sheet.drawing_plans if drawing.suspicious_total)
            + sum(1 for sheet in report.sheet_plans if sheet.broken_validations),
        )
        total_units = max(1, len(names) * 2 + clean_units)
        completed_units = 0
        for name in names:
            zipf.extract(name, temp_dir)
            completed_units += 1
            report_progress(progress, completed_units, total_units, f"正在解压：{Path(name).name}")

        root_dir = Path(temp_dir)

        for sheet_plan in report.sheet_plans:
            sheet_file = root_dir / Path(sheet_plan.sheet_path)
            if not sheet_file.exists():
                continue

            sheet_root = ET.fromstring(sheet_file.read_bytes())
            sheet_rel_path = Path(sheet_plan.sheet_path).parent / "_rels" / f"{Path(sheet_plan.sheet_path).name}.rels"
            sheet_rel_file = root_dir / sheet_rel_path
            sheet_rel_root = ET.fromstring(sheet_rel_file.read_bytes()) if sheet_rel_file.exists() else None
            rel_elements: Dict[str, ET.Element] = {}
            if sheet_rel_root is not None:
                for rel in sheet_rel_root.findall("rel:Relationship", NS):
                    rel_elements[rel.attrib["Id"]] = rel

            for drawing_plan in sheet_plan.drawing_plans:
                if not drawing_plan.suspicious_total:
                    continue

                drawing_file = root_dir / Path(drawing_plan.drawing_path)
                if not drawing_file.exists():
                    continue

                drawing_rel_path = Path(drawing_plan.drawing_path).parent / "_rels" / f"{Path(drawing_plan.drawing_path).name}.rels"
                drawing_rel_file = root_dir / drawing_rel_path
                drawing_rel_root = ET.fromstring(drawing_rel_file.read_bytes()) if drawing_rel_file.exists() else None
                drawing_rels: Dict[str, str] = {}
                drawing_rel_elements: Dict[str, ET.Element] = {}
                if drawing_rel_root is not None:
                    for rel in drawing_rel_root.findall("rel:Relationship", NS):
                        rel_id = rel.attrib["Id"]
                        drawing_rel_elements[rel_id] = rel
                        drawing_rels[rel_id] = resolve_zip_target(drawing_plan.drawing_path, rel.attrib["Target"])

                new_drawing_bytes, remaining_by_rel = remove_suspicious_anchors_from_drawing(
                    drawing_file.read_bytes(),
                    drawing_rels,
                    drawing_plan.suspicious_signatures,
                )
                drawing_file.write_bytes(new_drawing_bytes)
                report.removed_anchors += drawing_plan.suspicious_total

                if drawing_rel_root is not None:
                    for rel_id, rel in list(drawing_rel_elements.items()):
                        if remaining_by_rel.get(rel_id, 0) == 0:
                            target = rel.attrib.get("Target", "")
                            drawing_rel_root.remove(rel)
                            media_file = root_dir / resolve_zip_target(drawing_plan.drawing_path, target)
                            if media_file.exists():
                                media_file.unlink()
                    drawing_rel_file.write_bytes(
                        ET.tostring(drawing_rel_root, encoding="utf-8", xml_declaration=True)
                    )

                new_drawing_root = ET.fromstring(new_drawing_bytes)
                remaining_anchor_count = sum(
                    1 for child in list(new_drawing_root) if localname(child.tag).endswith("Anchor")
                )
                if remaining_anchor_count == 0 or drawing_plan.remove_entire_drawing:
                    for drawing_node in list(sheet_root.findall("main:drawing", NS)):
                        if drawing_node.attrib.get(f"{{{NS['docrel']}}}id") == drawing_plan.rel_id:
                            sheet_root.remove(drawing_node)
                    if sheet_rel_root is not None and drawing_plan.rel_id in rel_elements:
                        sheet_rel_root.remove(rel_elements[drawing_plan.rel_id])
                    if drawing_file.exists():
                        drawing_file.unlink()
                    if drawing_rel_file.exists():
                        drawing_rel_file.unlink()
                    report.removed_drawings += 1
                completed_units += 1
                report_progress(progress, completed_units, total_units, f"已清理绘图对象：{sheet_plan.sheet_name}")

            report.removed_validations += remove_broken_validations(sheet_root)
            sheet_file.write_bytes(ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True))
            if sheet_rel_root is not None:
                sheet_rel_file.write_bytes(ET.tostring(sheet_rel_root, encoding="utf-8", xml_declaration=True))
            completed_units += 1
            report_progress(progress, completed_units, total_units, f"已处理工作表：{sheet_plan.sheet_name}")

        files_to_write = [file_path for file_path in root_dir.rglob("*") if file_path.is_file()]
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
            for file_path in files_to_write:
                out_zip.write(file_path, file_path.relative_to(root_dir).as_posix())
                completed_units += 1
                report_progress(progress, completed_units, total_units, f"正在写入：{file_path.name}")

    report_progress(progress, total_units, total_units, "清理完成")
    return report


def clean_workbook(
    source: Path,
    output: Path,
    progress: Optional[ProgressCallback] = None,
) -> WorkbookReport:
    report = analyze_workbook(source, progress=progress)
    return clean_workbook_from_report(report, output, progress=progress)


def collect_xlsx_files(folder: Path) -> List[Path]:
    files = []
    for path in sorted(folder.rglob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        stem = path.stem.lower()
        if stem.endswith("_cleaned") or "_cleaned_" in stem:
            continue
        files.append(path)
    return files


def unique_output_path(output: Path) -> Path:
    if not output.exists():
        return output
    for index in range(2, 10000):
        candidate = output.with_name(f"{output.stem}_{index}{output.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法生成不覆盖原文件的输出路径: {output}")


def analyze_folder(folder: Path, progress: Optional[ProgressCallback] = None) -> BatchResult:
    result = BatchResult()
    files = collect_xlsx_files(folder)
    total = max(1, len(files))
    for index, path in enumerate(files, start=1):
        report_progress(progress, index - 1, total, f"正在检测：{path.name}")
        try:
            result.reports.append(analyze_workbook(path))
        except Exception as exc:
            result.failed.append((path, str(exc)))
        report_progress(progress, index, total, f"已检测：{path.name}")
    report_progress(progress, total, total, "文件夹检测完成")
    return result


def clean_folder_from_preview(
    preview: BatchResult,
    progress: Optional[ProgressCallback] = None,
) -> BatchResult:
    result = BatchResult()
    result.failed.extend(preview.failed)

    actionable = preview.actionable_reports
    total = max(1, len(actionable))
    cleaned_index = 0

    for report in preview.reports:
        if not report.needs_cleanup:
            result.reports.append(report)
            result.skipped.append((report.source, "正常，无需清理"))
            continue
        output = report.source.with_name(f"{report.source.stem}_cleaned{report.source.suffix}")
        cleaned_index += 1
        report_progress(progress, cleaned_index - 1, total, f"正在清理：{report.source.name}")
        try:
            result.reports.append(clean_workbook_from_report(report, output))
        except Exception as exc:
            result.failed.append((report.source, str(exc)))
        report_progress(progress, cleaned_index, total, f"已清理：{report.source.name}")

    report_progress(progress, total, total, "批量清理完成")
    return result


def clean_folder(folder: Path, progress: Optional[ProgressCallback] = None) -> BatchResult:
    preview = analyze_folder(folder, progress=progress)
    return clean_folder_from_preview(preview, progress=progress)


class CleanerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel-Dr")
        self.root.geometry("980x760")
        self.root.minsize(900, 680)

        self.mode_var = tk.StringVar(value="file")
        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.auto_output_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="请选择一个 .xlsx 文件或文件夹")
        self.percent_var = tk.StringVar(value="0%")
        self.progress_var = tk.DoubleVar(value=0)
        self.stat_file_var = tk.StringVar(value="-")
        self.stat_object_var = tk.StringVar(value="-")
        self.stat_rule_var = tk.StringVar(value="-")
        self.stat_output_var = tk.StringVar(value="-")

        self.task_queue: queue.Queue = queue.Queue()
        self.worker_active = False
        self.last_result: Optional[object] = None

        self.build_ui()
        self.root.after(100, self.process_task_queue)

    def build_ui(self) -> None:
        style = ttk.Style()
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Section.TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Primary.TButton", padding=(18, 8))
        style.configure("Action.TButton", padding=(18, 8))
        style.configure("Stat.TLabel", font=("Microsoft YaHei UI", 16, "bold"))

        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        header = ttk.Frame(frame)
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Excel-Dr", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="修复卡顿 Excel 报表。检测阶段只读取文件，清理阶段只新建副本，不覆盖原文件。",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        picker = ttk.LabelFrame(frame, text="1 选择处理对象", style="Section.TLabelframe")
        picker.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        picker.columnconfigure(1, weight=1)

        ttk.Label(picker, text="处理模式").grid(row=0, column=0, sticky="w", padx=14, pady=(14, 8))
        mode_bar = ttk.Frame(picker)
        mode_bar.grid(row=0, column=1, sticky="w", padx=8, pady=(14, 8))
        self.file_radio = ttk.Radiobutton(mode_bar, text="单文件", value="file", variable=self.mode_var, command=self.on_mode_change)
        self.file_radio.pack(side="left")
        self.folder_radio = ttk.Radiobutton(mode_bar, text="文件夹批量", value="folder", variable=self.mode_var, command=self.on_mode_change)
        self.folder_radio.pack(side="left", padx=(18, 0))

        self.source_label = ttk.Label(picker, text="源文件")
        self.source_label.grid(row=1, column=0, sticky="w", padx=14, pady=8)
        self.source_entry = ttk.Entry(picker, textvariable=self.source_var)
        self.source_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=8, ipady=4)
        self.pick_source_button = ttk.Button(picker, text="选择文件", command=self.pick_source, style="Action.TButton")
        self.pick_source_button.grid(row=1, column=2, sticky="ew", padx=(8, 14), pady=8)

        self.output_label = ttk.Label(picker, text="输出文件")
        self.output_label.grid(row=2, column=0, sticky="w", padx=14, pady=(8, 14))
        self.output_entry = ttk.Entry(picker, textvariable=self.output_var)
        self.output_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=(8, 14), ipady=4)
        self.output_button = ttk.Button(picker, text="另存为", command=self.pick_output, style="Action.TButton")
        self.output_button.grid(row=2, column=2, sticky="ew", padx=(8, 14), pady=(8, 14))

        progress = ttk.LabelFrame(frame, text="2 执行处理", style="Section.TLabelframe")
        progress.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        progress.columnconfigure(0, weight=1)

        status_row = ttk.Frame(progress)
        status_row.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 6))
        status_row.columnconfigure(0, weight=1)
        ttk.Label(status_row, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        ttk.Label(status_row, textvariable=self.percent_var).grid(row=0, column=1, sticky="e")

        self.progress_bar = ttk.Progressbar(progress, variable=self.progress_var, maximum=100, mode="determinate")
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 12))

        action_row = ttk.Frame(progress)
        action_row.grid(row=2, column=0, sticky="w", padx=14, pady=(0, 14))
        self.scan_button = ttk.Button(action_row, text="仅检测", command=self.scan_action, style="Action.TButton")
        self.scan_button.pack(side="left")
        self.clean_button = ttk.Button(action_row, text="清理并新建", command=self.clean_action, style="Primary.TButton")
        self.clean_button.pack(side="left", padx=(10, 0))

        result = ttk.LabelFrame(frame, text="3 查看结果", style="Section.TLabelframe")
        result.grid(row=3, column=0, sticky="nsew", pady=(14, 0))
        result.columnconfigure(0, weight=1)
        result.rowconfigure(1, weight=1)

        stats = ttk.Frame(result)
        stats.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 10))
        for col in range(4):
            stats.columnconfigure(col, weight=1)
        self.create_stat(stats, 0, "扫描文件", self.stat_file_var)
        self.create_stat(stats, 1, "可清理对象", self.stat_object_var)
        self.create_stat(stats, 2, "损坏规则", self.stat_rule_var)
        self.create_stat(stats, 3, "输出副本", self.stat_output_var)

        self.log = scrolledtext.ScrolledText(result, wrap="word", font=("Consolas", 10), height=12)
        self.log.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        self.set_log(
            "等待操作。\n\n"
            "建议流程：\n"
            "1. 选择 .xlsx 文件或文件夹。\n"
            "2. 点击“仅检测”查看问题。\n"
            "3. 确认后点击“清理并新建”，软件会输出新副本。"
        )

        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)
        self.on_mode_change()

    def create_stat(self, parent: ttk.Frame, column: int, title: str, variable: tk.StringVar) -> None:
        box = ttk.Frame(parent, padding=(12, 10))
        box.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
        ttk.Label(box, text=title).grid(row=0, column=0, sticky="w")
        ttk.Label(box, textvariable=variable, style="Stat.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))

    def on_mode_change(self) -> None:
        if self.worker_active:
            return
        folder_mode = self.mode_var.get() == "folder"
        self.source_var.set("")
        self.output_var.set("")
        self.source_label.configure(text="源文件夹" if folder_mode else "源文件")
        self.pick_source_button.configure(text="选择文件夹" if folder_mode else "选择文件")
        if folder_mode:
            self.output_label.configure(text="输出规则")
            self.output_var.set("每个异常文件会在原目录生成 *_cleaned.xlsx")
            self.output_entry.configure(state="disabled")
            self.output_button.configure(state="disabled")
        else:
            self.output_label.configure(text="输出文件")
            self.toggle_output_mode()

    def default_output(self, source: Path) -> Path:
        return source.with_name(f"{source.stem}_cleaned{source.suffix}")

    def toggle_output_mode(self) -> None:
        if self.mode_var.get() == "folder":
            self.output_entry.configure(state="disabled")
            self.output_button.configure(state="disabled")
            return
        self.output_entry.configure(state="disabled")
        self.output_button.configure(state="disabled")

    def pick_source(self) -> None:
        if self.worker_active:
            return
        if self.mode_var.get() == "folder":
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.source_var.set(path)
            if self.mode_var.get() == "file":
                self.output_var.set(str(self.default_output(Path(path))))
            self.reset_result()

    def pick_output(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.output_var.set(path)

    def validate_single_paths(self, needs_output: bool) -> Tuple[Optional[Path], Optional[Path]]:
        source_text = self.source_var.get().strip()
        if not source_text:
            messagebox.showerror("缺少文件", "请选择 Excel 文件")
            return None, None
        source = Path(source_text)
        if not source.exists():
            messagebox.showerror("文件不存在", f"找不到文件:\n{source}")
            return None, None
        if source.suffix.lower() != ".xlsx":
            messagebox.showerror("格式不支持", "当前版本只支持 .xlsx")
            return None, None

        output = None
        if needs_output:
            output = self.default_output(source)
            if output.resolve() == source.resolve():
                messagebox.showerror("输出非法", "输出文件不能覆盖原文件")
                return None, None
            output.parent.mkdir(parents=True, exist_ok=True)
            self.output_var.set(str(output))
        return source, output

    def validate_folder(self) -> Optional[Path]:
        source_text = self.source_var.get().strip()
        if not source_text:
            messagebox.showerror("缺少文件夹", "请选择文件夹")
            return None
        folder = Path(source_text)
        if not folder.exists() or not folder.is_dir():
            messagebox.showerror("文件夹不存在", f"找不到文件夹:\n{folder}")
            return None
        return folder

    def reset_result(self) -> None:
        self.progress_var.set(0)
        self.percent_var.set("0%")
        self.status_var.set("已选择文件，建议先检测")
        self.stat_file_var.set("-")
        self.stat_object_var.set("-")
        self.stat_rule_var.set("-")
        self.stat_output_var.set("-")

    def set_log(self, text: str) -> None:
        self.log.delete("1.0", tk.END)
        self.log.insert(tk.END, text)
        self.log.see(tk.END)

    def append_log(self, text: str) -> None:
        self.log.insert(tk.END, f"\n{text}")
        self.log.see(tk.END)

    def set_busy(self, busy: bool) -> None:
        self.worker_active = busy
        state = "disabled" if busy else "normal"
        for widget in (
            self.file_radio,
            self.folder_radio,
            self.source_entry,
            self.pick_source_button,
            self.scan_button,
            self.clean_button,
        ):
            widget.configure(state=state)

    def progress_callback(self, current: int, total: int, message: str) -> None:
        self.task_queue.put(("progress", current, total, message))

    def start_task(self, task_name: str, worker: Callable[[], object]) -> None:
        if self.worker_active:
            return
        self.set_busy(True)
        self.progress_var.set(0)
        self.percent_var.set("0%")
        self.status_var.set("正在准备处理")

        def run() -> None:
            try:
                result = worker()
                self.task_queue.put(("done", task_name, result))
            except Exception as exc:
                self.task_queue.put(("error", task_name, str(exc)))

        threading.Thread(target=run, daemon=True).start()

    def process_task_queue(self) -> None:
        try:
            while True:
                event = self.task_queue.get_nowait()
                if event[0] == "progress":
                    _, current, total, message = event
                    percent = min(100, int(current * 100 / max(1, total)))
                    self.progress_var.set(percent)
                    self.percent_var.set(f"{percent}%")
                    self.status_var.set(message)
                elif event[0] == "done":
                    _, task_name, result = event
                    self.set_busy(False)
                    self.progress_var.set(100)
                    self.percent_var.set("100%")
                    self.handle_task_done(task_name, result)
                elif event[0] == "error":
                    _, task_name, message = event
                    self.set_busy(False)
                    self.status_var.set("处理失败")
                    messagebox.showerror("处理失败", message)
        except queue.Empty:
            pass
        self.root.after(100, self.process_task_queue)

    def update_stats(self, result: object) -> None:
        if isinstance(result, WorkbookReport):
            self.stat_file_var.set("1")
            self.stat_object_var.set(str(result.suspicious_total))
            self.stat_rule_var.set(str(result.broken_validation_total))
            self.stat_output_var.set("1" if result.output else "-")
        elif isinstance(result, BatchResult):
            self.stat_file_var.set(str(result.file_count))
            self.stat_object_var.set(str(result.suspicious_total))
            self.stat_rule_var.set(str(result.broken_validation_total))
            self.stat_output_var.set(str(result.actionable_count))

    def handle_task_done(self, task_name: str, result: object) -> None:
        self.last_result = result
        if isinstance(result, WorkbookReport):
            self.update_stats(result)
            self.set_log(result.render())
        elif isinstance(result, BatchResult):
            self.update_stats(result)
            self.set_log(result.render())

        if task_name == "scan":
            self.status_var.set("检测完成")
        elif task_name == "preview_single_clean":
            report, output = result  # type: ignore[misc]
            report.output = output
            self.update_stats(report)
            self.set_log(report.render())
            if self.confirm_single_clean(report):
                self.start_task("clean_single", lambda: clean_workbook_from_report(report, output, self.progress_callback))
            else:
                self.status_var.set("已取消清理")
        elif task_name == "preview_folder_clean":
            preview = result
            if self.confirm_batch_clean(preview):  # type: ignore[arg-type]
                self.start_task("clean_folder", lambda: clean_folder_from_preview(preview, self.progress_callback))  # type: ignore[arg-type]
            else:
                self.status_var.set("已取消批量清理")
        elif task_name == "clean_single":
            self.status_var.set("清理完成，已新建副本")
            messagebox.showinfo("完成", f"清理完成。\n输出文件：\n{result.output}")  # type: ignore[union-attr]
        elif task_name == "clean_folder":
            self.status_var.set("批量清理完成")
            messagebox.showinfo("完成", f"批量清理完成，已处理 {result.actionable_count} 个文件。")  # type: ignore[union-attr]

    def confirm_single_clean(self, report: WorkbookReport) -> bool:
        if not report.needs_cleanup:
            messagebox.showinfo("无需清理", "这个文件没有发现需要清理的问题。")
            return False
        message = (
            f"将清理 {report.suspicious_total} 个异常隐藏对象"
            f"{'，并修复 ' + str(report.broken_validation_total) + ' 处数据有效性' if report.broken_validation_total else ''}。\n\n"
            f"输出文件：\n{report.output}\n\n"
            "原文件不会被覆盖。继续吗？"
        )
        return messagebox.askyesno("确认清理并新建", message)

    def confirm_batch_clean(self, result: BatchResult) -> bool:
        if result.actionable_count == 0:
            messagebox.showinfo("无需清理", "这个文件夹里的文件都正常，没有需要清理的内容。")
            return False
        message = (
            f"将处理 {result.actionable_count} 个文件。\n"
            f"预计清理 {result.suspicious_total} 个异常隐藏对象"
            f"{'，并修复 ' + str(result.broken_validation_total) + ' 处数据有效性' if result.broken_validation_total else ''}。\n\n"
            "正常文件会自动跳过，异常文件会在原目录新建副本。继续吗？"
        )
        return messagebox.askyesno("确认批量清理并新建", message)

    def scan_action(self) -> None:
        if self.mode_var.get() == "folder":
            folder = self.validate_folder()
            if not folder:
                return
            self.start_task("scan", lambda: analyze_folder(folder, self.progress_callback))
            return

        source, _ = self.validate_single_paths(needs_output=False)
        if not source:
            return
        self.start_task("scan", lambda: analyze_workbook(source, self.progress_callback))

    def clean_action(self) -> None:
        if self.mode_var.get() == "folder":
            folder = self.validate_folder()
            if not folder:
                return
            self.start_task("preview_folder_clean", lambda: analyze_folder(folder, self.progress_callback))
            return

        source, output = self.validate_single_paths(needs_output=True)
        if not source or not output:
            return
        self.start_task(
            "preview_single_clean",
            lambda: (analyze_workbook(source, self.progress_callback), output),
        )


def main() -> int:
    parser = argparse.ArgumentParser(description="Excel Dr")
    parser.add_argument("--scan", metavar="XLSX", help="scan workbook and print report")
    parser.add_argument("--clean", metavar="XLSX", help="clean workbook and write output")
    parser.add_argument("--output", metavar="XLSX", help="output path for --clean")
    parser.add_argument("--scan-folder", metavar="DIR", help="scan all xlsx files under folder")
    parser.add_argument("--clean-folder", metavar="DIR", help="clean all xlsx files under folder")
    args = parser.parse_args()

    if args.scan:
        print(analyze_workbook(Path(args.scan)).render())
        return 0

    if args.clean:
        source = Path(args.clean)
        output = Path(args.output) if args.output else source.with_name(f"{source.stem}_cleaned{source.suffix}")
        print(clean_workbook(source, output).render())
        return 0

    if args.scan_folder:
        print(analyze_folder(Path(args.scan_folder)).render())
        return 0

    if args.clean_folder:
        print(clean_folder(Path(args.clean_folder)).render())
        return 0

    root = tk.Tk()
    CleanerApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
