use anyhow::{anyhow, bail, Context, Result};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, HashMap};
use std::fs::{self, File};
use std::io::{Cursor, Read, Write};
use std::path::{Component, Path, PathBuf};
use tempfile::TempDir;
use walkdir::WalkDir;
use xmltree::{Element, XMLNode};
use zip::{CompressionMethod, ZipArchive, ZipWriter};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const XDR_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct AnchorSignature {
    pub start_col: Option<i32>,
    pub start_row: Option<i32>,
    pub embed_id: Option<String>,
    pub image_path: Option<String>,
    pub descr: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AnchorInfo {
    pub start_col: Option<i32>,
    pub start_row: Option<i32>,
    pub hidden: bool,
    pub embed_id: Option<String>,
    pub image_path: Option<String>,
    pub descr: String,
}

impl AnchorInfo {
    fn start_key(&self) -> (Option<i32>, Option<i32>) {
        (self.start_col, self.start_row)
    }

    fn signature(&self) -> AnchorSignature {
        AnchorSignature {
            start_col: self.start_col,
            start_row: self.start_row,
            embed_id: self.embed_id.clone(),
            image_path: self.image_path.clone(),
            descr: self.descr.clone(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DrawingPlan {
    pub rel_id: String,
    pub drawing_path: String,
    pub anchor_total: usize,
    pub hidden_total: usize,
    #[serde(skip_serializing, skip_deserializing)]
    pub suspicious_signatures: BTreeMap<AnchorSignature, usize>,
    pub suspicious_total: usize,
    pub remove_entire_drawing: bool,
    pub summary: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetPlan {
    pub sheet_name: String,
    pub sheet_path: String,
    pub drawing_plans: Vec<DrawingPlan>,
    pub broken_validations: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookReport {
    pub source: PathBuf,
    pub output: Option<PathBuf>,
    pub sheet_plans: Vec<SheetPlan>,
    pub findings: Vec<String>,
    pub warnings: Vec<String>,
    pub removed_anchors: usize,
    pub removed_drawings: usize,
    pub removed_validations: usize,
}

impl WorkbookReport {
    pub fn new(source: impl Into<PathBuf>) -> Self {
        Self {
            source: source.into(),
            output: None,
            sheet_plans: Vec::new(),
            findings: Vec::new(),
            warnings: Vec::new(),
            removed_anchors: 0,
            removed_drawings: 0,
            removed_validations: 0,
        }
    }

    pub fn suspicious_total(&self) -> usize {
        self.sheet_plans
            .iter()
            .flat_map(|sheet| &sheet.drawing_plans)
            .map(|drawing| drawing.suspicious_total)
            .sum()
    }

    pub fn broken_validation_total(&self) -> usize {
        self.sheet_plans
            .iter()
            .map(|sheet| sheet.broken_validations.len())
            .sum()
    }

    pub fn needs_cleanup(&self) -> bool {
        self.suspicious_total() > 0 || self.broken_validation_total() > 0
    }

    pub fn render(&self) -> String {
        let mut lines = vec![format!("文件: {}", self.source.display())];
        if let Some(output) = &self.output {
            lines.push(format!("输出: {}", output.display()));
        }
        lines.push(String::new());
        lines.push("检测结果:".to_string());
        if self.findings.is_empty() {
            lines.push("- 未发现需要清理的问题".to_string());
        } else {
            lines.extend(self.findings.iter().map(|line| format!("- {line}")));
        }

        if !self.sheet_plans.is_empty() {
            lines.push(String::new());
            lines.push("明细:".to_string());
            for sheet in &self.sheet_plans {
                lines.push(format!("- 工作表 {}", sheet.sheet_name));
                for drawing in &sheet.drawing_plans {
                    lines.push(format!(
                        "  绘图关系 {}: 总对象 {}，隐藏 {}，计划清理 {}",
                        drawing.rel_id,
                        drawing.anchor_total,
                        drawing.hidden_total,
                        drawing.suspicious_total
                    ));
                    lines.extend(drawing.summary.iter().map(|item| format!("    {item}")));
                }
                lines.extend(
                    sheet
                        .broken_validations
                        .iter()
                        .map(|item| format!("  损坏数据有效性: {item}")),
                );
            }
        }

        if !self.warnings.is_empty() {
            lines.push(String::new());
            lines.push("提示:".to_string());
            lines.extend(self.warnings.iter().map(|line| format!("- {line}")));
        }

        if self.output.is_some() {
            lines.push(String::new());
            lines.push("清理结果:".to_string());
            lines.push(format!("- 删除异常隐藏对象: {}", self.removed_anchors));
            lines.push(format!("- 移除空绘图关系: {}", self.removed_drawings));
            lines.push(format!(
                "- 移除损坏数据有效性: {}",
                self.removed_validations
            ));
        }

        lines.join("\n")
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct BatchResult {
    pub reports: Vec<WorkbookReport>,
    pub skipped: Vec<(PathBuf, String)>,
    pub failed: Vec<(PathBuf, String)>,
}

impl BatchResult {
    pub fn file_count(&self) -> usize {
        self.reports.len() + self.failed.len()
    }

    pub fn actionable_count(&self) -> usize {
        self.reports
            .iter()
            .filter(|report| report.needs_cleanup())
            .count()
    }

    pub fn suspicious_total(&self) -> usize {
        self.reports
            .iter()
            .map(WorkbookReport::suspicious_total)
            .sum()
    }

    pub fn broken_validation_total(&self) -> usize {
        self.reports
            .iter()
            .map(WorkbookReport::broken_validation_total)
            .sum()
    }

    pub fn render(&self) -> String {
        let mut lines = vec![
            format!("扫描文件数: {}", self.file_count()),
            format!("需要处理的文件: {}", self.actionable_count()),
            format!("异常隐藏对象: {}", self.suspicious_total()),
            format!("损坏数据有效性: {}", self.broken_validation_total()),
        ];
        if !self.skipped.is_empty() {
            lines.push(String::new());
            lines.push("跳过:".to_string());
            lines.extend(
                self.skipped
                    .iter()
                    .map(|(path, reason)| format!("- {}: {reason}", file_name(path))),
            );
        }
        if !self.failed.is_empty() {
            lines.push(String::new());
            lines.push("失败:".to_string());
            lines.extend(
                self.failed
                    .iter()
                    .map(|(path, reason)| format!("- {}: {reason}", file_name(path))),
            );
        }
        if !self.reports.is_empty() {
            lines.push(String::new());
            lines.push("文件明细:".to_string());
            for report in &self.reports {
                let status = if report.needs_cleanup() {
                    "需要清理"
                } else {
                    "正常"
                };
                lines.push(format!(
                    "- {}: {status}，异常对象 {}，损坏有效性 {}",
                    file_name(&report.source),
                    report.suspicious_total(),
                    report.broken_validation_total()
                ));
            }
        }
        lines.join("\n")
    }
}

pub fn analyze_file(path: impl AsRef<Path>) -> Result<WorkbookReport> {
    let path = path.as_ref();
    validate_xlsx(path)?;
    let mut report = WorkbookReport::new(path);
    let mut zip = open_zip(path)?;
    let names = zip_names(&mut zip);
    if !names.iter().any(|name| name == "xl/workbook.xml") {
        bail!("MissingWorkbook: 文件缺少 xl/workbook.xml");
    }

    let workbook_root = parse_zip_xml(&mut zip, "xl/workbook.xml")?;
    let workbook_rels = read_relationships(&mut zip, "xl/_rels/workbook.xml.rels")?;

    for sheet_node in find_descendants(&workbook_root, MAIN_NS, "sheet") {
        let sheet_name = sheet_node
            .attributes
            .get("name")
            .cloned()
            .unwrap_or_default();
        let sheet_rid = attr_by_local(sheet_node, "id").unwrap_or_default();
        let sheet_path = match workbook_rels.get(&sheet_rid) {
            Some(path) => path.clone(),
            None => continue,
        };
        if !names.contains(&sheet_path) {
            continue;
        }

        let mut sheet_plan = SheetPlan {
            sheet_name: sheet_name.clone(),
            sheet_path: sheet_path.clone(),
            drawing_plans: Vec::new(),
            broken_validations: Vec::new(),
        };
        let sheet_root = parse_zip_xml(&mut zip, &sheet_path)?;
        let sheet_rel_path = rels_path_for_part(&sheet_path);
        let sheet_rels = read_relationships(&mut zip, &sheet_rel_path)?;

        for drawing_node in find_children(&sheet_root, MAIN_NS, "drawing") {
            let rel_id = attr_by_local(drawing_node, "id").unwrap_or_default();
            let drawing_path = match sheet_rels.get(&rel_id) {
                Some(path) if names.contains(path) => path.clone(),
                _ => continue,
            };
            let drawing_root = parse_zip_xml(&mut zip, &drawing_path)?;
            let drawing_rel_path = rels_path_for_part(&drawing_path);
            let drawing_rels = read_relationships(&mut zip, &drawing_rel_path)?;
            let anchors: Vec<_> = drawing_root
                .children
                .iter()
                .filter_map(as_element)
                .filter(|element| local_name_element(element).ends_with("Anchor"))
                .map(|anchor| parse_anchor(anchor, &drawing_rels))
                .collect();
            if anchors.is_empty() {
                continue;
            }

            let suspicious = detect_suspicious_signatures(&anchors);
            let suspicious_total = suspicious.values().sum();
            let drawing_plan = DrawingPlan {
                rel_id: rel_id.clone(),
                drawing_path: drawing_path.clone(),
                anchor_total: anchors.len(),
                hidden_total: anchors.iter().filter(|anchor| anchor.hidden).count(),
                suspicious_signatures: suspicious.clone(),
                suspicious_total,
                remove_entire_drawing: suspicious_total == anchors.len(),
                summary: summarize_signature_counts(&suspicious),
            };
            if drawing_plan.suspicious_total > 0 {
                report.findings.push(format!(
                    "{} 存在异常隐藏绘图对象 {} 个，来源关系 {}",
                    sheet_name, drawing_plan.suspicious_total, rel_id
                ));
            }
            sheet_plan.drawing_plans.push(drawing_plan);
        }

        for dv in find_descendants(&sheet_root, MAIN_NS, "dataValidation") {
            let formula1 = dv
                .attributes
                .get("formula1")
                .cloned()
                .or_else(|| child_text(dv, MAIN_NS, "formula1"))
                .unwrap_or_default();
            if formula1.contains("#REF!") {
                let sqref = dv.attributes.get("sqref").cloned().unwrap_or_default();
                sheet_plan
                    .broken_validations
                    .push(format!("{sqref} -> {formula1}"));
            }
        }

        if !sheet_plan.broken_validations.is_empty() {
            report.findings.push(format!(
                "{} 存在损坏的数据有效性 {} 处",
                sheet_name,
                sheet_plan.broken_validations.len()
            ));
        }

        if !sheet_plan.drawing_plans.is_empty() || !sheet_plan.broken_validations.is_empty() {
            report.sheet_plans.push(sheet_plan);
        }
    }

    if report.findings.is_empty() {
        report
            .warnings
            .push("未命中异常阈值，默认不会做任何删除".to_string());
    }
    report
        .warnings
        .push("本工具不会修改单元格公式、共享字符串、样式或 cellImages 图片公式".to_string());
    report
        .warnings
        .push("默认仅清理被判定为异常膨胀的隐藏 drawing 对象和损坏数据有效性".to_string());
    Ok(report)
}

pub fn clean_file(source: impl AsRef<Path>, output: impl AsRef<Path>) -> Result<WorkbookReport> {
    let report = analyze_file(source.as_ref())?;
    clean_file_from_report(report, output)
}

pub fn clean_file_from_report(
    mut report: WorkbookReport,
    output: impl AsRef<Path>,
) -> Result<WorkbookReport> {
    if !report.needs_cleanup() {
        report
            .warnings
            .push("文件未发现需要清理的问题，已跳过生成新文件".to_string());
        return Ok(report);
    }

    let output = output.as_ref();
    if let Some(parent) = output.parent() {
        fs::create_dir_all(parent)
            .with_context(|| format!("创建输出目录失败: {}", parent.display()))?;
    }
    if same_path(&report.source, output) {
        bail!("OutputExistsLocked: 输出路径不能覆盖原文件");
    }

    let source = report.source.clone();
    let mut zip = open_zip(&source)?;
    let temp = TempDir::new().context("创建临时目录失败")?;
    extract_zip(&mut zip, temp.path())?;

    for sheet_plan in report.sheet_plans.clone() {
        let sheet_file = temp.path().join(path_from_zip(&sheet_plan.sheet_path));
        if !sheet_file.exists() {
            continue;
        }
        let mut sheet_root = parse_file_xml(&sheet_file)?;
        let sheet_rel_zip_path = rels_path_for_part(&sheet_plan.sheet_path);
        let sheet_rel_file = temp.path().join(path_from_zip(&sheet_rel_zip_path));
        let mut sheet_rel_root = if sheet_rel_file.exists() {
            Some(parse_file_xml(&sheet_rel_file)?)
        } else {
            None
        };

        for drawing_plan in sheet_plan
            .drawing_plans
            .iter()
            .filter(|drawing| drawing.suspicious_total > 0)
        {
            let drawing_file = temp.path().join(path_from_zip(&drawing_plan.drawing_path));
            if !drawing_file.exists() {
                continue;
            }
            let drawing_rel_zip_path = rels_path_for_part(&drawing_plan.drawing_path);
            let drawing_rel_file = temp.path().join(path_from_zip(&drawing_rel_zip_path));
            let mut drawing_rel_root = if drawing_rel_file.exists() {
                Some(parse_file_xml(&drawing_rel_file)?)
            } else {
                None
            };
            let drawing_rels = if let Some(root) = &drawing_rel_root {
                relationships_from_root(root, &drawing_plan.drawing_path)
            } else {
                HashMap::new()
            };

            let mut drawing_root = parse_file_xml(&drawing_file)?;
            let remaining_by_rel = remove_suspicious_anchors_from_drawing(
                &mut drawing_root,
                &drawing_rels,
                &drawing_plan.suspicious_signatures,
            );
            write_xml_file(&drawing_file, &drawing_root)?;
            report.removed_anchors += drawing_plan.suspicious_total;

            if let Some(root) = &mut drawing_rel_root {
                let rel_ids: Vec<_> = root
                    .children
                    .iter()
                    .filter_map(as_element)
                    .filter(|rel| {
                        element_is(rel, REL_NS, "Relationship")
                            || local_name_element(rel) == "Relationship"
                    })
                    .filter_map(|rel| rel.attributes.get("Id").cloned())
                    .collect();
                for rel_id in rel_ids {
                    if remaining_by_rel.get(&rel_id).copied().unwrap_or_default() == 0 {
                        if let Some(target) = relationship_target(root, &rel_id) {
                            remove_relationship(root, &rel_id);
                            let media_file = temp.path().join(path_from_zip(&resolve_zip_target(
                                &drawing_plan.drawing_path,
                                &target,
                            )));
                            if media_file.exists() {
                                let _ = fs::remove_file(media_file);
                            }
                        }
                    }
                }
                write_xml_file(&drawing_rel_file, root)?;
            }

            let remaining_anchor_count = drawing_root
                .children
                .iter()
                .filter_map(as_element)
                .filter(|child| local_name_element(child).ends_with("Anchor"))
                .count();
            if remaining_anchor_count == 0 || drawing_plan.remove_entire_drawing {
                remove_sheet_drawing(&mut sheet_root, &drawing_plan.rel_id);
                if let Some(root) = &mut sheet_rel_root {
                    remove_relationship(root, &drawing_plan.rel_id);
                }
                let _ = fs::remove_file(&drawing_file);
                let _ = fs::remove_file(&drawing_rel_file);
                report.removed_drawings += 1;
            }
        }

        report.removed_validations += remove_broken_validations(&mut sheet_root);
        write_xml_file(&sheet_file, &sheet_root)?;
        if let Some(root) = &sheet_rel_root {
            write_xml_file(&sheet_rel_file, root)?;
        }
    }

    let final_output = unique_output_path(output);
    write_zip_from_dir(temp.path(), &final_output)?;
    report.output = Some(final_output);
    Ok(report)
}

pub fn analyze_folder(folder: impl AsRef<Path>) -> Result<BatchResult> {
    validate_folder(folder.as_ref())?;
    let mut result = BatchResult::default();
    for path in collect_xlsx_files(folder.as_ref()) {
        match analyze_file(&path) {
            Ok(report) => result.reports.push(report),
            Err(error) => result.failed.push((path, error.to_string())),
        }
    }
    Ok(result)
}

pub fn clean_folder(folder: impl AsRef<Path>) -> Result<BatchResult> {
    validate_folder(folder.as_ref())?;
    let preview = analyze_folder(folder)?;
    clean_folder_from_preview(preview)
}

pub fn clean_folder_from_preview(preview: BatchResult) -> Result<BatchResult> {
    let mut result = BatchResult::default();
    result.failed.extend(preview.failed);
    for report in preview.reports {
        let source = report.source.clone();
        if !report.needs_cleanup() {
            result
                .skipped
                .push((report.source.clone(), "正常，无需清理".to_string()));
            result.reports.push(report);
            continue;
        }
        let output = default_output_path(&report.source);
        match clean_file_from_report(report, output) {
            Ok(cleaned) => result.reports.push(cleaned),
            Err(error) => result.failed.push((source, error.to_string())),
        }
    }
    Ok(result)
}

pub fn collect_xlsx_files(folder: &Path) -> Vec<PathBuf> {
    if !folder.is_dir() {
        return Vec::new();
    }
    let mut files: Vec<_> = WalkDir::new(folder)
        .into_iter()
        .filter_map(Result::ok)
        .filter(|entry| entry.file_type().is_file())
        .map(|entry| entry.into_path())
        .filter(|path| {
            path.extension()
                .map(|ext| ext.eq_ignore_ascii_case("xlsx"))
                .unwrap_or(false)
        })
        .filter(|path| {
            let name = file_name(path);
            let stem = path
                .file_stem()
                .and_then(|item| item.to_str())
                .unwrap_or_default();
            !name.starts_with("~$") && !stem.ends_with("_cleaned") && !stem.contains("_cleaned_")
        })
        .collect();
    files.sort();
    files
}

pub fn default_output_path(source: &Path) -> PathBuf {
    let stem = source
        .file_stem()
        .and_then(|item| item.to_str())
        .unwrap_or("output");
    source.with_file_name(format!("{stem}_cleaned.xlsx"))
}

fn validate_xlsx(path: &Path) -> Result<()> {
    if !path.exists() {
        bail!("FileNotFound: 找不到选择的文件");
    }
    if !path.is_file() {
        bail!("UnsupportedFormat: 请选择 .xlsx 文件");
    }
    if !path
        .extension()
        .map(|ext| ext.eq_ignore_ascii_case("xlsx"))
        .unwrap_or(false)
    {
        bail!("UnsupportedFormat: 当前版本只支持 .xlsx 文件");
    }
    Ok(())
}

fn validate_folder(path: &Path) -> Result<()> {
    if !path.exists() {
        bail!("FolderNotFound: 找不到选择的文件夹");
    }
    if !path.is_dir() {
        bail!("UnsupportedTarget: 请选择文件夹");
    }
    Ok(())
}

fn open_zip(path: &Path) -> Result<ZipArchive<File>> {
    let file = File::open(path)
        .with_context(|| format!("PermissionDenied: 无法读取 {}", path.display()))?;
    ZipArchive::new(file).map_err(|error| anyhow!("InvalidZip: {error}"))
}

fn zip_names(zip: &mut ZipArchive<File>) -> Vec<String> {
    (0..zip.len())
        .filter_map(|index| {
            zip.by_index(index)
                .ok()
                .map(|file| file.name().replace('\\', "/"))
        })
        .collect()
}

fn parse_zip_xml(zip: &mut ZipArchive<File>, name: &str) -> Result<Element> {
    let mut file = zip
        .by_name(name)
        .with_context(|| format!("MissingPart: 缺少 {name}"))?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    Element::parse(Cursor::new(bytes)).map_err(|error| anyhow!("XmlParseFailed: {name}: {error}"))
}

fn parse_file_xml(path: &Path) -> Result<Element> {
    let bytes = fs::read(path).with_context(|| format!("读取 XML 失败: {}", path.display()))?;
    Element::parse(Cursor::new(bytes))
        .map_err(|error| anyhow!("XmlParseFailed: {}: {error}", path.display()))
}

fn write_xml_file(path: &Path, element: &Element) -> Result<()> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    let mut file =
        File::create(path).with_context(|| format!("写入 XML 失败: {}", path.display()))?;
    element.write_with_config(
        &mut file,
        xmltree::EmitterConfig::new()
            .perform_indent(false)
            .write_document_declaration(true),
    )?;
    Ok(())
}

fn read_relationships(
    zip: &mut ZipArchive<File>,
    rel_path: &str,
) -> Result<HashMap<String, String>> {
    match parse_zip_xml(zip, rel_path) {
        Ok(root) => Ok(relationships_from_root(
            &root,
            &source_part_from_rels(rel_path),
        )),
        Err(_) => Ok(HashMap::new()),
    }
}

fn relationships_from_root(root: &Element, owner_part: &str) -> HashMap<String, String> {
    root.children
        .iter()
        .filter_map(as_element)
        .filter(|rel| {
            element_is(rel, REL_NS, "Relationship") || local_name_element(rel) == "Relationship"
        })
        .filter_map(|rel| {
            let id = rel.attributes.get("Id")?.clone();
            let target = rel.attributes.get("Target")?.clone();
            Some((id, resolve_zip_target(owner_part, &target)))
        })
        .collect()
}

fn source_part_from_rels(rel_path: &str) -> String {
    let rel_path = rel_path.replace('\\', "/");
    if !rel_path.contains("/_rels/") || !rel_path.ends_with(".rels") {
        return rel_path;
    }
    let (left, right) = rel_path.split_once("/_rels/").unwrap();
    let owner_name = right.trim_end_matches(".rels");
    if left.is_empty() {
        owner_name.to_string()
    } else {
        format!("{left}/{owner_name}")
    }
}

pub fn resolve_zip_target(base_part: &str, target: &str) -> String {
    let target = target.replace('\\', "/");
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }
    let base_dir = Path::new(base_part)
        .parent()
        .unwrap_or_else(|| Path::new(""));
    let target_path = base_dir.join(target);
    normalize_zip_path(&target_path)
}

fn normalize_zip_path(path: &Path) -> String {
    let mut parts = Vec::new();
    for component in path.components() {
        match component {
            Component::Normal(part) => parts.push(part.to_string_lossy().to_string()),
            Component::ParentDir => {
                parts.pop();
            }
            Component::CurDir => {}
            _ => {}
        }
    }
    parts.join("/")
}

fn rels_path_for_part(part: &str) -> String {
    let path = Path::new(part);
    let parent = path.parent().map(normalize_zip_path).unwrap_or_default();
    let file = path
        .file_name()
        .and_then(|item| item.to_str())
        .unwrap_or_default();
    if parent.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{parent}/_rels/{file}.rels")
    }
}

fn parse_anchor(anchor: &Element, drawing_rels: &HashMap<String, String>) -> AnchorInfo {
    let from_node = find_child(anchor, XDR_NS, "from");
    let start_col = from_node
        .and_then(|node| child_text(node, XDR_NS, "col"))
        .and_then(|text| text.parse().ok());
    let start_row = from_node
        .and_then(|node| child_text(node, XDR_NS, "row"))
        .and_then(|text| text.parse().ok());
    let c_nv_pr = find_descendants(anchor, XDR_NS, "cNvPr").into_iter().next();
    let hidden = c_nv_pr
        .and_then(|node| node.attributes.get("hidden"))
        .map(|value| value == "1")
        .unwrap_or(false);
    let descr = c_nv_pr
        .and_then(|node| node.attributes.get("descr"))
        .cloned()
        .unwrap_or_default();
    let blip = find_descendants(anchor, DRAWING_NS, "blip")
        .into_iter()
        .next();
    let embed_id = blip.and_then(|node| attr_by_local(node, "embed"));
    let image_path = embed_id
        .as_ref()
        .and_then(|id| drawing_rels.get(id))
        .cloned();
    AnchorInfo {
        start_col,
        start_row,
        hidden,
        embed_id,
        image_path,
        descr,
    }
}

fn detect_suspicious_signatures(anchors: &[AnchorInfo]) -> BTreeMap<AnchorSignature, usize> {
    let hidden: Vec<_> = anchors.iter().filter(|anchor| anchor.hidden).collect();
    if hidden.len() < 100 {
        return BTreeMap::new();
    }

    let mut by_position: HashMap<(Option<i32>, Option<i32>), usize> = HashMap::new();
    let mut by_signature: BTreeMap<AnchorSignature, usize> = BTreeMap::new();
    for anchor in &hidden {
        *by_position.entry(anchor.start_key()).or_default() += 1;
        *by_signature.entry(anchor.signature()).or_default() += 1;
    }

    let mut suspicious = BTreeMap::new();
    for anchor in &hidden {
        let signature = anchor.signature();
        if by_position
            .get(&anchor.start_key())
            .copied()
            .unwrap_or_default()
            >= 50
            && by_signature.get(&signature).copied().unwrap_or_default() >= 100
        {
            *suspicious.entry(signature).or_default() += 1;
        }
    }
    if !suspicious.is_empty() {
        return suspicious;
    }

    if hidden.len() >= 5000 {
        if let Some((dominant_signature, dominant_count)) =
            by_signature.iter().max_by_key(|(_, count)| *count)
        {
            if (*dominant_count as f64 / hidden.len() as f64) >= 0.9 {
                suspicious.insert(dominant_signature.clone(), *dominant_count);
            }
        }
    }
    suspicious
}

fn summarize_signature_counts(signatures: &BTreeMap<AnchorSignature, usize>) -> Vec<String> {
    if signatures.is_empty() {
        return Vec::new();
    }
    let mut by_position: BTreeMap<(Option<i32>, Option<i32>), usize> = BTreeMap::new();
    let mut by_image: BTreeMap<String, usize> = BTreeMap::new();
    for (signature, count) in signatures {
        *by_position
            .entry((signature.start_col, signature.start_row))
            .or_default() += *count;
        *by_image
            .entry(
                signature
                    .image_path
                    .clone()
                    .unwrap_or_else(|| "未知".to_string()),
            )
            .or_default() += *count;
    }
    let mut result = Vec::new();
    let mut positions: Vec<_> = by_position.into_iter().collect();
    positions.sort_by(|left, right| right.1.cmp(&left.1));
    if !positions.is_empty() {
        result.push(format!(
            "重复锚点 {}",
            positions
                .into_iter()
                .take(3)
                .map(|((col, row), count)| format!(
                    "(col={}, row={}) x {count}",
                    opt_i32(col),
                    opt_i32(row)
                ))
                .collect::<Vec<_>>()
                .join("、")
        ));
    }
    let mut images: Vec<_> = by_image.into_iter().collect();
    images.sort_by(|left, right| right.1.cmp(&left.1));
    if !images.is_empty() {
        result.push(format!(
            "重复图片资源 {}",
            images
                .into_iter()
                .take(3)
                .map(|(image, count)| format!("{image} x {count}"))
                .collect::<Vec<_>>()
                .join("、")
        ));
    }
    result
}

fn remove_suspicious_anchors_from_drawing(
    root: &mut Element,
    drawing_rels: &HashMap<String, String>,
    suspicious_signatures: &BTreeMap<AnchorSignature, usize>,
) -> HashMap<String, usize> {
    let mut remaining_suspicious = suspicious_signatures.clone();
    let mut remaining_by_rel = HashMap::new();
    let mut new_children = Vec::with_capacity(root.children.len());
    for child in root.children.drain(..) {
        let Some(anchor) = as_element(&child) else {
            new_children.push(child);
            continue;
        };
        if !local_name_element(anchor).ends_with("Anchor") {
            new_children.push(child);
            continue;
        }
        let info = parse_anchor(anchor, drawing_rels);
        let signature = info.signature();
        if let Some(count) = remaining_suspicious.get_mut(&signature) {
            if *count > 0 {
                *count -= 1;
                continue;
            }
        }
        if let Some(embed_id) = info.embed_id {
            *remaining_by_rel.entry(embed_id).or_default() += 1;
        }
        new_children.push(child);
    }
    root.children = new_children;
    remaining_by_rel
}

fn remove_broken_validations(sheet_root: &mut Element) -> usize {
    let Some(index) = sheet_root.children.iter().position(|child| {
        as_element(child)
            .map(|element| element_is(element, MAIN_NS, "dataValidations"))
            .unwrap_or(false)
    }) else {
        return 0;
    };

    let Some(data_validations) = as_element_mut(&mut sheet_root.children[index]) else {
        return 0;
    };
    let mut removed = 0;
    data_validations.children.retain(|child| {
        let Some(dv) = as_element(child) else {
            return true;
        };
        if !element_is(dv, MAIN_NS, "dataValidation") {
            return true;
        }
        let formula1 = dv
            .attributes
            .get("formula1")
            .cloned()
            .or_else(|| child_text(dv, MAIN_NS, "formula1"))
            .unwrap_or_default();
        if formula1.contains("#REF!") {
            removed += 1;
            false
        } else {
            true
        }
    });

    if removed > 0 {
        let remaining = data_validations
            .children
            .iter()
            .filter_map(as_element)
            .filter(|element| element_is(element, MAIN_NS, "dataValidation"))
            .count();
        if remaining > 0 {
            data_validations
                .attributes
                .insert("count".to_string(), remaining.to_string());
        } else {
            sheet_root.children.remove(index);
        }
    }
    removed
}

fn extract_zip(zip: &mut ZipArchive<File>, target_dir: &Path) -> Result<()> {
    for index in 0..zip.len() {
        let mut file = zip.by_index(index)?;
        if file.name().ends_with('/') {
            continue;
        }
        let out_path = target_dir.join(path_from_zip(file.name()));
        if let Some(parent) = out_path.parent() {
            fs::create_dir_all(parent)?;
        }
        let mut output = File::create(out_path)?;
        std::io::copy(&mut file, &mut output)?;
    }
    Ok(())
}

fn write_zip_from_dir(source_dir: &Path, output: &Path) -> Result<()> {
    if output.exists() {
        fs::remove_file(output)
            .with_context(|| format!("OutputExistsLocked: 无法删除旧输出 {}", output.display()))?;
    }
    let file = File::create(output)
        .with_context(|| format!("PermissionDenied: 无法写入 {}", output.display()))?;
    let mut zip = ZipWriter::new(file);
    let options =
        zip::write::FileOptions::default().compression_method(CompressionMethod::Deflated);
    let mut files: Vec<_> = WalkDir::new(source_dir)
        .into_iter()
        .filter_map(Result::ok)
        .filter(|entry| entry.file_type().is_file())
        .map(|entry| entry.into_path())
        .collect();
    files.sort();
    for path in files {
        let name = path
            .strip_prefix(source_dir)?
            .to_string_lossy()
            .replace('\\', "/");
        zip.start_file(name, options)?;
        let bytes = fs::read(&path)?;
        zip.write_all(&bytes)?;
    }
    zip.finish()?;
    Ok(())
}

fn unique_output_path(output: &Path) -> PathBuf {
    if !output.exists() {
        return output.to_path_buf();
    }
    let parent = output.parent().unwrap_or_else(|| Path::new(""));
    let stem = output
        .file_stem()
        .and_then(|item| item.to_str())
        .unwrap_or("output");
    let ext = output
        .extension()
        .and_then(|item| item.to_str())
        .unwrap_or("xlsx");
    for index in 2.. {
        let candidate = parent.join(format!("{stem}_{index}.{ext}"));
        if !candidate.exists() {
            return candidate;
        }
    }
    unreachable!()
}

fn same_path(left: &Path, right: &Path) -> bool {
    match (left.canonicalize(), right.canonicalize()) {
        (Ok(left), Ok(right)) => left == right,
        _ => left == right,
    }
}

fn path_from_zip(path: &str) -> PathBuf {
    let mut safe = PathBuf::new();
    for part in path.replace('\\', "/").split('/') {
        if part.is_empty() || part == "." {
            continue;
        }
        if part == ".." {
            safe.pop();
            continue;
        }
        safe.push(part);
    }
    safe
}

fn find_child<'a>(element: &'a Element, ns: &str, name: &str) -> Option<&'a Element> {
    element
        .children
        .iter()
        .filter_map(as_element)
        .find(|child| element_is(child, ns, name))
}

fn find_children<'a>(element: &'a Element, ns: &str, name: &str) -> Vec<&'a Element> {
    element
        .children
        .iter()
        .filter_map(as_element)
        .filter(|child| element_is(child, ns, name))
        .collect()
}

fn find_descendants<'a>(element: &'a Element, ns: &str, name: &str) -> Vec<&'a Element> {
    let mut result = Vec::new();
    for child in element.children.iter().filter_map(as_element) {
        if element_is(child, ns, name) {
            result.push(child);
        }
        result.extend(find_descendants(child, ns, name));
    }
    result
}

fn child_text(element: &Element, ns: &str, name: &str) -> Option<String> {
    find_child(element, ns, name).and_then(|child| child.get_text().map(|text| text.to_string()))
}

fn element_is(element: &Element, ns: &str, name: &str) -> bool {
    local_name_element(element) == name
        && element
            .namespace
            .as_deref()
            .map(|value| value == ns)
            .unwrap_or(true)
}

fn local_name_element(element: &Element) -> &str {
    element
        .name
        .rsplit_once(':')
        .map(|(_, name)| name)
        .unwrap_or(&element.name)
}

fn attr_by_local(element: &Element, name: &str) -> Option<String> {
    element.attributes.iter().find_map(|(key, value)| {
        let local = key.rsplit_once(':').map(|(_, item)| item).unwrap_or(key);
        if local == name {
            Some(value.clone())
        } else {
            None
        }
    })
}

fn as_element(node: &XMLNode) -> Option<&Element> {
    match node {
        XMLNode::Element(element) => Some(element),
        _ => None,
    }
}

fn as_element_mut(node: &mut XMLNode) -> Option<&mut Element> {
    match node {
        XMLNode::Element(element) => Some(element),
        _ => None,
    }
}

fn relationship_target(root: &Element, rel_id: &str) -> Option<String> {
    root.children.iter().filter_map(as_element).find_map(|rel| {
        if rel
            .attributes
            .get("Id")
            .map(|id| id == rel_id)
            .unwrap_or(false)
        {
            rel.attributes.get("Target").cloned()
        } else {
            None
        }
    })
}

fn remove_relationship(root: &mut Element, rel_id: &str) {
    root.children.retain(|child| {
        as_element(child)
            .map(|rel| {
                rel.attributes
                    .get("Id")
                    .map(|id| id != rel_id)
                    .unwrap_or(true)
            })
            .unwrap_or(true)
    });
}

fn remove_sheet_drawing(sheet_root: &mut Element, rel_id: &str) {
    sheet_root.children.retain(|child| {
        let Some(element) = as_element(child) else {
            return true;
        };
        if !element_is(element, MAIN_NS, "drawing") {
            return true;
        }
        attr_by_local(element, "id")
            .map(|value| value != rel_id)
            .unwrap_or(true)
    });
}

fn file_name(path: &Path) -> String {
    path.file_name()
        .and_then(|item| item.to_str())
        .unwrap_or_default()
        .to_string()
}

fn opt_i32(value: Option<i32>) -> String {
    value
        .map(|item| item.to_string())
        .unwrap_or_else(|| "None".to_string())
}
