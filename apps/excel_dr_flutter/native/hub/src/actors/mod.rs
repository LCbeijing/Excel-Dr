use std::path::Path;

use crate::signals::{ResultRow, TaskProgress, TaskRequest, TaskResult};
use excel_dr_core::{self, BatchResult, WorkbookReport};
use rinf::{DartSignal, RustSignal};
use tokio::task::spawn_blocking;

pub async fn create_actors() {
    let receiver = TaskRequest::get_dart_signal_receiver();
    while let Some(signal_pack) = receiver.recv().await {
        let request = signal_pack.message;
        spawn_blocking(move || run_request(request));
    }
}

fn run_request(request: TaskRequest) {
    send_progress(&request, 3, "开始处理", "正在准备任务");
    let result = match (
        request.operation.as_str(),
        request.target_kind.as_str(),
        request.path.trim(),
    ) {
        (_, _, "") => Err("请选择一个 Excel 文件或文件夹".to_string()),
        ("analyze", "file", path) => excel_dr_core::analyze_file(path)
            .map(|report| result_from_report(&request, report, false))
            .map_err(|error| error.to_string()),
        ("clean", "file", path) => {
            let output = excel_dr_core::default_output_path(Path::new(path));
            excel_dr_core::clean_file(path, output)
                .map(|report| result_from_report(&request, report, true))
                .map_err(|error| error.to_string())
        }
        ("analyze", "folder", path) => Ok(analyze_folder_with_progress(&request, path)),
        ("clean", "folder", path) => Ok(clean_folder_with_progress(&request, path)),
        _ => Err("不支持的任务类型".to_string()),
    };

    match result {
        Ok(message) => message.send_signal_to_dart(),
        Err(error) => TaskResult {
            request_id: request.request_id,
            operation: request.operation,
            target_kind: request.target_kind,
            status: "failed".to_string(),
            message: error.clone(),
            output_path: String::new(),
            scanned_files: 0,
            actionable_files: 0,
            suspicious_objects: 0,
            broken_rules: 0,
            output_files: 0,
            skipped_files: 0,
            failed_files: 1,
            rows: Vec::new(),
            warnings: Vec::new(),
            errors: vec![error],
        }
        .send_signal_to_dart(),
    }
}

fn analyze_folder_with_progress(request: &TaskRequest, folder: &str) -> TaskResult {
    let files = excel_dr_core::collect_xlsx_files(Path::new(folder));
    let total = files.len().max(1);
    let mut batch = BatchResult::default();
    for (index, path) in files.into_iter().enumerate() {
        send_progress(
            request,
            progress_percent(index, total, 8, 86),
            &file_label(&path),
            "正在检测",
        );
        match excel_dr_core::analyze_file(&path) {
            Ok(report) => batch.reports.push(report),
            Err(error) => batch.failed.push((path, error.to_string())),
        }
    }
    send_progress(request, 96, "生成报告", "正在整理结果");
    result_from_batch(request, batch, false)
}

fn clean_folder_with_progress(request: &TaskRequest, folder: &str) -> TaskResult {
    let files = excel_dr_core::collect_xlsx_files(Path::new(folder));
    let total = files.len().max(1);
    let mut batch = BatchResult::default();
    for (index, path) in files.into_iter().enumerate() {
        send_progress(
            request,
            progress_percent(index, total, 8, 88),
            &file_label(&path),
            "正在检测并清理",
        );
        match excel_dr_core::analyze_file(&path) {
            Ok(report) if report.needs_cleanup() => {
                let output = excel_dr_core::default_output_path(&report.source);
                match excel_dr_core::clean_file_from_report(report, output) {
                    Ok(cleaned) => batch.reports.push(cleaned),
                    Err(error) => batch.failed.push((path, error.to_string())),
                }
            }
            Ok(report) => {
                batch
                    .skipped
                    .push((report.source.clone(), "正常，无需清理".to_string()));
                batch.reports.push(report);
            }
            Err(error) => batch.failed.push((path, error.to_string())),
        }
    }
    send_progress(request, 96, "生成报告", "正在整理结果");
    result_from_batch(request, batch, true)
}

fn send_progress(request: &TaskRequest, percent: i32, current: &str, detail: &str) {
    TaskProgress {
        request_id: request.request_id,
        percent,
        current: current.to_string(),
        detail: detail.to_string(),
    }
    .send_signal_to_dart();
}

fn progress_percent(index: usize, total: usize, start: i32, end: i32) -> i32 {
    let span = (end - start).max(1) as usize;
    start + ((index * span) / total) as i32
}

fn result_from_report(request: &TaskRequest, report: WorkbookReport, cleaned: bool) -> TaskResult {
    let actionable = report.needs_cleanup();
    let output_path = report
        .output
        .as_ref()
        .map(|path| path_to_string(path.as_path()))
        .unwrap_or_default();
    let rows = rows_from_report(&report, cleaned);
    let message = if cleaned && !output_path.is_empty() {
        "清理完成，已生成新文件".to_string()
    } else if actionable {
        "发现可清理内容".to_string()
    } else {
        "未发现需要清理的问题".to_string()
    };

    TaskResult {
        request_id: request.request_id,
        operation: request.operation.clone(),
        target_kind: request.target_kind.clone(),
        status: "success".to_string(),
        message,
        output_path,
        scanned_files: 1,
        actionable_files: if actionable { 1 } else { 0 },
        suspicious_objects: to_i32(report.suspicious_total()),
        broken_rules: to_i32(report.broken_validation_total()),
        output_files: if cleaned && report.output.is_some() { 1 } else { 0 },
        skipped_files: if actionable { 0 } else { 1 },
        failed_files: 0,
        rows,
        warnings: report.warnings,
        errors: Vec::new(),
    }
}

fn result_from_batch(request: &TaskRequest, batch: BatchResult, cleaned: bool) -> TaskResult {
    let output_path = first_output_path(&batch).unwrap_or_default();
    let rows = batch
        .reports
        .iter()
        .flat_map(|report| rows_from_report(report, cleaned))
        .collect::<Vec<_>>();
    let output_files = batch
        .reports
        .iter()
        .filter(|report| report.output.is_some())
        .count();
    let errors = batch
        .failed
        .iter()
        .map(|(path, reason)| format!("{}: {reason}", file_label(path)))
        .collect::<Vec<_>>();
    let message = if !errors.is_empty() {
        "部分文件处理失败，其余文件已继续完成".to_string()
    } else if cleaned && output_files > 0 {
        "批量清理完成，已生成新文件".to_string()
    } else if batch.actionable_count() > 0 {
        "发现可清理内容".to_string()
    } else {
        "未发现需要清理的问题".to_string()
    };

    TaskResult {
        request_id: request.request_id,
        operation: request.operation.clone(),
        target_kind: request.target_kind.clone(),
        status: if errors.is_empty() { "success" } else { "partial" }.to_string(),
        message,
        output_path,
        scanned_files: to_i32(batch.file_count()),
        actionable_files: to_i32(batch.actionable_count()),
        suspicious_objects: to_i32(batch.suspicious_total()),
        broken_rules: to_i32(batch.broken_validation_total()),
        output_files: to_i32(output_files),
        skipped_files: to_i32(batch.skipped.len()),
        failed_files: to_i32(batch.failed.len()),
        rows,
        warnings: Vec::new(),
        errors,
    }
}

fn rows_from_report(report: &WorkbookReport, cleaned: bool) -> Vec<ResultRow> {
    let mut rows = Vec::new();
    for sheet in &report.sheet_plans {
        for drawing in &sheet.drawing_plans {
            if drawing.suspicious_total > 0 {
                rows.push(ResultRow {
                    location: sheet.sheet_name.clone(),
                    issue: "隐藏对象过多".to_string(),
                    count: to_i32(drawing.suspicious_total),
                    action: if cleaned { "已清理" } else { "可清理" }.to_string(),
                });
            }
        }
        if !sheet.broken_validations.is_empty() {
            rows.push(ResultRow {
                location: sheet.sheet_name.clone(),
                issue: "无效数据规则".to_string(),
                count: to_i32(sheet.broken_validations.len()),
                action: if cleaned { "已修复" } else { "可修复" }.to_string(),
            });
        }
    }
    if rows.is_empty() {
        rows.push(ResultRow {
            location: file_label(&report.source),
            issue: "未发现异常".to_string(),
            count: 0,
            action: "跳过".to_string(),
        });
    }
    rows
}

fn first_output_path(batch: &BatchResult) -> Option<String> {
    batch
        .reports
        .iter()
        .find_map(|report| report.output.as_ref().map(|path| path_to_string(path.as_path())))
}

fn path_to_string(path: &Path) -> String {
    path.display().to_string()
}

fn file_label(path: &Path) -> String {
    path.file_name()
        .and_then(|name| name.to_str())
        .unwrap_or_else(|| path.to_str().unwrap_or(""))
        .to_string()
}

fn to_i32(value: usize) -> i32 {
    i32::try_from(value).unwrap_or(i32::MAX)
}
