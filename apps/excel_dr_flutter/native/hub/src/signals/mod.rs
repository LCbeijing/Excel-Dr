use rinf::{DartSignal, RustSignal, SignalPiece};
use serde::{Deserialize, Serialize};

#[derive(Deserialize, DartSignal)]
pub struct TaskRequest {
    pub request_id: i32,
    pub operation: String,
    pub target_kind: String,
    pub path: String,
}

#[derive(Serialize, RustSignal)]
pub struct TaskProgress {
    pub request_id: i32,
    pub percent: i32,
    pub current: String,
    pub detail: String,
}

#[derive(Serialize, RustSignal)]
pub struct TaskResult {
    pub request_id: i32,
    pub operation: String,
    pub target_kind: String,
    pub status: String,
    pub message: String,
    pub output_path: String,
    pub scanned_files: i32,
    pub actionable_files: i32,
    pub suspicious_objects: i32,
    pub broken_rules: i32,
    pub output_files: i32,
    pub skipped_files: i32,
    pub failed_files: i32,
    pub rows: Vec<ResultRow>,
    pub warnings: Vec<String>,
    pub errors: Vec<String>,
}

#[derive(Serialize, SignalPiece)]
pub struct ResultRow {
    pub location: String,
    pub issue: String,
    pub count: i32,
    pub action: String,
}
