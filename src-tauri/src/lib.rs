pub mod excel;

use excel::{ConversionResult, ConvertedRow, ColumnInfo, ColumnMapping};
use std::sync::Mutex;
use std::collections::HashMap;
use tauri::State;

struct AppState {
    converted_data: Mutex<Option<Vec<ConvertedRow>>>,
    source_path: Mutex<Option<String>>,
}

#[tauri::command]
fn read_columns(path: String) -> Result<Vec<ColumnInfo>, String> {
    excel::read_columns(&path)
}

#[tauri::command]
fn convert_with_mapping(
    path: String,
    mappings: HashMap<String, ColumnMapping>,
    state: State<AppState>,
) -> Result<ConversionResult, String> {
    let result = excel::read_and_convert_with_mapping(&path, mappings)?;
    let mut data = state.converted_data.lock().map_err(|e| format!("锁定状态失败: {}", e))?;
    *data = Some(result.rows.clone());
    let mut src = state.source_path.lock().map_err(|e| format!("锁定状态失败: {}", e))?;
    *src = Some(path);
    Ok(result)
}

#[tauri::command]
fn convert_file(path: String, state: State<AppState>) -> Result<ConversionResult, String> {
    let result = excel::read_and_convert(&path)?;
    let mut data = state.converted_data.lock().map_err(|e| format!("锁定状态失败: {}", e))?;
    *data = Some(result.rows.clone());
    Ok(result)
}

#[tauri::command]
fn export_file(output_path: String, state: State<AppState>) -> Result<String, String> {
    let data = state.converted_data.lock().map_err(|e| format!("锁定状态失败: {}", e))?;
    let rows = data.as_ref().ok_or("没有可导出的数据，请先选择并转换文件")?;
    excel::write_output(rows, &output_path)?;
    Ok(format!("成功导出 {} 条数据", rows.len()))
}

#[tauri::command]
fn merge_duplicates(state: State<AppState>) -> Result<ConversionResult, String> {
    let mut data = state.converted_data.lock().map_err(|e| format!("锁定状态失败: {}", e))?;
    let rows = data.as_ref().ok_or("没有可合并的数据，请先选择并转换文件")?;
    let merged = excel::merge_duplicates(rows);
    let total_rows = merged.len();
    *data = Some(merged.clone());
    Ok(ConversionResult {
        rows: merged,
        total_rows,
        has_duplicates: false,
        duplicate_count: 0,
    })
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .manage(AppState {
            converted_data: Mutex::new(None),
            source_path: Mutex::new(None),
        })
        .invoke_handler(tauri::generate_handler![convert_file, export_file, merge_duplicates, read_columns, convert_with_mapping])
        .run(tauri::generate_context!())
        .expect("启动应用失败");
}
