use serde_json::Value;
use std::{
    io::{BufRead, BufReader, Read, Write},
    path::{Path, PathBuf},
    process::{Command, Stdio},
    thread,
};
use tauri::{AppHandle, Emitter, Manager};

fn backend_path(app: &AppHandle) -> Result<PathBuf, String> {
    let root = app.path().resource_dir().map_err(|e| e.to_string())?;
    [root.join("backend/word-batch-backend.exe"), root.join("resources/backend/word-batch-backend.exe")]
        .into_iter().find(|p| p.exists()).ok_or_else(|| "Pythonバックエンドが見つかりません。".to_string())
}

fn spawn_backend(app: &AppHandle, command: &str, request: Value) -> Result<std::process::Child, String> {
    let mut child = Command::new(backend_path(app)?)
        .arg(command).stdin(Stdio::piped()).stdout(Stdio::piped()).stderr(Stdio::piped())
        .spawn().map_err(|e| e.to_string())?;
    child.stdin.as_mut().ok_or_else(|| "Pythonバックエンドの標準入力を開けません。".to_string())?
        .write_all(&serde_json::to_vec(&request).map_err(|e| e.to_string())?).map_err(|e| e.to_string())?;
    Ok(child)
}

fn execute_backend(app: &AppHandle, command: &str, request: Value) -> Result<Value, String> {
    let child = spawn_backend(app, command, request)?;
    let output = child.wait_with_output().map_err(|e| e.to_string())?;
    if !output.status.success() { return Err(String::from_utf8_lossy(&output.stderr).to_string()); }
    serde_json::from_slice(&output.stdout).map_err(|e| format!("バックエンド応答解析エラー: {e}"))
}

#[tauri::command]
fn warmup_backend(app: AppHandle) -> Result<Value, String> { execute_backend(&app, "warmup", serde_json::json!({})) }

#[tauri::command]
fn inspect_data(app: AppHandle, request: Value) -> Result<Value, String> { execute_backend(&app, "inspect", request) }

#[tauri::command]
fn generate_documents(app: AppHandle, request: Value) -> Result<Value, String> {
    let mut child = spawn_backend(&app, "generate", request)?;
    let stdout = child.stdout.take().ok_or_else(|| "Pythonバックエンドの標準出力を開けません。".to_string())?;
    let mut stderr = child.stderr.take().ok_or_else(|| "Pythonバックエンドの標準エラーを開けません。".to_string())?;
    let stderr_handle = thread::spawn(move || { let mut text = String::new(); let _ = stderr.read_to_string(&mut text); text });
    let mut result: Option<Value> = None;
    for line in BufReader::new(stdout).lines() {
        let line = line.map_err(|e| e.to_string())?;
        if line.trim().is_empty() { continue; }
        let message: Value = serde_json::from_str(&line).map_err(|e| format!("進捗応答解析エラー: {e}"))?;
        match message.get("type").and_then(Value::as_str) {
            Some("progress") => { app.emit("generation-progress", &message).map_err(|e| e.to_string())?; }
            Some("result") => { result = message.get("data").cloned(); }
            _ => {}
        }
    }
    let status = child.wait().map_err(|e| e.to_string())?;
    let error_text = stderr_handle.join().unwrap_or_else(|_| "エラー情報を取得できませんでした。".to_string());
    if !status.success() { return Err(error_text); }
    result.ok_or_else(|| "処理結果を取得できませんでした。".to_string())
}

#[tauri::command]
fn open_output_folder(path: String) -> Result<(), String> {
    let requested = Path::new(&path);
    if !requested.exists() { return Err(format!("出力先が見つかりません。\n{path}")); }
    let folder = if requested.is_dir() { requested } else { requested.parent().ok_or_else(|| format!("出力先を特定できません。\n{path}"))? };
    #[cfg(target_os = "windows")]
    { Command::new("explorer.exe").arg(folder).spawn().map_err(|e| format!("エクスプローラーを起動できませんでした: {e}"))?; return Ok(()); }
    #[cfg(target_os = "macos")]
    { Command::new("open").arg(folder).spawn().map_err(|e| e.to_string())?; return Ok(()); }
    #[cfg(all(unix, not(target_os = "macos")))]
    { Command::new("xdg-open").arg(folder).spawn().map_err(|e| e.to_string())?; return Ok(()); }
    #[allow(unreachable_code)] Err("このOSでは出力先を開けません。".to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init()).plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![warmup_backend, inspect_data, generate_documents, open_output_folder])
        .run(tauri::generate_context!()).expect("error while running tauri application");
}
