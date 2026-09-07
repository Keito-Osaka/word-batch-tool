use serde_json::Value;
use std::{
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Stdio},
};
use tauri::{AppHandle, Manager};

fn backend_path(app: &AppHandle) -> Result<PathBuf, String> {
    let root = app.path().resource_dir().map_err(|error| error.to_string())?;
    let candidates = [
        root.join("backend/word-batch-backend.exe"),
        root.join("resources/backend/word-batch-backend.exe"),
    ];
    candidates
        .into_iter()
        .find(|path| path.exists())
        .ok_or_else(|| "Pythonバックエンドが見つかりません。".to_string())
}

fn execute_backend(app: &AppHandle, command: &str, request: Value) -> Result<Value, String> {
    let executable = backend_path(app)?;
    let mut child = Command::new(executable)
        .arg(command)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|error| error.to_string())?;

    let request_bytes = serde_json::to_vec(&request).map_err(|error| error.to_string())?;
    child
        .stdin
        .as_mut()
        .ok_or_else(|| "Pythonバックエンドの標準入力を開けません。".to_string())?
        .write_all(&request_bytes)
        .map_err(|error| error.to_string())?;

    let output = child.wait_with_output().map_err(|error| error.to_string())?;
    if !output.status.success() {
        return Err(String::from_utf8_lossy(&output.stderr).to_string());
    }
    serde_json::from_slice(&output.stdout)
        .map_err(|error| format!("バックエンド応答解析エラー: {error}"))
}

#[tauri::command]
fn warmup_backend(app: AppHandle) -> Result<Value, String> {
    execute_backend(&app, "warmup", serde_json::json!({}))
}

#[tauri::command]
fn inspect_data(app: AppHandle, request: Value) -> Result<Value, String> {
    execute_backend(&app, "inspect", request)
}

#[tauri::command]
fn generate_documents(app: AppHandle, request: Value) -> Result<Value, String> {
    execute_backend(&app, "generate", request)
}

#[tauri::command]
fn open_output_folder(path: String) -> Result<(), String> {
    let requested = Path::new(&path);
    if !requested.exists() {
        return Err(format!("出力先フォルダが見つかりません。\n{path}"));
    }
    let folder = if requested.is_dir() { requested } else { requested.parent().ok_or_else(|| format!("出力先を特定できません。\n{path}"))? };

    #[cfg(target_os = "windows")]
    {
        Command::new("explorer.exe")
            .arg(folder)
            .spawn()
            .map_err(|error| format!("エクスプローラーを起動できませんでした: {error}"))?;
        return Ok(());
    }

    #[cfg(target_os = "macos")]
    {
        Command::new("open").arg(folder).spawn().map_err(|error| error.to_string())?;
        return Ok(());
    }

    #[cfg(all(unix, not(target_os = "macos")))]
    {
        Command::new("xdg-open").arg(folder).spawn().map_err(|error| error.to_string())?;
        return Ok(());
    }

    #[allow(unreachable_code)]
    Err("このOSでは出力先を開けません。".to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            warmup_backend,
            inspect_data,
            generate_documents,
            open_output_folder
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
