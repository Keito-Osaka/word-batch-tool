mod edge_integration;
use serde::Deserialize;
use serde_json::Value;
use std::{
    fs,
    io::{BufRead, BufReader, Read, Write},
    path::{Path, PathBuf},
    process::{Command, Stdio},
    thread,
};
use tauri::{AppHandle, Emitter, Manager};

#[cfg(target_os = "windows")]
use std::os::windows::process::CommandExt;

#[cfg(target_os = "windows")]
const CREATE_NO_WINDOW: u32 = 0x08000000;

fn backend_path(app: &AppHandle) -> Result<PathBuf, String> {
    let root = app.path().resource_dir().map_err(|e| e.to_string())?;
    [
        root.join("backend/word-batch-backend.exe"),
        root.join("resources/backend/word-batch-backend.exe"),
    ]
    .into_iter()
    .find(|p| p.exists())
    .ok_or_else(|| "Pythonバックエンドが見つかりません。".to_string())
}

fn spawn_backend(
    app: &AppHandle,
    command: &str,
    request: Value,
) -> Result<std::process::Child, String> {
    let mut backend_command = Command::new(backend_path(app)?);

    backend_command
        .arg(command)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());

    #[cfg(target_os = "windows")]
    backend_command.creation_flags(CREATE_NO_WINDOW);

    let mut child = backend_command.spawn().map_err(|e| e.to_string())?;
    let request_bytes = serde_json::to_vec(&request).map_err(|e| e.to_string())?;

    {
        let mut stdin = child
            .stdin
            .take()
            .ok_or_else(|| "Pythonバックエンドの標準入力を開けません。".to_string())?;
        stdin.write_all(&request_bytes).map_err(|e| e.to_string())?;
        stdin.flush().map_err(|e| e.to_string())?;
    }

    // stdinをここで確実に閉じ、Python側のread()へEOFを通知する。
    Ok(child)
}

fn execute_backend(
    app: &AppHandle,
    command: &str,
    request: Value,
) -> Result<Value, String> {
    let child = spawn_backend(app, command, request)?;
    let output = child.wait_with_output().map_err(|e| e.to_string())?;

    if !output.status.success() {
        return Err(String::from_utf8_lossy(&output.stderr).to_string());
    }

    serde_json::from_slice(&output.stdout)
        .map_err(|e| format!("バックエンド応答解析エラー: {e}"))
}

#[tauri::command]
async fn warmup_backend(app: AppHandle) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || {
        execute_backend(&app, "warmup", serde_json::json!({}))
    })
    .await
    .map_err(|e| format!("ウォームアップ処理の実行に失敗しました: {e}"))?
}

#[tauri::command]
async fn warmup_template(app: AppHandle, template_path: String) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || {
        execute_backend(
            &app,
            "warmup-template",
            serde_json::json!({"template_path": template_path}),
        )
    })
    .await
    .map_err(|e| format!("テンプレート確認処理の実行に失敗しました: {e}"))?
}

#[tauri::command]
async fn inspect_data(app: AppHandle, request: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || execute_backend(&app, "inspect", request))
        .await
        .map_err(|e| format!("置換データ確認処理の実行に失敗しました: {e}"))?
}

fn generate_documents_blocking(app: AppHandle, request: Value) -> Result<Value, String> {
    let mut child = spawn_backend(&app, "generate", request)?;
    let stdout = child
        .stdout
        .take()
        .ok_or_else(|| "Pythonバックエンドの標準出力を開けません。".to_string())?;
    let mut stderr = child
        .stderr
        .take()
        .ok_or_else(|| "Pythonバックエンドの標準エラーを開けません。".to_string())?;

    let stderr_handle = thread::spawn(move || {
        let mut text = String::new();
        let _ = stderr.read_to_string(&mut text);
        text
    });

    let mut result: Option<Value> = None;

    for line in BufReader::new(stdout).lines() {
        let line = line.map_err(|e| e.to_string())?;
        if line.trim().is_empty() {
            continue;
        }

        let message: Value = serde_json::from_str(&line)
            .map_err(|e| format!("進捗応答解析エラー: {e}"))?;

        match message.get("type").and_then(Value::as_str) {
            Some("progress") => {
                app.emit("generation-progress", &message)
                    .map_err(|e| e.to_string())?;
            }
            Some("result") => {
                result = message.get("data").cloned();
            }
            _ => {}
        }
    }

    let status = child.wait().map_err(|e| e.to_string())?;
    let error_text = stderr_handle
        .join()
        .unwrap_or_else(|_| "エラー情報を取得できませんでした。".to_string());

    if !status.success() {
        return Err(error_text);
    }

    result.ok_or_else(|| "処理結果を取得できませんでした。".to_string())
}

#[tauri::command]
async fn generate_documents(app: AppHandle, request: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || generate_documents_blocking(app, request))
        .await
        .map_err(|e| format!("文書生成処理の実行に失敗しました: {e}"))?
}


#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DocumentNumberRecord {
    version: u32,
    document_number: String,
    source: String,
    captured_at: String,
    #[serde(default)]
    document_title: Option<String>,
    #[serde(default)]
    page_type: Option<String>,
    #[serde(default)]
    page_title: Option<String>,
    #[serde(default)]
    enforcement_date: Option<String>,
}

#[tauri::command]
fn read_document_number() -> Result<Value, String> {
    let local_app_data = std::env::var_os("LOCALAPPDATA")
        .ok_or_else(|| "文書番号の保存先を取得できませんでした。".to_string())?;
    let path = PathBuf::from(local_app_data)
        .join("WordBatchTool")
        .join("integration")
        .join("document-number.json");

    if !path.exists() {
        return Err(
            "文書番号の連携データがありません。行政文書管理システムで文書番号を取得し、Edge拡張機能の「アプリへ送信」を押してください。"
                .to_string(),
        );
    }

    let bytes = fs::read(&path)
        .map_err(|e| format!("文書番号の連携データを読み取れませんでした: {e}"))?;
    let record: DocumentNumberRecord = serde_json::from_slice(&bytes)
        .map_err(|e| format!("文書番号の連携データが壊れています。Edge拡張機能から再送信してください: {e}"))?;

    if record.version != 1 && record.version != 2 {
        return Err("対応していない文書番号データです。Edge拡張機能から再送信してください。".to_string());
    }
    let document_number = record.document_number.trim();
    if document_number.is_empty() || document_number.chars().count() > 100 {
        return Err("文書番号の形式が正しくありません。Edge拡張機能から再送信してください。".to_string());
    }
    if record.source != "panfocus" {
        return Err("文書番号の取得元を確認できませんでした。".to_string());
    }

    Ok(serde_json::json!({
        "version": record.version,
        "documentNumber": document_number,
        "source": record.source,
        "capturedAt": record.captured_at,
        "path": path.to_string_lossy(),
        "documentTitle": record.document_title,
        "pageType": record.page_type,
        "pageTitle": record.page_title,
        "enforcementDate": record.enforcement_date,
    }))
}


#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DocumentContextBridgeResponse {
    version: u32,
    #[serde(rename = "type")]
    kind: String,
    request_id: String,
    ok: bool,
    #[serde(default)]
    error: Option<String>,
    #[serde(default)]
    document_number: Option<String>,
    #[serde(default)]
    document_title: Option<String>,
    #[serde(default)]
    page_type: Option<String>,
    #[serde(default)]
    page_title: Option<String>,
    #[serde(default)]
    enforcement_date: Option<String>,
    #[serde(default)]
    source: Option<String>,
    #[serde(default)]
    captured_at: Option<String>,
}

fn document_bridge_directory() -> Result<PathBuf, String> {
    let local_app_data = std::env::var_os("LOCALAPPDATA")
        .ok_or_else(|| "文書情報連携の保存先を取得できませんでした。".to_string())?;
    Ok(PathBuf::from(local_app_data).join("WordBatchTool").join("integration"))
}

fn create_bridge_request_id() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH)
        .map(|value| value.as_nanos()).unwrap_or_default();
    format!("{}-{}", std::process::id(), nanos)
}

fn validate_bridge_response(response: DocumentContextBridgeResponse, request_id: &str) -> Result<Value, String> {
    if response.version != 1 || response.kind != "current-document-context-result" || response.request_id != request_id {
        return Err("文書情報連携の応答形式が正しくありません。".to_string());
    }
    if !response.ok {
        return Err(response.error.unwrap_or_else(|| "Microsoft Edgeから文書情報を取得できませんでした。".to_string()));
    }
    let document_number = response.document_number.unwrap_or_default().trim().to_string();
    if document_number.is_empty() || document_number.chars().count() > 100 {
        return Err("取得した文書番号の形式が正しくありません。".to_string());
    }
    if response.source.as_deref() != Some("panfocus") {
        return Err("文書情報の取得元を確認できませんでした。".to_string());
    }
    Ok(serde_json::json!({
        "version": 2,
        "requestId": request_id,
        "documentNumber": document_number,
        "documentTitle": response.document_title,
        "pageType": response.page_type,
        "pageTitle": response.page_title,
        "enforcementDate": response.enforcement_date,
        "source": response.source,
        "capturedAt": response.captured_at,
        "path": document_bridge_directory()?.join("bridge-response.json").to_string_lossy(),
    }))
}

#[tauri::command]
async fn request_current_document_context() -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(|| {
        use std::time::{Duration, Instant};
        let directory = document_bridge_directory()?;
        fs::create_dir_all(&directory).map_err(|e| format!("連携フォルダーを作成できませんでした: {e}"))?;
        let request_path = directory.join("bridge-request.json");
        let response_path = directory.join("bridge-response.json");
        let _ = fs::remove_file(&response_path);
        let request_id = create_bridge_request_id();
        let request = serde_json::json!({
            "version": 1,
            "type": "get-current-document-context",
            "requestId": request_id,
            "createdAtUnixMs": std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)
                .map(|value| value.as_millis()).unwrap_or_default(),
        });
        let temporary_path = request_path.with_extension("json.tmp");
        fs::write(&temporary_path, serde_json::to_vec_pretty(&request).map_err(|e| e.to_string())?)
            .map_err(|e| format!("文書情報の取得要求を書き込めませんでした: {e}"))?;
        if request_path.exists() { let _ = fs::remove_file(&request_path); }
        fs::rename(&temporary_path, &request_path)
            .map_err(|e| format!("文書情報の取得要求を確定できませんでした: {e}"))?;
        let started = Instant::now();
        while started.elapsed() < Duration::from_secs(20) {
            if response_path.exists() {
                let bytes = fs::read(&response_path).map_err(|e| e.to_string())?;
                if let Ok(response) = serde_json::from_slice::<DocumentContextBridgeResponse>(&bytes) {
                    if response.request_id == request_id {
                        let result = validate_bridge_response(response, &request_id);
                        let _ = fs::remove_file(&request_path);
                        let _ = fs::remove_file(&response_path);
                        return result;
                    }
                }
            }
            std::thread::sleep(Duration::from_millis(100));
        }
        let _ = fs::remove_file(&request_path);
        Err("Microsoft Edgeから応答がありませんでした。対象文書を表示し、連携ブリッジと拡張機能が導入されていることを確認してください。".to_string())
    }).await.map_err(|e| format!("文書情報取得処理に失敗しました: {e}"))?
}

#[tauri::command]
fn classify_dropped_paths(paths: Vec<String>) -> Value {
    let mut template: Option<String> = None;
    let mut data: Option<String> = None;
    let mut output: Option<String> = None;
    let mut unsupported: Vec<String> = Vec::new();

    for value in paths {
        let path = Path::new(&value);
        if path.is_dir() {
            output = Some(value);
            continue;
        }

        let extension = path
            .extension()
            .and_then(|item| item.to_str())
            .unwrap_or("")
            .to_ascii_lowercase();

        match extension.as_str() {
            "docx" => template = Some(value),
            "xlsx" | "xlsm" | "xls" | "csv" => data = Some(value),
            _ => unsupported.push(value),
        }
    }

    serde_json::json!({
        "template_path": template,
        "data_path": data,
        "output_path": output,
        "unsupported_paths": unsupported,
    })
}

#[tauri::command]
fn open_output_folder(path: String) -> Result<(), String> {
    let requested = Path::new(&path);
    if !requested.exists() {
        return Err(format!("出力先が見つかりません。\n{path}"));
    }

    let folder = if requested.is_dir() {
        requested
    } else {
        requested
            .parent()
            .ok_or_else(|| format!("出力先を特定できません。\n{path}"))?
    };

    #[cfg(target_os = "windows")]
    {
        let mut explorer = Command::new("explorer.exe");
        explorer.arg(folder).creation_flags(CREATE_NO_WINDOW);
        explorer
            .spawn()
            .map_err(|e| format!("エクスプローラーを起動できませんでした: {e}"))?;
        return Ok(());
    }

    #[cfg(target_os = "macos")]
    {
        Command::new("open")
            .arg(folder)
            .spawn()
            .map_err(|e| e.to_string())?;
        return Ok(());
    }

    #[cfg(all(unix, not(target_os = "macos")))]
    {
        Command::new("xdg-open")
            .arg(folder)
            .spawn()
            .map_err(|e| e.to_string())?;
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
            warmup_template,
            inspect_data,
            generate_documents,
            classify_dropped_paths,
            open_output_folder,
            read_document_number,
            request_current_document_context,
            edge_integration::get_edge_integration_status,
            edge_integration::check_edge_integration,
            edge_integration::open_edge_integration_setup,
            edge_integration::repair_edge_integration
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
