use serde_json::{json, Value};
use std::{
    fs,
    path::{Path, PathBuf},
    process::Command,
    time::{Duration, Instant, SystemTime, UNIX_EPOCH},
};
use tauri::{AppHandle, Manager};

fn root() -> Result<PathBuf, String> {
    let local_app_data = std::env::var_os("LOCALAPPDATA")
        .ok_or_else(|| "LOCALAPPDATAを取得できませんでした。".to_string())?;
    Ok(PathBuf::from(local_app_data).join("WordBatchTool"))
}

fn integration_directory() -> Result<PathBuf, String> {
    Ok(root()?.join("integration"))
}

fn configured() -> Result<bool, String> {
    let app_root = root()?;
    Ok(app_root
        .join("NativeMessaging")
        .join("word-batch-native-host.exe")
        .exists()
        && app_root
            .join("EdgeExtension")
            .join("manifest.json")
            .exists())
}

fn last_response_at() -> Option<String> {
    let path = integration_directory().ok()?.join("edge-state.json");
    let value: Value = serde_json::from_slice(&fs::read(path).ok()?).ok()?;
    value
        .get("lastResponseAt")?
        .as_str()
        .map(ToString::to_string)
}

fn now_unix_milliseconds() -> u128 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis()
}

pub fn get_edge_integration_status() -> Result<Value, String> {
    Ok(json!({
        "configured": configured()?,
        "lastResponseAt": last_response_at()
    }))
}

pub fn open_edge_integration_setup() -> Result<(), String> {
    Command::new("explorer.exe")
        .arg(root()?.join("EdgeExtension"))
        .spawn()
        .map_err(|error| error.to_string())?;

    Command::new("cmd.exe")
        .args(["/C", "start", "", "msedge.exe", "edge://extensions/"])
        .spawn()
        .map_err(|error| error.to_string())?;

    Ok(())
}

pub fn repair_edge_integration(app: AppHandle) -> Result<(), String> {
    let resource_directory = app
        .path()
        .resource_dir()
        .map_err(|error| error.to_string())?;
    let app_root = root()?;

    copy_directory(
        &resource_directory.join("edge-extension"),
        &app_root.join("EdgeExtension"),
    )?;

    let native_messaging_directory = app_root.join("NativeMessaging");
    fs::create_dir_all(&native_messaging_directory)
        .map_err(|error| error.to_string())?;

    fs::copy(
        resource_directory
            .join("native-messaging")
            .join("word-batch-native-host.exe"),
        native_messaging_directory.join("word-batch-native-host.exe"),
    )
    .map_err(|error| error.to_string())?;

    Ok(())
}

pub async fn check_edge_integration() -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(|| {
        let directory = integration_directory()?;
        fs::create_dir_all(&directory).map_err(|error| error.to_string())?;

        let request_id = format!(
            "health-{}",
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap_or_default()
                .as_nanos()
        );

        let request = json!({
            "version": 1,
            "type": "bridge-health-check",
            "requestId": request_id,
            "createdAtUnixMs": now_unix_milliseconds()
        });

        fs::write(
            directory.join("bridge-request.json"),
            serde_json::to_vec_pretty(&request).map_err(|error| error.to_string())?,
        )
        .map_err(|error| error.to_string())?;

        let started = Instant::now();
        while started.elapsed() < Duration::from_secs(10) {
            if let Ok(bytes) = fs::read(directory.join("bridge-response.json")) {
                if let Ok(response) = serde_json::from_slice::<Value>(&bytes) {
                    let response_request_id =
                        response.get("requestId").and_then(Value::as_str);
                    let response_type = response.get("type").and_then(Value::as_str);

                    if response_request_id == Some(request_id.as_str())
                        && response_type == Some("bridge-health-check-result")
                    {
                        let response_at = now_unix_milliseconds().to_string();
                        fs::write(
                            directory.join("edge-state.json"),
                            serde_json::to_vec_pretty(&json!({
                                "lastResponseAt": response_at
                            }))
                            .map_err(|error| error.to_string())?,
                        )
                        .map_err(|error| error.to_string())?;

                        return Ok(json!({
                            "configured": configured()?,
                            "lastResponseAt": response_at
                        }));
                    }
                }
            }

            std::thread::sleep(Duration::from_millis(100));
        }

        Err("Edge拡張機能から診断応答がありませんでした。".to_string())
    })
    .await
    .map_err(|error| error.to_string())?
}

fn copy_directory(source: &Path, destination: &Path) -> Result<(), String> {
    if !source.exists() {
        return Err(format!(
            "連携用リソースが見つかりません: {}",
            source.display()
        ));
    }

    fs::create_dir_all(destination).map_err(|error| error.to_string())?;

    for entry_result in fs::read_dir(source).map_err(|error| error.to_string())? {
        let entry = entry_result.map_err(|error| error.to_string())?;
        let source_path = entry.path();
        let destination_path = destination.join(entry.file_name());

        if source_path.is_dir() {
            copy_directory(&source_path, &destination_path)?;
        } else {
            fs::copy(&source_path, &destination_path)
                .map_err(|error| error.to_string())?;
        }
    }

    Ok(())
}
