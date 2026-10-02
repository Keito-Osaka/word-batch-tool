use serde_json::{json,Value};use std::{fs,path::{Path,PathBuf},process::Command,time::{Duration,Instant,SystemTime,UNIX_EPOCH}};use tauri::{AppHandle,Manager};
fn root()->Result<PathBuf,String>{Ok(PathBuf::from(std::env::var_os("LOCALAPPDATA").ok_or("LOCALAPPDATA unavailable")?).join("WordBatchTool"))}fn int()->Result<PathBuf,String>{Ok(root()?.join("integration"))}fn conf()->Result<bool,String>{let r=root()?;Ok(r.join("NativeMessaging/word-batch-native-host.exe").exists()&&r.join("EdgeExtension/manifest.json").exists())}fn last()->Option<String>{let v:Value=serde_json::from_slice(&fs::read(int().ok()?.join("edge-state.json")).ok()?).ok()?;v["lastResponseAt"].as_str().map(str::to_string)}
#[tauri::command]pub fn get_edge_integration_status()->Result<Value,String>{Ok(json!({"configured":conf()?,"lastResponseAt":last()}))}#[tauri::command]pub fn open_edge_integration_setup()->Result<(),String>{Command::new("explorer.exe").arg(root()?.join("EdgeExtension")).spawn().map_err(|e|e.to_string())?;Command::new("cmd.exe").args(["/C","start","","msedge.exe","edge://extensions/"]).spawn().map_err(|e|e.to_string())?;Ok(())}#[tauri::command]pub fn repair_edge_integration(app: AppHandle) -> Result<(), String> {
    let resource_directory = app.path().resource_dir().map_err(|error| error.to_string())?;
    let app_root = root()?;
    let extension_directory = app_root.join("EdgeExtension");
    let native_directory = app_root.join("NativeMessaging");
    let integration_directory = app_root.join("integration");
    copy_directory(&resource_directory.join("edge-extension"), &extension_directory)?;
    fs::create_dir_all(&native_directory).map_err(|error| error.to_string())?;
    fs::create_dir_all(&integration_directory).map_err(|error| error.to_string())?;
    let host_path = native_directory.join("word-batch-native-host.exe");
    fs::copy(resource_directory.join("native-messaging").join("word-batch-native-host.exe"), &host_path).map_err(|error| error.to_string())?;
    let manifest_path = native_directory.join("jp.keito.word_batch_tool.json");
    let manifest = json!({
        "name": "jp.keito.word_batch_tool",
        "description": "WordBatchTool Native Messaging Host",
        "path": host_path.to_string_lossy(),
        "type": "stdio",
        "allowed_origins": ["chrome-extension://omgfaofenjcomglepofcopbfmjhcgdmd/"]
    });
    fs::write(&manifest_path, serde_json::to_vec_pretty(&manifest).map_err(|error| error.to_string())?).map_err(|error| error.to_string())?;
    let key = r"HKCU\Software\Microsoft\Edge\NativeMessagingHosts\jp.keito.word_batch_tool";
    let status = Command::new("reg.exe").args(["add", key, "/ve", "/t", "REG_SZ", "/d", &manifest_path.to_string_lossy(), "/f"]).status().map_err(|error| error.to_string())?;
    if !status.success() { return Err("Native Messagingの登録を修復できませんでした。".to_string()); }
    for file in ["bridge-request.json", "bridge-response.json"] { let _ = fs::remove_file(integration_directory.join(file)); }
    Ok(())
}

pub async fn check_edge_integration()->Result<Value,String>{tauri::async_runtime::spawn_blocking(||{let d=int()?;fs::create_dir_all(&d).map_err(|e|e.to_string())?;let id=format!("health-{}",SystemTime::now().duration_since(UNIX_EPOCH).unwrap_or_default().as_nanos());fs::write(d.join("bridge-request.json"),serde_json::to_vec(&json!({"version":1,"type":"bridge-health-check","requestId":id,"createdAtUnixMs":SystemTime::now().duration_since(UNIX_EPOCH).unwrap_or_default().as_millis()})).unwrap()).map_err(|e|e.to_string())?;let st=Instant::now();while st.elapsed()<Duration::from_secs(10){if let Ok(b)=fs::read(d.join("bridge-response.json")){if let Ok(v)=serde_json::from_slice::<Value>(&b){if v["requestId"]==id&&v["type"]=="bridge-health-check-result"{let n=chrono::Local::now().to_rfc3339();fs::write(d.join("edge-state.json"),serde_json::to_vec(&json!({"lastResponseAt":n})).unwrap()).map_err(|e|e.to_string())?;return Ok(json!({"configured":conf()?,"lastResponseAt":n}))}}}std::thread::sleep(Duration::from_millis(100));}Err("Edge拡張機能から診断応答がありませんでした。".into())}).await.map_err(|e|e.to_string())?}fn copy(s:&Path,d:&Path)->Result<(),String>{fs::create_dir_all(d).map_err(|e|e.to_string())?;for e in fs::read_dir(s).map_err(|e|e.to_string())?{let e=e.map_err(|e|e.to_string())?,t=d.join(e.file_name());if e.path().is_dir(){copy(&e.path(),&t)?}else{fs::copy(e.path(),t).map_err(|e|e.to_string())?;}}Ok(())}
