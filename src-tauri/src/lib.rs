use serde_json::Value;
use std::{io::Write, path::PathBuf, process::{Command, Stdio}};
use tauri::{AppHandle, Manager};

fn backend_path(app: &AppHandle) -> Result<PathBuf,String> {
  let root=app.path().resource_dir().map_err(|e|e.to_string())?;
  let candidates=[root.join("backend/word-batch-backend.exe"),root.join("resources/backend/word-batch-backend.exe")];
  candidates.into_iter().find(|p|p.exists()).ok_or_else(||"Pythonバックエンドが見つかりません。".into())
}
fn execute(app:&AppHandle, command:&str, request:Value)->Result<Value,String>{
  let exe=backend_path(app)?;
  let mut child=Command::new(exe).arg(command).stdin(Stdio::piped()).stdout(Stdio::piped()).stderr(Stdio::piped()).spawn().map_err(|e|e.to_string())?;
  child.stdin.as_mut().ok_or("標準入力を開けません。")?.write_all(serde_json::to_string(&request).map_err(|e|e.to_string())?.as_bytes()).map_err(|e|e.to_string())?;
  let output=child.wait_with_output().map_err(|e|e.to_string())?;
  if !output.status.success(){return Err(String::from_utf8_lossy(&output.stderr).to_string())}
  serde_json::from_slice(&output.stdout).map_err(|e|format!("バックエンド応答解析エラー: {e}"))
}
#[tauri::command] fn inspect_data(app:AppHandle,request:Value)->Result<Value,String>{execute(&app,"inspect",request)}
#[tauri::command] fn generate_documents(app:AppHandle,request:Value)->Result<Value,String>{execute(&app,"generate",request)}
#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run(){tauri::Builder::default().plugin(tauri_plugin_dialog::init()).plugin(tauri_plugin_opener::init()).invoke_handler(tauri::generate_handler![inspect_data,generate_documents]).run(tauri::generate_context!()).expect("error while running tauri application");}
