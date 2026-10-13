#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use serde::{Deserialize, Serialize};
use std::env;
use std::fs;
use tauri_plugin_dialog::DialogExt;

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct SystemPrinter {
    pub name: String,
    pub is_default: bool,
    pub port_name: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct FileFilter {
    pub name: String,
    pub extensions: Vec<String>,
}

#[tauri::command]
async fn open_files(app: tauri::AppHandle, filters: Vec<FileFilter>) -> Result<Vec<String>, String> {
    let (tx, rx) = std::sync::mpsc::channel();
    let mut builder = app.dialog().file();

    for filter in filters {
        let exts: Vec<&str> = filter.extensions.iter().map(|s| s.as_str()).collect();
        builder = builder.add_filter(filter.name, &exts);
    }

    builder.pick_files(move |paths| {
        let _ = tx.send(paths);
    });

    match rx.recv().map_err(|e| e.to_string())? {
        Some(paths) => {
            let mut result = Vec::new();
            for path in paths {
                let resolved_path = path.into_path().map_err(|e| e.to_string())?;
                result.push(resolved_path.display().to_string());
            }
            Ok(result)
        }
        None => Ok(vec![]),
    }
}

#[tauri::command]
async fn save_file(app: tauri::AppHandle, suggested_name: String, filters: Vec<FileFilter>) -> Result<String, String> {
    let (tx, rx) = std::sync::mpsc::channel();
    let mut builder = app.dialog().file().set_file_name(suggested_name);

    for filter in filters {
        let exts: Vec<&str> = filter.extensions.iter().map(|s| s.as_str()).collect();
        builder = builder.add_filter(filter.name, &exts);
    }

    builder.save_file(move |path| {
        let _ = tx.send(path);
    });

    match rx.recv().map_err(|e| e.to_string())? {
        Some(path) => {
            let resolved_path = path.into_path().map_err(|e| e.to_string())?;
            Ok(resolved_path.display().to_string())
        }
        None => Err("No file selected".to_string()),
    }
}

#[tauri::command]
fn read_file(path: String) -> Result<Vec<u8>, String> {
    fs::read(&path).map_err(|e| e.to_string())
}

#[tauri::command]
fn write_file(path: String, data: Vec<u8>) -> Result<(), String> {
    fs::write(&path, data).map_err(|e| e.to_string())
}

#[tauri::command]
fn get_temp_dir() -> String {
    env::temp_dir().display().to_string()
}

#[tauri::command]
fn open_url(url: String) -> Result<(), String> {
    if !url.starts_with("http://") && !url.starts_with("https://") {
        return Err("Invalid URL scheme".to_string());
    }

    #[cfg(target_os = "windows")]
    {
        use std::os::windows::process::CommandExt;
        const CREATE_NO_WINDOW: u32 = 0x08000000;
        std::process::Command::new("cmd")
            .args(["/c", "start", "", &url])
            .creation_flags(CREATE_NO_WINDOW)
            .spawn()
            .map_err(|e| e.to_string())?;
        Ok(())
    }
    #[cfg(target_os = "macos")]
    {
        std::process::Command::new("open")
            .arg(&url)
            .spawn()
            .map_err(|e| e.to_string())?;
        Ok(())
    }
    #[cfg(target_os = "linux")]
    {
        std::process::Command::new("xdg-open")
            .arg(&url)
            .spawn()
            .map_err(|e| e.to_string())?;
        Ok(())
    }
    #[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
    {
        Err("Unsupported platform for opening URL".to_string())
    }
}

#[tauri::command]
fn get_system_printers() -> Result<Vec<SystemPrinter>, String> {
    #[cfg(target_os = "windows")]
    {
        use std::os::windows::process::CommandExt;
        const CREATE_NO_WINDOW: u32 = 0x08000000;
        let script = r#"
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
            $printers = Get-CimInstance Win32_Printer | Select-Object -Property Name, Default, PortName;
            if ($printers) {
                if ($printers -isnot [array]) {
                    $printers = @($printers)
                }
                ConvertTo-Json -InputObject $printers -Compress
            } else {
                "[]"
            }
        "#;
        let output = std::process::Command::new("powershell")
            .args(["-NoProfile", "-NonInteractive", "-Command", script])
            .creation_flags(CREATE_NO_WINDOW)
            .output()
            .map_err(|e| e.to_string())?;

        if !output.status.success() {
            let err_str = String::from_utf8_lossy(&output.stderr);
            return Err(format!("Failed to query printers: {}", err_str));
        }

        let json_str = String::from_utf8_lossy(&output.stdout).trim().to_string();
        if json_str.is_empty() || json_str == "null" {
            return Ok(vec![]);
        }

        #[derive(Deserialize)]
        struct RawPrinter {
            #[serde(rename = "Name")]
            name: Option<String>,
            #[serde(rename = "Default")]
            is_default: Option<bool>,
            #[serde(rename = "PortName")]
            port_name: Option<String>,
        }

        if let Ok(raw_list) = serde_json::from_str::<Vec<RawPrinter>>(&json_str) {
            let printers = raw_list
                .into_iter()
                .filter_map(|p| {
                    p.name.map(|name| SystemPrinter {
                        name,
                        is_default: p.is_default.unwrap_or(false),
                        port_name: p.port_name,
                    })
                })
                .collect();
            return Ok(printers);
        } else if let Ok(single) = serde_json::from_str::<RawPrinter>(&json_str) {
            if let Some(name) = single.name {
                return Ok(vec![SystemPrinter {
                    name,
                    is_default: single.is_default.unwrap_or(false),
                    port_name: single.port_name,
                }]);
            }
        }

        Ok(vec![])
    }

    #[cfg(target_os = "macos")]
    {
        let output = std::process::Command::new("lpstat")
            .args(["-p"])
            .output()
            .map_err(|e| e.to_string())?;
        let stdout = String::from_utf8_lossy(&output.stdout);
        let mut printers = Vec::new();
        for line in stdout.lines() {
            if let Some(rest) = line.strip_prefix("printer ") {
                if let Some(name) = rest.split_whitespace().next() {
                    printers.push(SystemPrinter {
                        name: name.to_string(),
                        is_default: false,
                        port_name: None,
                    });
                }
            }
        }
        Ok(printers)
    }

    #[cfg(target_os = "linux")]
    {
        let output = std::process::Command::new("lpstat")
            .args(["-p"])
            .output()
            .map_err(|e| e.to_string())?;
        let stdout = String::from_utf8_lossy(&output.stdout);
        let mut printers = Vec::new();
        for line in stdout.lines() {
            if let Some(rest) = line.strip_prefix("printer ") {
                if let Some(name) = rest.split_whitespace().next() {
                    printers.push(SystemPrinter {
                        name: name.to_string(),
                        is_default: false,
                        port_name: None,
                    });
                }
            }
        }
        Ok(printers)
    }

    #[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
    {
        Ok(vec![])
    }
}

#[tauri::command]
fn print_pdf_file(data: Vec<u8>, printer_name: Option<String>) -> Result<(), String> {
    let timestamp = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or(0);
    let temp_file_path = env::temp_dir().join(format!("pdfcraft_print_{}.pdf", timestamp));
    fs::write(&temp_file_path, &data).map_err(|e| e.to_string())?;

    let file_path_str = temp_file_path.to_string_lossy().to_string();

    #[cfg(target_os = "windows")]
    {
        use std::os::windows::process::CommandExt;
        const CREATE_NO_WINDOW: u32 = 0x08000000;

        let script = if let Some(printer) = printer_name {
            format!(
                "Start-Process -FilePath '{}' -Verb PrintTo -ArgumentList '\"{}\"'",
                file_path_str.replace("'", "''"),
                printer.replace("'", "''")
            )
        } else {
            format!(
                "Start-Process -FilePath '{}' -Verb Print",
                file_path_str.replace("'", "''")
            )
        };

        let output = std::process::Command::new("powershell")
            .args(["-NoProfile", "-NonInteractive", "-Command", &script])
            .creation_flags(CREATE_NO_WINDOW)
            .output()
            .map_err(|e| e.to_string())?;

        if !output.status.success() {
            let err_str = String::from_utf8_lossy(&output.stderr);
            return Err(format!("Print execution failed: {}", err_str));
        }

        Ok(())
    }

    #[cfg(any(target_os = "macos", target_os = "linux"))]
    {
        let mut cmd = std::process::Command::new("lpr");
        if let Some(printer) = printer_name {
            cmd.arg("-P").arg(printer);
        }
        cmd.arg(&file_path_str);
        let output = cmd.output().map_err(|e| e.to_string())?;
        if !output.status.success() {
            let err_str = String::from_utf8_lossy(&output.stderr);
            return Err(format!("Print execution failed: {}", err_str));
        }
        Ok(())
    }

    #[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
    {
        Err("Unsupported operating system for printing".to_string())
    }
}

fn main() {
    #[cfg(target_os = "linux")]
    {
        // Mitigate WebKitGTK / Mesa EGL display initialization failure under modern Wayland (EGL_BAD_PARAMETER)
        if env::var("WEBKIT_DISABLE_DMABUF_RENDERER").is_err() {
            #[allow(unused_unsafe)]
            unsafe {
                env::set_var("WEBKIT_DISABLE_DMABUF_RENDERER", "1");
            }
        }
    }

    #[cfg(target_os = "windows")]
    {
        if let Ok(exe_path) = env::current_exe() {
            if let Some(exe_dir) = exe_path.parent() {
                let data_dir = exe_dir.join("data");
                let portable_marker = exe_dir.join("portable.txt");
                let is_portable = exe_path
                    .file_name()
                    .map(|n| n.to_string_lossy().to_lowercase().contains("portable"))
                    .unwrap_or(false)
                    || data_dir.exists()
                    || portable_marker.exists();

                if is_portable {
                    let _ = fs::create_dir_all(&data_dir);
                    #[allow(unused_unsafe)]
                    unsafe {
                        env::set_var("WEBVIEW2_USER_DATA_FOLDER", &data_dir);
                    }
                }
            }
        }
    }

    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![
            open_files,
            save_file,
            read_file,
            write_file,
            get_temp_dir,
            open_url,
            get_system_printers,
            print_pdf_file
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
