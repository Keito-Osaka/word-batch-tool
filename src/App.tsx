import { useEffect, useMemo, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { open } from "@tauri-apps/plugin-dialog";
import { getCurrentWebviewWindow } from "@tauri-apps/api/webviewWindow";
import {
  Check,
  ChevronDown,
  ChevronRight,
  ChevronUp,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Moon,
  Play,
  RotateCcw,
  Settings as SettingsIcon,
  Sun,
  Trash2,
  X,
} from "lucide-react";
import type { DataPreview, DroppedPathClassification, GenerateResult, GenerationProgress, Settings } from "./types";

const defaults: Settings = {
  outputFormat: "word",
  outputMethod: "folder",
  addSerialNumber: true,
  serialDigits: 2,
  formatAmountWithComma: true,
  rowExcludeMode: "selected_column_number_empty",
  targetColumnNumber: 2,
  filenameKeys: [],
  fastPdfSplitEnabled: true,
};

const loadSettings = (): Settings => {
  try {
    const saved = JSON.parse(localStorage.getItem("wordBatchSettings") || "{}");
    const migratedKeys = Array.isArray(saved.filenameKeys) ? saved.filenameKeys : [];
    const wasOldImplicitDefault = migratedKeys.length === 1 && migratedKeys[0] === "名前（漢字）" && !localStorage.getItem("wordBatchFilenameRuleV2");
    const filenameKeys = wasOldImplicitDefault ? [] : migratedKeys;
    const addSerialNumber = filenameKeys.length === 0 ? true : saved.addSerialNumber ?? true;
    localStorage.setItem("wordBatchFilenameRuleV2", "1");
    return { ...defaults, ...saved, filenameKeys, addSerialNumber };
  } catch {
    return defaults;
  }
};

function Segmented<T extends string>({
  value,
  onChange,
  items,
}: {
  value: T;
  onChange: (value: T) => void;
  items: { value: T; label: string; disabled?: boolean }[];
}) {
  return (
    <div className="segmented">
      {items.map((item) => (
        <button
          key={item.value}
          className={value === item.value ? "active" : ""}
          disabled={item.disabled}
          onClick={() => !item.disabled && onChange(item.value)}
        >
          {item.label}
        </button>
      ))}
    </div>
  );
}

function Picker({
  kind,
  title,
  path,
  meta,
  disabled,
  onPick,
}: {
  kind: "word" | "excel" | "folder";
  title: string;
  path: string;
  meta?: string;
  disabled?: boolean;
  onPick: () => void;
}) {
  const Icon = kind === "word" ? FileText : kind === "excel" ? FileSpreadsheet : FolderOpen;
  return (
    <button className="picker" disabled={disabled} onClick={onPick}>
      <span className="picker-icon"><Icon size={20} /></span>
      <span className="picker-body">
        <small>{title}</small>
        <strong>{path ? path.split(/[\\/]/).pop() : "クリックして選択"}</strong>
        <span>{meta || (path ? "選択済み" : "クリックして選択、または画面へドラッグ＆ドロップ")}</span>
      </span>
      <span className="change">{path ? "変更" : "選択"}</span>
    </button>
  );
}

function Confirmation({
  title,
  message,
  confirmLabel,
  onCancel,
  onConfirm,
}: {
  title: string;
  message: string;
  confirmLabel: string;
  onCancel: () => void;
  onConfirm: () => void;
}) {
  return (
    <div className="overlay center-overlay" role="presentation" onMouseDown={onCancel}>
      <section className="confirm-dialog" role="dialog" aria-modal="true" onMouseDown={(event) => event.stopPropagation()}>
        <h2>{title}</h2>
        <p>{message}</p>
        <div className="dialog-actions">
          <button onClick={onCancel}>キャンセル</button>
          <button className="danger-button" onClick={onConfirm}>{confirmLabel}</button>
        </div>
      </section>
    </div>
  );
}

export default function App() {
  const [dark, setDark] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [previewOpen, setPreviewOpen] = useState(false);
  const [previewTab, setPreviewTab] = useState<"included" | "excluded">("included");
  const [resetConfirmOpen, setResetConfirmOpen] = useState(false);
  const [templatePath, setTemplatePath] = useState("");
  const [dataPath, setDataPath] = useState("");
  const [outputPath, setOutputPath] = useState("");
  const [preview, setPreview] = useState<DataPreview | null>(null);
  const [settings, setSettings] = useState<Settings>(loadSettings);
  const [busy, setBusy] = useState(false);
  const [warmupState, setWarmupState] = useState<"running" | "ready" | "error">("running");
  const [warmupError, setWarmupError] = useState("");
  const [status, setStatus] = useState("準備完了");
  const [result, setResult] = useState<GenerateResult | null>(null);
  const [error, setError] = useState("");
  const [progress, setProgress] = useState<GenerationProgress | null>(null);
  const [dragActive, setDragActive] = useState(false);
  const [dropNotice, setDropNotice] = useState("");
  const [templateWarmupState, setTemplateWarmupState] = useState<"idle" | "running" | "ready" | "error">("idle");

  useEffect(() => {
    localStorage.setItem("wordBatchSettings", JSON.stringify(settings));
  }, [settings]);

  useEffect(() => {
    let unlisten: (() => void) | undefined;
    listen<GenerationProgress>("generation-progress", (event) => setProgress(event.payload)).then((fn) => { unlisten = fn; });
    return () => { if (unlisten) unlisten(); };
  }, []);

  useEffect(() => {
    let unlisten: (() => void) | undefined;
    getCurrentWebviewWindow().onDragDropEvent(async (event) => {
      if (busy) return;
      if (event.payload.type === "enter" || event.payload.type === "over") { setDragActive(true); return; }
      if (event.payload.type === "leave") { setDragActive(false); return; }
      if (event.payload.type !== "drop") return;
      setDragActive(false);
      try {
        const classified = await invoke<DroppedPathClassification>("classify_dropped_paths", { paths: event.payload.paths });
        const accepted: string[] = [];
        if (classified.template_path) { setTemplatePath(classified.template_path); accepted.push("テンプレート"); await prepareTemplate(classified.template_path); }
        if (classified.output_path) { setOutputPath(classified.output_path); accepted.push("出力先"); }
        if (classified.data_path) {
          if (warmupState !== "ready") throw new Error("文書処理の準備が完了してから置換データをドロップしてください。");
          setDataPath(classified.data_path); accepted.push("置換データ"); await inspect(classified.data_path);
        }
        if (classified.unsupported_paths.length) setError(`対応していないファイル形式です。
${classified.unsupported_paths.join("
")}`); else setError("");
        if (accepted.length) { setResult(null); setDropNotice(`${accepted.join("・")}を選択しました`); window.setTimeout(() => setDropNotice(""), 2400); }
      } catch (reason) { setError(String(reason)); }
    }).then((fn) => { unlisten = fn; });
    return () => { if (unlisten) unlisten(); };
  }, [busy, warmupState, settings.formatAmountWithComma, settings.rowExcludeMode, settings.targetColumnNumber]);

  useEffect(() => {
    let active = true;
    (async () => {
      try {
        await invoke("warmup_backend");
        if (active) setWarmupState("ready");
      } catch (reason) {
        if (active) {
          setWarmupError(String(reason));
          setWarmupState("error");
        }
      }
    })();
    return () => { active = false; };
  }, []);

  useEffect(() => {
    if (!preview) return;
    setSettings((current) => ({
      ...current,
      filenameKeys: current.filenameKeys.filter((key) => preview.columns.includes(key)),
    }));
  }, [preview?.columns.join("|")]);

  const canRun = Boolean(
    templatePath && dataPath && outputPath && preview?.included_count && !busy && warmupState === "ready" && templateWarmupState !== "running" && (settings.addSerialNumber || settings.filenameKeys.length > 0),
  );
  const hasWork = Boolean(templatePath || dataPath || outputPath || preview || result || error);

  const exampleName = useMemo(() => {
    if (settings.outputMethod !== "folder") {
      const extension = settings.outputMethod === "zip" ? "zip" : settings.outputFormat === "pdf" ? "pdf" : "docx";
      const templateName = (templatePath.split(/[\\/]/).pop() || "テンプレート.docx").replace(/\.docx$/i, "");
      const count = preview?.included_count || 0;
      return `${templateName}_${count}件一式.${extension}`;
    }
    const row = preview?.included_rows[0];
    const parts: string[] = [];
    if (settings.addSerialNumber) {
      parts.push(String(1).padStart(settings.serialDigits, "0"));
    }
    for (const key of settings.filenameKeys) {
      const value = row?.[key]?.trim();
      if (value) parts.push(value.replace(/\s+/g, " "));
    }
    const sourceName = templatePath.split(/[\\/]/).pop() || "テンプレート.docx";
    parts.push(sourceName.replace(/\.docx$/i, settings.outputFormat === "pdf" ? ".pdf" : ".docx"));
    return parts.join("_");
  }, [preview, settings.addSerialNumber, settings.filenameKeys, settings.serialDigits, settings.outputFormat, settings.outputMethod, templatePath]);

  async function prepareTemplate(path: string) {
    setTemplateWarmupState("running");
    try {
      await invoke("warmup_template", { templatePath: path });
      setTemplateWarmupState("ready");
    } catch (reason) {
      setTemplateWarmupState("error");
      setError(`テンプレートを確認できませんでした。\n${String(reason)}`);
    }
  }

  async function chooseTemplate() {
    const path = await open({ multiple: false, filters: [{ name: "Word", extensions: ["docx"] }] });
    if (typeof path === "string") {
      setTemplatePath(path);
      setResult(null);
      setError("");
      await prepareTemplate(path);
    }
  }

  async function chooseData() {
    const path = await open({
      multiple: false,
      filters: [{ name: "置換データ", extensions: ["xlsx", "xlsm", "xls", "csv"] }],
    });
    if (typeof path === "string") {
      setDataPath(path);
      setResult(null);
      await inspect(path);
    }
  }

  async function chooseOutput() {
    const path = await open({ directory: true, multiple: false });
    if (typeof path === "string") {
      setOutputPath(path);
      setResult(null);
      setError("");
    }
  }

  async function inspect(path = dataPath) {
    if (!path) return;
    setBusy(true);
    setError("");
    setStatus("置換データを確認しています...");
    try {
      const response = await invoke<DataPreview>("inspect_data", {
        request: {
          data_path: path,
          format_amount_with_comma: settings.formatAmountWithComma,
          row_exclude_mode: settings.rowExcludeMode,
          row_exclude_target_column_number: settings.targetColumnNumber,
          output_format: settings.outputFormat,
          output_method: settings.outputMethod,
          fast_pdf_split_enabled: settings.fastPdfSplitEnabled,
        },
      });
      setPreview(response);
      setStatus(`使用${response.included_count}件・除外${response.excluded_count}件`);
    } catch (reason) {
      setPreview(null);
      setError(String(reason));
      setStatus("読込エラー");
    } finally {
      setBusy(false);
    }
  }

  async function generate() {
    if (!canRun) return;
    setBusy(true);
    setError("");
    setResult(null);
    setProgress({ type: "progress", stage: "prepare", label: "処理を開始しています", current: 0, total: preview?.included_count || 0, percent: 0, detail: "", state: "running" });
    setStatus("処理を開始しています...");
    try {
      const response = await invoke<GenerateResult>("generate_documents", {
        request: {
          template_path: templatePath,
          data_path: dataPath,
          output_path: outputPath,
          filename_keys: settings.filenameKeys,
          add_serial_number: settings.addSerialNumber,
          serial_digits: settings.serialDigits,
          format_amount_with_comma: settings.formatAmountWithComma,
          row_exclude_mode: settings.rowExcludeMode,
          row_exclude_target_column_number: settings.targetColumnNumber,
          output_format: settings.outputFormat,
          output_method: settings.outputMethod,
          fast_pdf_split_enabled: settings.fastPdfSplitEnabled,
        },
      });
      setResult(response);
      setProgress({ type: "progress", stage: "complete", label: "処理が完了しました", current: response.generated_count, total: response.generated_count, percent: 100, detail: "", state: "done" });
      setStatus("完了");
    } catch (reason) {
      setError(String(reason));
      setProgress((current) => current ? { ...current, state: "error", label: `${current.label}でエラーが発生しました` } : null);
      setStatus("処理エラー");
    } finally {
      setBusy(false);
    }
  }

  async function openOutputFolder() {
    if (!outputPath || busy) return;
    setError("");
    try {
      await invoke("open_output_folder", { path: result?.output_path || outputPath });
    } catch (reason) {
      setError(`出力先を開けませんでした。\n${String(reason)}`);
    }
  }

  function resetWork() {
    setTemplatePath("");
    setTemplateWarmupState("idle");
    setDataPath("");
    setOutputPath("");
    setPreview(null);
    setResult(null);
    setError("");
    setStatus("準備完了");
    setProgress(null);
    setPreviewOpen(false);
    setResetConfirmOpen(false);
  }

  function requestReset() {
    if (!hasWork) return;
    setResetConfirmOpen(true);
  }

  function moveFilenameKey(index: number, direction: -1 | 1) {
    setSettings((current) => {
      const next = [...current.filenameKeys];
      const target = index + direction;
      if (target < 0 || target >= next.length) return current;
      [next[index], next[target]] = [next[target], next[index]];
      return { ...current, filenameKeys: next };
    });
  }

  function toggleFilenameKey(key: string) {
    setSettings((current) => {
      const isSelected = current.filenameKeys.includes(key);
      if (isSelected && current.filenameKeys.length === 1 && !current.addSerialNumber) return current;
      return {
        ...current,
        filenameKeys: isSelected
          ? current.filenameKeys.filter((item) => item !== key)
          : [...current.filenameKeys, key],
      };
    });
  }

  const tableRows = previewTab === "included"
    ? preview?.included_rows.map((row, index) => ({ key: `i-${index}`, number: index + 1, row, reason: "" })) || []
    : preview?.excluded_rows.map((item, index) => ({ key: `e-${index}`, number: item.source_row_number, row: item.row, reason: item.reason })) || [];

  return (
    <div className={dark ? "app dark" : "app"}>
      {dragActive && <div className="native-drop-overlay" aria-live="polite"><div className="native-drop-panel"><span className="drop-symbol"><FileSpreadsheet size={28} /></span><strong>ここにドロップして選択</strong><p>Word、Excel・CSV、または出力先フォルダを自動で判別します。</p><div><span>Word</span><span>Excel / CSV</span><span>フォルダ</span></div></div></div>}
      {dropNotice && <div className="drop-toast"><Check size={15} />{dropNotice}</div>}
      <header>
        <div className="brand-block">
          <div className="app-symbol" aria-hidden="true"><FileText size={19} /></div>
          <div>
            <span className="eyebrow">DOCUMENT AUTOMATION</span>
            <h1>Wordファイル一括作成</h1>
            <p>テンプレートと置換データから、文書をすばやく作成します。</p>
          </div>
        </div>
        <div className="header-actions">
          <span className={`readiness ${warmupState}`}><i />{warmupState === "ready" ? "準備完了" : warmupState === "running" ? "準備中" : "要確認"}</span>
          <button className="icon-button" onClick={() => setDark((value) => !value)} aria-label="テーマ切替">
            {dark ? <Sun size={18} /> : <Moon size={18} />}
          </button>
          <button className="icon-button" disabled={busy} onClick={() => setSettingsOpen(true)} aria-label="設定">
            <SettingsIcon size={18} />
          </button>
        </div>
      </header>

      <main>
        {warmupState === "running" && (
          <section className="warmup-card" aria-live="polite">
            <span className="warmup-spinner" />
            <div><strong>文書処理を準備しています</strong><p>初回のみ数秒かかることがあります。準備が整うと置換データを選択できます。</p></div>
          </section>
        )}
        {warmupState === "error" && (
          <section className="error" role="alert"><strong>文書処理を準備できませんでした</strong><p>{warmupError}</p></section>
        )}
        {templateWarmupState === "running" && (
          <section className="template-warmup" aria-live="polite"><span className="warmup-spinner" /><div><strong>テンプレートを確認しています</strong><p>初回の文書作成をすばやく開始できるよう準備しています。</p></div></section>
        )}
        <section className="picker-grid">
          <Picker kind="word" title="テンプレート" path={templatePath} disabled={busy} meta={templateWarmupState === "running" ? "テンプレートを確認しています..." : templateWarmupState === "ready" ? "文書作成の準備が整いました" : undefined} onPick={chooseTemplate} />
          <Picker
            kind="excel"
            title="置換データ"
            path={dataPath}
            disabled={busy || warmupState !== "ready"}
            meta={warmupState === "running" ? "文書処理の準備が整うまでお待ちください" : preview ? `使用${preview.included_count}件・除外${preview.excluded_count}件` : undefined}
            onPick={chooseData}
          />
        </section>

        {preview && (
          <button className="preview-link" onClick={() => { setPreviewTab("included"); setPreviewOpen(true); }}>
            <span>
              <strong>置換データを確認</strong>
              <small>読込{preview.original_count}件、使用{preview.included_count}件、除外{preview.excluded_count}件</small>
            </span>
            <ChevronRight size={18} />
          </button>
        )}

        <Picker kind="folder" title="出力先" path={outputPath} disabled={busy} onPick={chooseOutput} />

        <section className="options-card">
          <div>
            <label>出力形式</label>
            <Segmented
              value={settings.outputFormat}
              onChange={(value) => setSettings((current) => ({ ...current, outputFormat: value }))}
              items={[{ value: "word", label: "Word" }, { value: "pdf", label: "PDF" }]}
            />
            
          </div>
          <div>
            <label>出力方法</label>
            <Segmented
              value={settings.outputMethod}
              onChange={(value) => setSettings((current) => ({ ...current, outputMethod: value }))}
              items={[
                { value: "folder", label: "個別" },
                { value: "merged", label: "結合" },
                { value: "zip", label: "ZIP" },
              ]}
            />
            
          </div>
          <button className="details" onClick={() => setSettingsOpen(true)}>
            ファイル名と詳細設定 <ChevronRight size={17} />
          </button>
        </section>

        {templatePath && (
          <section className="filename">
            <small>出力ファイル名の例</small>
            <code>{exampleName}</code>
          </section>
        )}

        {busy && progress && (
          <section className="progress-card" aria-live="polite">
            <div className="progress-heading"><div><strong>{progress.label}</strong><p>{progress.detail || "処理が完了するまで、この画面を閉じずにお待ちください。"}</p></div><span>{progress.percent == null ? "処理中" : `${progress.percent}%`}</span></div>
            <div className={progress.percent == null ? "progress-track indeterminate" : "progress-track"}><span style={progress.percent == null ? undefined : { width: `${progress.percent}%` }} /></div>
            <div className="progress-meta"><span>{progress.current > 0 && progress.total > 0 ? `${progress.current} / ${progress.total}件` : "処理を準備中"}</span><span>{progress.stage === "fallback" ? "通常方式へ切替" : ""}</span></div>
            <div className="stage-list">
              {[
                ["prepare", "データ準備"],
                ["word", settings.outputFormat === "pdf" ? "PDF変換用文書の作成" : "Word文書作成"],
                ...(settings.outputFormat === "pdf" ? [["pdf", "PDF変換"], ...(settings.outputMethod === "folder" && settings.fastPdfSplitEnabled ? [["split", "個別PDFへ分割"]] : [])] : []),
                ...(settings.outputMethod === "zip" ? [["zip", "ZIP作成"]] : []),
                ["complete", "完了処理"],
              ].map(([key, label], index, all) => {
                const currentIndex = all.findIndex(([stage]) => stage === progress.stage);
                const state = progress.stage === "fallback" ? (key === "word" ? "active" : index < 1 ? "done" : "waiting") : index < currentIndex ? "done" : index === currentIndex ? "active" : "waiting";
                return <div className={`stage-item ${state}`} key={key}><span>{state === "done" ? "✓" : state === "active" ? "●" : "○"}</span><strong>{label}</strong><small>{state === "done" ? "完了" : state === "active" ? "実行中" : "待機中"}</small></div>;
              })}
            </div>
          </section>
        )}

        {error && (
          <section className="error" role="alert">
            <strong>処理できませんでした</strong>
            <p>{error}</p>
            {dataPath && <button onClick={() => inspect()}>置換データを再読み込み</button>}
          </section>
        )}

        {result ? (
          <section className="success">
            <span><Check size={18} /></span>
            <div>
              <strong>{result.generated_count}件の{result.format === "pdf" ? "PDF" : "Word"}処理が完了しました</strong>
              <p>{result.output_path}</p>
              <div className="success-actions">
                <button className="primary" onClick={openOutputFolder}><FolderOpen size={16} />出力先を開く</button>
                <button onClick={() => { setResult(null); setProgress(null); setStatus("準備完了"); }}><X size={16} />閉じる</button>
              </div>
            </div>
          </section>
        ) : (
          <section className="runbar">
            <div>
              <span><strong>{preview?.included_count || 0}件</strong>の{settings.outputFormat === "pdf" ? "PDF" : "Word"}を{settings.outputMethod === "folder" ? "個別作成" : settings.outputMethod === "merged" ? "結合して作成" : "ZIPにまとめて作成"}します</span>
              <small>{status}</small>
            </div>
            <div className="run-actions">
              <button className="primary" disabled={!canRun} onClick={generate}>
                <Play size={16} />{busy ? "処理中..." : "複製を開始"}
              </button>
              <button disabled={!hasWork || busy} onClick={requestReset}><RotateCcw size={16} />リセット</button>
              <button disabled={!outputPath || busy} onClick={openOutputFolder}><FolderOpen size={16} />出力先を開く</button>
            </div>
          </section>
        )}
      </main>

      {settingsOpen && (
        <div className="overlay" onMouseDown={() => setSettingsOpen(false)}>
          <aside className="drawer" onMouseDown={(event) => event.stopPropagation()}>
            <div className="drawer-head">
              <div><h2>詳細設定</h2><p>変更内容は自動保存されます。</p></div>
              <button className="icon-button" onClick={() => setSettingsOpen(false)}><X size={18} /></button>
            </div>
            <div className="drawer-body">
              <h3>ファイル名</h3>
              <label className="check">
                <input
                  type="checkbox"
                  checked={settings.addSerialNumber}
                  disabled={settings.filenameKeys.length === 0}
                  onChange={(event) => setSettings((current) => ({ ...current, addSerialNumber: event.target.checked || current.filenameKeys.length === 0 }))}
                />
                先頭に通し番号を付ける
              </label>
              {settings.filenameKeys.length === 0 && <small className="setting-note">列が選択されていないため、通し番号は必須です。</small>}
              <label>
                桁数
                <input
                  type="number"
                  min="1"
                  max="6"
                  disabled={!settings.addSerialNumber}
                  value={settings.serialDigits}
                  onChange={(event) => setSettings((current) => ({ ...current, serialDigits: Math.min(6, Math.max(1, Number(event.target.value) || 1)) }))}
                />
              </label>

              <div className="setting-subhead">
                <strong>ファイル名に使用する列</strong>
                <small>選択した順にファイル名へ追加します。</small>
              </div>
              {!preview ? (
                <p className="muted-box">置換データを読み込むと列を選択できます。</p>
              ) : (
                <>
                  <div className="column-choice-list">
                    {preview.columns.map((column) => (
                      <label className="check column-choice" key={column}>
                        <input
                          type="checkbox"
                          checked={settings.filenameKeys.includes(column)}
                          disabled={settings.filenameKeys.includes(column) && settings.filenameKeys.length === 1 && !settings.addSerialNumber}
                          onChange={() => toggleFilenameKey(column)}
                        />
                        <span>{column}</span>
                      </label>
                    ))}
                  </div>
                  <div className="sort-list">
                    {settings.filenameKeys.length === 0 ? (
                      <p className="empty-selection">列は選択されていません。</p>
                    ) : settings.filenameKeys.map((key, index) => (
                      <div className="sort-item" key={key}>
                        <span className="sort-number">{index + 1}</span>
                        <strong>{key}</strong>
                        <div>
                          <button disabled={index === 0} onClick={() => moveFilenameKey(index, -1)} aria-label="上へ"><ChevronUp size={16} /></button>
                          <button disabled={index === settings.filenameKeys.length - 1} onClick={() => moveFilenameKey(index, 1)} aria-label="下へ"><ChevronDown size={16} /></button>
                          <button disabled={settings.filenameKeys.length === 1 && !settings.addSerialNumber} onClick={() => toggleFilenameKey(key)} aria-label="削除"><Trash2 size={16} /></button>
                        </div>
                      </div>
                    ))}
                  </div>
                </>
              )}
              <div className="drawer-preview"><small>出力例</small><code>{exampleName}</code></div>

              <h3>行の除外</h3>
              <label>
                条件
                <select
                  value={settings.rowExcludeMode}
                  onChange={(event) => setSettings((current) => ({ ...current, rowExcludeMode: event.target.value as Settings["rowExcludeMode"] }))}
                >
                  <option value="any_empty_except_first">1列目以外に空欄があれば除外</option>
                  <option value="any_empty">どこかに空欄があれば除外</option>
                  <option value="all_empty_except_first">1列目以外がすべて空なら除外</option>
                  <option value="selected_column_number_empty">指定列が空なら除外</option>
                </select>
              </label>
              <label>
                指定列番号
                <input
                  type="number"
                  min="1"
                  disabled={settings.rowExcludeMode !== "selected_column_number_empty"}
                  value={settings.targetColumnNumber}
                  onChange={(event) => setSettings((current) => ({ ...current, targetColumnNumber: Math.max(1, Number(event.target.value) || 1) }))}
                />
              </label>

              <h3>金額</h3>
              <label className="check">
                <input
                  type="checkbox"
                  checked={settings.formatAmountWithComma}
                  onChange={(event) => setSettings((current) => ({ ...current, formatAmountWithComma: event.target.checked }))}
                />
                金額を3桁区切りにする
              </label>

              <h3>PDF</h3>
              <label className="check">
                <input type="checkbox" checked={settings.fastPdfSplitEnabled} onChange={(event) => setSettings((current) => ({ ...current, fastPdfSplitEnabled: event.target.checked }))} />
                個別PDF出力を高速化する
              </label>
              <small className="setting-note">分割できない場合は自動的に通常方式へ切り替えます。</small>

              <button className="primary full" onClick={async () => { setSettingsOpen(false); if (dataPath) await inspect(); }}>
                設定を適用
              </button>
            </div>
          </aside>
        </div>
      )}

      {previewOpen && preview && (
        <div className="overlay center-overlay" onMouseDown={() => setPreviewOpen(false)}>
          <section className="preview-modal" onMouseDown={(event) => event.stopPropagation()}>
            <div className="drawer-head">
              <div>
                <h2>置換データの確認</h2>
                <p>使用{preview.included_count}件・除外{preview.excluded_count}件・記入例{preview.example_count}件</p>
              </div>
              <button className="icon-button" onClick={() => setPreviewOpen(false)}><X size={18} /></button>
            </div>
            <div className="preview-tabs">
              <button className={previewTab === "included" ? "active" : ""} onClick={() => setPreviewTab("included")}>使用するデータ <span>{preview.included_count}</span></button>
              <button className={previewTab === "excluded" ? "active" : ""} onClick={() => setPreviewTab("excluded")}>除外されたデータ <span>{preview.excluded_count}</span></button>
            </div>
            <div className="table-help">表の下部にあるスクロールバーで、右側の列まで確認できます。</div>
            <div className="table-scroll" tabIndex={0}>
              <table>
                <thead>
                  <tr>
                    <th className="sticky-number">No.</th>
                    {preview.columns.map((column) => <th key={column}>{column}</th>)}
                    {previewTab === "excluded" && <th className="reason-column">除外理由</th>}
                  </tr>
                </thead>
                <tbody>
                  {tableRows.length === 0 ? (
                    <tr><td className="empty-table" colSpan={preview.columns.length + 2}>該当するデータはありません。</td></tr>
                  ) : tableRows.map((item) => (
                    <tr key={item.key}>
                      <td className="sticky-number">{item.number}</td>
                      {preview.columns.map((column) => <td key={column}>{item.row[column]}</td>)}
                      {previewTab === "excluded" && <td className="reason-column">{item.reason}</td>}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      )}

      {resetConfirmOpen && (
        <Confirmation
          title="現在の選択内容をクリアしますか？"
          message="テンプレート、置換データ、出力先の選択を解除します。保存済みの環境設定は変更されません。"
          confirmLabel="クリア"
          onCancel={() => setResetConfirmOpen(false)}
          onConfirm={resetWork}
        />
      )}
    </div>
  );
}
