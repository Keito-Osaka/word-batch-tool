import { useEffect, useMemo, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { open } from "@tauri-apps/plugin-dialog";
import { getCurrentWebviewWindow } from "@tauri-apps/api/webviewWindow";
import {
  Check,
  CircleHelp,
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
import type { DataPreview, DroppedPathClassification, GenerateResult, GenerationProgress, Settings, TemplateInspection, CommonValues } from "./types";

const APP_VERSION = "Ver.1.3.0";
const DEFAULT_AMOUNT_INCLUDE_KEYWORDS = [
  "交付申請額",
  "交付決定額",
  "申請額",
  "金額",
  "費用",
  "参加費用",
  "支援金",
];

const defaults: Settings = {
  outputFormat: "word",
  outputMethod: "folder",
  addSerialNumber: true,
  serialDigits: 2,
  formatAmountWithComma: true,
  amountIncludeKeywords: DEFAULT_AMOUNT_INCLUDE_KEYWORDS,
  rowExcludeMode: "selected_columns_any_empty",
  excludeExampleRows: true,
  rowExcludeColumns: [],
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
    const amountIncludeKeywords = Array.isArray(saved.amountIncludeKeywords)
      ? saved.amountIncludeKeywords
      : DEFAULT_AMOUNT_INCLUDE_KEYWORDS;
    const oldMode = saved.rowExcludeMode === "selected_column_number_empty" ? "selected_columns_any_empty" : saved.rowExcludeMode;
    const rowExcludeColumns = Array.isArray(saved.rowExcludeColumns) ? saved.rowExcludeColumns : [];
    return { ...defaults, ...saved, rowExcludeMode: oldMode ?? defaults.rowExcludeMode, rowExcludeColumns,
      excludeExampleRows: saved.excludeExampleRows ?? true, filenameKeys, addSerialNumber, amountIncludeKeywords };
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
        <strong>{path ? path.split(/[\\/]/).pop() : kind === "folder" ? "フォルダを選択" : "ファイルを選択"}</strong>
        <span>{meta || (path ? "選択済み" : "クリックして選択、または画面へドラッグ＆ドロップ")}</span>
      </span>
      <span className="change">{path ? "変更" : "選択"}</span>
    </button>
  );
}

function HelpGuide({ onClose }: { onClose: () => void }) {
  const [section, setSection] = useState<"basic" | "template" | "data" | "output" | "history" | "about">("basic");

  return (
    <div className="overlay center-overlay" role="presentation" onMouseDown={onClose}>
      <section className="help-modal help-modal-wide" role="dialog" aria-modal="true" aria-labelledby="help-title" onMouseDown={(event) => event.stopPropagation()}>
        <div className="drawer-head">
          <div><h2 id="help-title">使い方とアプリ情報</h2><p>基本操作、データの準備、更新情報を確認できます。</p></div>
          <button className="icon-button" onClick={onClose} aria-label="使い方を閉じる"><X size={18} /></button>
        </div>
        <div className="help-layout">
          <nav className="help-nav" aria-label="使い方の目次">
            <button className={section === "basic" ? "active" : ""} onClick={() => setSection("basic")}>基本操作</button>
            <button className={section === "template" ? "active" : ""} onClick={() => setSection("template")}>テンプレートの作り方</button>
            <button className={section === "data" ? "active" : ""} onClick={() => setSection("data")}>置換データの作り方</button>
            <button className={section === "output" ? "active" : ""} onClick={() => setSection("output")}>ファイル名と出力</button>
            <button className={section === "history" ? "active" : ""} onClick={() => setSection("history")}>更新履歴</button>
            <button className={section === "about" ? "active" : ""} onClick={() => setSection("about")}>このアプリについて</button>
          </nav>

          <div className="help-content">
            {section === "basic" && (
              <div className="help-section">
                <h3>基本操作</h3>
                <ol className="help-steps">
                  <li><span>1</span><div><strong>テンプレートを選択</strong><p>差し込み項目を含むWordファイルを選びます。ドラッグ＆ドロップにも対応しています。</p></div></li>
                  <li><span>2</span><div><strong>置換データを選択</strong><p>ExcelまたはCSVを選びます。「置換データを確認」から内容と除外行を確認できます。</p></div></li>
                  <li><span>3</span><div><strong>出力先を選択</strong><p>作成したファイルを保存するフォルダを選びます。</p></div></li>
                  <li><span>4</span><div><strong>形式と方法を指定</strong><p>WordまたはPDF、個別・結合・ZIPを選びます。</p></div></li>
                  <li><span>5</span><div><strong>複製を開始</strong><p>件数と出力ファイル名の例を確認してから開始します。</p></div></li>
                </ol>
              </div>
            )}

            {section === "template" && (
              <div className="help-section">
                <h3>テンプレートの作り方</h3>
                <p>Word内の置換したい部分を、半角の二重波括弧で囲みます。括弧内の文字は、置換データの列名と完全に一致させてください。</p>
                <div className="help-example"><small>テンプレートの例</small><code>学校名：{`{{学校名}}`}<br />氏名：{`{{氏名}}`}<br />交付決定額：金{`{{交付決定額}}`}円</code></div>
                <div className="help-rule-list">
                  <p><strong>使用できる形式</strong><span>.docx</span></p>
                  <p><strong>正しい書き方</strong><span><code>{`{{学校名}}`}</code></span></p>
                  <p><strong>避ける書き方</strong><span><code>{`{{ 学校名 }}`}</code></span></p>
                </div>
                <div className="help-note"><strong>全文書で共通する項目</strong><p>通知日や回答期限など、すべての文書で同じ値を使う部分は、半角の&lt;&lt;項目名&gt;&gt;で囲みます。例：&lt;&lt;通知日&gt;&gt;</p></div>
                <div className="help-note"><strong>書式を維持するために</strong><p>プレースホルダー全体を同じ文字サイズ・フォント・装飾にしてください。同じ項目はWord内で複数回使用できます。</p></div>
              </div>
            )}

            {section === "data" && (
              <div className="help-section">
                <h3>置換データの作り方</h3>
                <p>Excel・CSVの1行目を列名として読み込み、2行目以降を置換データとして使用します。</p>
                <div className="help-data-sample" role="table" aria-label="置換データの例">
                  <div className="head">学校名</div><div className="head">氏名</div><div className="head">交付決定額</div>
                  <div>あいうえお高等学校</div><div>山田 太郎</div><div>100000</div>
                  <div>かきくけこ高等学校</div><div>佐藤 花子</div><div>120000</div>
                </div>
                <ul className="help-bullets">
                  <li>1行目の列名を空欄にしないでください。</li>
                  <li>Wordの二重波括弧内と列名を完全に一致させてください。</li>
                  <li>結合セルや、データ途中の説明行は使用しないでください。</li>
                  <li>口座番号など先頭ゼロが必要な値は、Excel上で文字列として管理してください。</li>
                  <li>複数シートがある場合は、先頭のシートを読み込みます。</li>
                </ul>
              </div>
            )}

            {section === "output" && (
              <div className="help-section">
                <h3>ファイル名と出力</h3>
                <div className="help-rule-list">
                  <p><strong>出力形式</strong><span>Word / PDF</span></p>
                  <p><strong>出力方法</strong><span>個別 / 結合 / ZIP</span></p>
                  <p><strong>結合・ZIP名</strong><span>テンプレート名_件数件一式</span></p>
                  <p><strong>同名ファイル</strong><span>上書きせず、末尾に番号を付加</span></p>
                </div>
                <div className="help-note"><strong>ファイル名の識別</strong><p>初期状態では通し番号が付きます。通し番号を付けない場合は、ファイル名に使用する列を1つ以上選択してください。</p></div>
                <div className="help-note"><strong>数値のカンマ区切り</strong><p>登録したキーワードを列名に含む列の値を、カンマ区切り形式へ整形します。対象キーワードは「ファイル名と詳細設定」の金額欄で追加・削除できます。</p></div>
                <div className="help-note"><strong>PDF出力</strong><p>PDF作成には、デスクトップ版Microsoft Wordが必要です。</p></div>
              </div>
            )}

            {section === "history" && (
              <div className="help-section">
                <h3>更新履歴</h3>
                <article className="release-card"><div><strong>{APP_VERSION}</strong><span>共通項目の置換</span></div><ul><li>&lt;&lt;項目名&gt;&gt;による全文書共通の置換に対応</li><li>テンプレートから共通項目を自動検出</li><li>共通項目の入力・確認画面と未入力チェックを追加</li></ul></article>
                <article className="release-card release-card-previous"><div><strong>Ver.1.2.0</strong><span>設定とデータ除外の改善</span></div><ul><li>詳細設定を中央モーダルへ変更</li><li>記入例行の除外を選択可能に変更</li><li>除外なし、複数列のAND・OR条件を追加</li><li>除外理由と条件概要の表示を改善</li><li>数値のカンマ区切りとして名称と説明を整理</li></ul></article>
                <article className="release-card"><div><strong>Ver.1.1.0</strong><span>操作性・互換性の改善</span></div><ul><li>数字を含む置換値でもテンプレートのフォントを維持するよう修正</li><li>初回セットアップと文書処理をバックグラウンド化し、画面の応答性を改善</li><li>初回セットアップの所要時間案内を改善</li><li>準備中もテンプレート、置換データ、出力先を選択可能に変更</li><li>準備中に選択したファイルを、完了後に自動確認する機能を追加</li></ul></article>
                <article className="release-card release-card-previous"><div><strong>Ver.1.0.0</strong><span>初回正式版</span></div><ul><li>Word・PDFの個別、結合、ZIP出力に対応</li><li>Excel・CSV、ドラッグ＆ドロップ、進捗表示に対応</li><li>ファイル名設定、データ確認、ライト・ダークテーマを実装</li><li>金額列の3桁区切りと対象キーワード編集に対応</li></ul></article>
                <p className="help-footnote">更新履歴はVer.1.0.0以降を掲載します。</p>
              </div>
            )}

            {section === "about" && (
              <div className="help-section">
                <h3>このアプリについて</h3>
                <div className="about-card"><div className="about-symbol"><FileText size={24} /></div><div><strong>Wordファイル一括作成</strong><span>{APP_VERSION}</span></div></div>
                <div className="help-rule-list about-list">
                  <p><strong>制作者</strong><span>今井 啓登</span></p>
                  <p><strong>連絡先</strong><span><a href="mailto:ImaiK@mbox.pref.osaka.lg.jp">ImaiK@mbox.pref.osaka.lg.jp</a></span></p>
                </div>
                <div className="help-note"><strong>利用上の注意</strong><p>生成した文書は、配布・送信・印刷前に必ず内容を確認してください。</p></div>
              </div>
            )}
          </div>
        </div>
        <div className="help-actions"><button className="primary" onClick={onClose}>閉じる</button></div>
      </section>
    </div>
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
  const [settingsSection, setSettingsSection] = useState<"filename" | "exclude" | "numeric" | "pdf">("filename");
  const [helpOpen, setHelpOpen] = useState(false);
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
  const [templateWarmupState, setTemplateWarmupState] = useState<"idle" | "queued" | "running" | "ready" | "error">("idle");
  const [dataLoadState, setDataLoadState] = useState<"idle" | "queued" | "running" | "ready" | "error">("idle");
  const [amountKeywordInput, setAmountKeywordInput] = useState("");
  const [templateInspection, setTemplateInspection] = useState<TemplateInspection | null>(null);
  const [commonValues, setCommonValues] = useState<CommonValues>({});
  const [commonDraft, setCommonDraft] = useState<CommonValues>({});
  const [commonOpen, setCommonOpen] = useState(false);
  const [isFirstSetup] = useState(() => localStorage.getItem("wordBatchSetupCompleted") !== "1");
  const pendingTemplateRef = useRef<string | null>(null);
  const pendingDataRef = useRef<string | null>(null);

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
        if (classified.template_path) {
          setTemplatePath(classified.template_path); accepted.push("テンプレート"); setResult(null);
          if (warmupState === "ready") await prepareTemplate(classified.template_path);
          else { pendingTemplateRef.current = classified.template_path; setTemplateWarmupState("queued"); }
        }
        if (classified.output_path) { setOutputPath(classified.output_path); accepted.push("出力先"); }
        if (classified.data_path) {
          setDataPath(classified.data_path); accepted.push("置換データ"); setResult(null);
          if (warmupState === "ready") await inspect(classified.data_path);
          else { pendingDataRef.current = classified.data_path; setDataLoadState("queued"); }
        }
        if (classified.unsupported_paths.length) {
          setError(`対応していないファイル形式です。\n${classified.unsupported_paths.join("\n")}`);
        } else {
          setError("");
        }
        if (accepted.length) { setResult(null); setDropNotice(`${accepted.join("・")}を選択しました`); window.setTimeout(() => setDropNotice(""), 2400); }
      } catch (reason) { setError(String(reason)); }
    }).then((fn) => { unlisten = fn; });
    return () => { if (unlisten) unlisten(); };
  }, [busy, warmupState, settings.formatAmountWithComma, settings.rowExcludeMode, settings.rowExcludeColumns.join("|"), settings.excludeExampleRows, settings.targetColumnNumber]);

  useEffect(() => {
    let active = true;
    (async () => {
      try {
        await invoke("warmup_backend");
        if (active) {
          localStorage.setItem("wordBatchSetupCompleted", "1");
          setWarmupState("ready");
        }
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
    if (warmupState !== "ready") return;
    let cancelled = false;
    (async () => {
      const queuedTemplate = pendingTemplateRef.current;
      pendingTemplateRef.current = null;
      if (queuedTemplate && !cancelled) await prepareTemplate(queuedTemplate);
      const queuedData = pendingDataRef.current;
      pendingDataRef.current = null;
      if (queuedData && !cancelled) await inspect(queuedData);
    })();
    return () => { cancelled = true; };
  }, [warmupState]);

  useEffect(() => {
    if (!preview) return;
    setSettings((current) => ({
      ...current,
      filenameKeys: current.filenameKeys.filter((key) => preview.columns.includes(key)),
      rowExcludeColumns: current.rowExcludeColumns.filter((key) => preview.columns.includes(key)),
    }));
  }, [preview?.columns.join("|")]);

  const commonFields = templateInspection?.common_fields ?? [];
  const missingCommonFields = commonFields.filter((name) => !commonValues[name]?.trim());
  const commonCompletedCount = commonFields.length - missingCommonFields.length;
  const canRun = Boolean(
    templatePath && dataPath && outputPath && preview?.included_count && !busy && warmupState === "ready" && templateWarmupState !== "running" && missingCommonFields.length === 0 && (!(settings.rowExcludeMode.startsWith("selected_columns")) || settings.rowExcludeColumns.length > 0) && (settings.addSerialNumber || settings.filenameKeys.length > 0),
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
      const inspection = await invoke<TemplateInspection>("warmup_template", { templatePath: path });
      setTemplateInspection(inspection);
      setCommonValues({});
      setCommonDraft({});
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
      if (warmupState === "ready") await prepareTemplate(path);
      else { pendingTemplateRef.current = path; setTemplateWarmupState("queued"); }
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
      if (warmupState === "ready") await inspect(path);
      else { pendingDataRef.current = path; setDataLoadState("queued"); }
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
    setDataLoadState("running");
    setError("");
    setStatus("置換データを確認しています...");
    try {
      const response = await invoke<DataPreview>("inspect_data", {
        request: {
          data_path: path,
          format_amount_with_comma: settings.formatAmountWithComma,
          amount_include_keywords: settings.amountIncludeKeywords,
          row_exclude_mode: settings.rowExcludeMode,
          row_exclude_columns: settings.rowExcludeColumns,
          exclude_example_rows: settings.excludeExampleRows,
          row_exclude_target_column_number: settings.targetColumnNumber,
          output_format: settings.outputFormat,
          output_method: settings.outputMethod,
          fast_pdf_split_enabled: settings.fastPdfSplitEnabled,
        },
      });
      setPreview(response);
      setDataLoadState("ready");
      setStatus(`使用${response.included_count}件・除外${response.excluded_count}件`);
    } catch (reason) {
      setPreview(null);
      setDataLoadState("error");
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
          common_values: commonValues,
          add_serial_number: settings.addSerialNumber,
          serial_digits: settings.serialDigits,
          format_amount_with_comma: settings.formatAmountWithComma,
          amount_include_keywords: settings.amountIncludeKeywords,
          row_exclude_mode: settings.rowExcludeMode,
          row_exclude_columns: settings.rowExcludeColumns,
          exclude_example_rows: settings.excludeExampleRows,
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
    pendingTemplateRef.current = null;
    pendingDataRef.current = null;
    setTemplatePath("");
    setTemplateInspection(null);
    setCommonValues({});
    setCommonDraft({});
    setCommonOpen(false);
    setTemplateWarmupState("idle");
    setDataPath("");
    setDataLoadState("idle");
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

  function addAmountKeyword() {
    const keyword = amountKeywordInput.trim().replace(/\s+/g, "");
    if (!keyword) return;
    setSettings((current) => {
      if (current.amountIncludeKeywords.includes(keyword)) return current;
      return { ...current, amountIncludeKeywords: [...current.amountIncludeKeywords, keyword] };
    });
    setAmountKeywordInput("");
  }

  function removeAmountKeyword(keyword: string) {
    setSettings((current) => ({
      ...current,
      amountIncludeKeywords: current.amountIncludeKeywords.filter((item) => item !== keyword),
    }));
  }

  function resetAmountKeywords() {
    setSettings((current) => ({
      ...current,
      amountIncludeKeywords: [...DEFAULT_AMOUNT_INCLUDE_KEYWORDS],
    }));
    setAmountKeywordInput("");
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
            <div className="title-row"><h1>Wordファイル一括作成</h1><span className="version-badge">{APP_VERSION}</span></div>
            <p>テンプレートと置換データから、文書をすばやく作成します。</p>
          </div>
        </div>
        <div className="header-actions">
          <span className={`readiness ${warmupState}`}><i />{warmupState === "ready" ? "準備完了" : warmupState === "running" ? "準備中" : "要確認"}</span>
          <button className="icon-button" disabled={busy} onClick={() => setHelpOpen(true)} aria-label="使い方を確認" title="使い方を確認"><CircleHelp size={18} /></button>
          <button className="icon-button" onClick={() => setDark((value) => !value)} aria-label="テーマ切替" title="テーマを切り替え">
            {dark ? <Sun size={18} /> : <Moon size={18} />}
          </button>
          <button className="icon-button" disabled={busy} onClick={() => setSettingsOpen(true)} aria-label="設定" title="設定">
            <SettingsIcon size={18} />
          </button>
        </div>
      </header>

      <main>
        {warmupState === "running" && (
          <section className="warmup-card" aria-live="polite">
            <span className="warmup-spinner" />
            <div><strong>{isFirstSetup ? "初回セットアップを行っています" : "文書処理を準備しています"}</strong><p>{isFirstSetup ? "初回起動時のみ10～20秒ほどかかることがあります。準備中もテンプレート、置換データ、出力先を選択できます。" : "まもなく利用できます。準備中も各ファイルを選択できます。"}</p></div>
          </section>
        )}
        {warmupState === "error" && (
          <section className="error" role="alert"><strong>文書処理を準備できませんでした</strong><p>{warmupError}</p></section>
        )}
        <section className="picker-grid">
          <Picker kind="word" title="テンプレート" path={templatePath} disabled={busy} meta={templateWarmupState === "queued" ? "セットアップ完了後に確認します" : templateWarmupState === "running" ? "テンプレートを確認しています…" : templateWarmupState === "ready" ? "読み込みが完了しました" : undefined} onPick={chooseTemplate} />
          <Picker
            kind="excel"
            title="置換データ"
            path={dataPath}
            disabled={busy}
            meta={dataLoadState === "queued" ? "セットアップ完了後に確認します" : dataLoadState === "running" ? "置換データを確認しています…" : dataLoadState === "ready" ? "読み込みが完了しました" : undefined}
            onPick={chooseData}
          />
        </section>

        {commonFields.length > 0 && (
          <button className={`common-card ${missingCommonFields.length ? "incomplete" : "complete"}`} onClick={() => { setCommonDraft({ ...commonValues }); setCommonOpen(true); }} disabled={busy}>
            <span className="common-card-symbol">&lt;&lt;&gt;&gt;</span><span><strong>共通項目を入力</strong><small>{missingCommonFields.length ? `${commonFields.length}項目のうち${missingCommonFields.length}項目が未入力です` : `${commonFields.length}項目すべて入力済みです`}</small></span><em>{commonCompletedCount}/{commonFields.length}</em><ChevronRight size={18}/>
          </button>
        )}
        {templateInspection?.conflicting_fields?.length ? <section className="template-warning"><strong>同じ名前が行別項目と共通項目にあります</strong><p>{templateInspection.conflicting_fields.join("、")}。記号の指定が正しいか確認してください。</p></section> : null}
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
          <div className="settings-footer">
            <div className="filename-inline" title={templatePath ? exampleName : undefined}>
              <small>出力ファイル名の例</small>
              <code>{templatePath ? exampleName : "テンプレート選択後に表示します"}</code>
            </div>
            <button className="details" disabled={busy} onClick={() => setSettingsOpen(true)}>
              ファイル名と詳細設定 <ChevronRight size={17} />
            </button>
          </div>
        </section>

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
              <small>{warmupState === "running" ? (isFirstSetup ? "初回セットアップ中です" : "文書処理を準備しています") : status}</small>
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

      {helpOpen && <HelpGuide onClose={() => setHelpOpen(false)} />}
      {commonOpen && (
        <div className="overlay center-overlay" onMouseDown={() => setCommonOpen(false)}>
          <section className="common-modal" role="dialog" aria-modal="true" aria-labelledby="common-title" onMouseDown={(event)=>event.stopPropagation()}>
            <div className="drawer-head"><div><h2 id="common-title">共通項目の入力</h2><p>&lt;&lt;項目名&gt;&gt;へ、すべての文書で共通する文字を挿入します。</p></div><button className="icon-button" onClick={()=>setCommonOpen(false)}><X size={18}/></button></div>
            <div className="common-form">{commonFields.map(name=><label key={name}><span>{name}</span><input autoFocus={name===commonFields[0]} value={commonDraft[name] ?? ""} placeholder={`${name}を入力`} onChange={(e)=>setCommonDraft(current=>({...current,[name]:e.target.value}))}/></label>)}</div>
            <div className="common-actions"><button onClick={()=>setCommonOpen(false)}>キャンセル</button><button className="primary" disabled={commonFields.some(name=>!commonDraft[name]?.trim())} onClick={()=>{setCommonValues({...commonDraft});setCommonOpen(false);setResult(null);}}>適用</button></div>
          </section>
        </div>
      )}
      {settingsOpen && (
        <div className="overlay center-overlay" onMouseDown={() => setSettingsOpen(false)}>
          <section className="settings-modal" role="dialog" aria-modal="true" aria-labelledby="settings-title" onMouseDown={(event) => event.stopPropagation()}>
            <div className="drawer-head"><div><h2 id="settings-title">詳細設定</h2><p>変更内容は自動保存されます。</p></div><button className="icon-button" onClick={() => setSettingsOpen(false)}><X size={18} /></button></div>
            <div className="settings-layout">
              <nav className="settings-nav">
                <button className={settingsSection === "filename" ? "active" : ""} onClick={() => setSettingsSection("filename")}>ファイル名</button>
                <button className={settingsSection === "exclude" ? "active" : ""} onClick={() => setSettingsSection("exclude")}>行の除外</button>
                <button className={settingsSection === "numeric" ? "active" : ""} onClick={() => setSettingsSection("numeric")}>数値の整形</button>
                <button className={settingsSection === "pdf" ? "active" : ""} onClick={() => setSettingsSection("pdf")}>PDF</button>
              </nav>
              <div className="settings-content">
                {settingsSection === "filename" && <section className="setting-panel"><h3>ファイル名</h3>
                  <div className="serial-row"><label className="check"><input type="checkbox" checked={settings.addSerialNumber} disabled={settings.filenameKeys.length === 0} onChange={(event) => setSettings((c) => ({...c, addSerialNumber:event.target.checked || c.filenameKeys.length===0}))}/>先頭に通し番号を付ける</label><label className="digit-field">桁数<input type="number" min="1" max="6" disabled={!settings.addSerialNumber} value={settings.serialDigits} onChange={(e)=>setSettings(c=>({...c,serialDigits:Math.min(6,Math.max(1,Number(e.target.value)||1))}))}/></label></div>
                  {settings.filenameKeys.length===0 && <small className="setting-note">列が選択されていないため、通し番号は必須です。</small>}
                  <div className="setting-subhead"><strong>ファイル名に使用する列</strong><small>選択した順にファイル名へ追加します。</small></div>
                  {!preview ? <p className="muted-box">置換データを読み込むと列を選択できます。</p> : <><div className="column-choice-list">{preview.columns.map(column=><label className="check column-choice" key={column}><input type="checkbox" checked={settings.filenameKeys.includes(column)} disabled={settings.filenameKeys.includes(column)&&settings.filenameKeys.length===1&&!settings.addSerialNumber} onChange={()=>toggleFilenameKey(column)}/><span>{column}</span></label>)}</div><div className="sort-list">{settings.filenameKeys.map((key,index)=><div className="sort-item" key={key}><span className="sort-number">{index+1}</span><strong>{key}</strong><div><button disabled={index===0} onClick={()=>moveFilenameKey(index,-1)}><ChevronUp size={16}/></button><button disabled={index===settings.filenameKeys.length-1} onClick={()=>moveFilenameKey(index,1)}><ChevronDown size={16}/></button><button disabled={settings.filenameKeys.length===1&&!settings.addSerialNumber} onClick={()=>toggleFilenameKey(key)}><Trash2 size={16}/></button></div></div>)}</div></>}
                  <div className="drawer-preview"><small>出力例</small><code>{exampleName}</code></div>
                </section>}
                {settingsSection === "exclude" && <section className="setting-panel"><h3>行の除外</h3>
                  <label className="check example-toggle"><input type="checkbox" checked={settings.excludeExampleRows} onChange={(e)=>setSettings(c=>({...c,excludeExampleRows:e.target.checked}))}/>1列目に「例」を含む行を除外する</label><small className="setting-note">必要なデータにも「例」が含まれる場合はOFFにしてください。</small>
                  <label className="field-label">空欄による除外条件<select value={settings.rowExcludeMode} onChange={(e)=>setSettings(c=>({...c,rowExcludeMode:e.target.value as Settings["rowExcludeMode"]}))}><option value="none">何も除外しない</option><option value="any_empty_except_first">1列目以外に空欄があれば除外</option><option value="any_empty">どこかに空欄があれば除外</option><option value="all_empty_except_first">1列目以外がすべて空なら除外</option><option value="selected_columns_all_empty">指定列がすべて空なら除外</option><option value="selected_columns_any_empty">指定列のどれか1つでも空なら除外</option></select></label>
                  {settings.rowExcludeMode.startsWith("selected_columns") && <><div className="setting-subhead"><strong>除外判定に使用する列</strong><small>{settings.rowExcludeMode==="selected_columns_all_empty"?"選択した列がすべて空欄の場合に除外します。":"選択した列に1つでも空欄がある場合に除外します。"}</small></div>{!preview?<p className="muted-box">置換データを読み込むと列を選択できます。</p>:<div className="column-choice-list">{preview.columns.map(column=><label className="check column-choice" key={column}><input type="checkbox" checked={settings.rowExcludeColumns.includes(column)} onChange={()=>setSettings(c=>({...c,rowExcludeColumns:c.rowExcludeColumns.includes(column)?c.rowExcludeColumns.filter(x=>x!==column):[...c.rowExcludeColumns,column]}))}/><span>{column}</span></label>)}</div>}{settings.rowExcludeColumns.length===0&&<small className="validation-note">除外判定に使用する列を1つ以上選択してください。</small>}</>}
                  <div className="setting-summary"><strong>現在の除外条件</strong><span>{settings.excludeExampleRows?'記入例を除外':'記入例も使用'} / {settings.rowExcludeMode==='none'?'空欄による除外なし':settings.rowExcludeMode==='selected_columns_all_empty'?'指定列がすべて空なら除外':settings.rowExcludeMode==='selected_columns_any_empty'?'指定列のどれか1つでも空なら除外':settings.rowExcludeMode==='any_empty_except_first'?'1列目以外に空欄があれば除外':settings.rowExcludeMode==='any_empty'?'どこかに空欄があれば除外':'1列目以外がすべて空なら除外'}</span></div>
                </section>}
                {settingsSection === "numeric" && <section className="setting-panel"><h3>数値の整形</h3><label className="check"><input type="checkbox" checked={settings.formatAmountWithComma} onChange={(e)=>setSettings(c=>({...c,formatAmountWithComma:e.target.checked}))}/>数値をカンマ区切りにする</label><small className="setting-note">登録したキーワードを列名に含む列の値を、カンマ区切り形式へ整形します。</small><div className={settings.formatAmountWithComma?"amount-keyword-editor":"amount-keyword-editor disabled"}><div className="setting-subhead amount-keyword-heading"><div><strong>対象となる列名キーワード</strong><small>部分一致で判定します。</small></div><button disabled={!settings.formatAmountWithComma} onClick={resetAmountKeywords}>初期設定に戻す</button></div><div className="keyword-tags">{settings.amountIncludeKeywords.length===0?<span className="keyword-empty">対象キーワードがありません。</span>:settings.amountIncludeKeywords.map(keyword=><span className="keyword-tag" key={keyword}>{keyword}<button disabled={!settings.formatAmountWithComma} onClick={()=>removeAmountKeyword(keyword)}><X size={13}/></button></span>)}</div><div className="keyword-add-row"><input type="text" value={amountKeywordInput} disabled={!settings.formatAmountWithComma} placeholder="例：補助対象経費" onChange={(e)=>setAmountKeywordInput(e.target.value)} onKeyDown={(e)=>{if(e.key==='Enter'){e.preventDefault();addAmountKeyword();}}}/><button disabled={!settings.formatAmountWithComma||!amountKeywordInput.trim()} onClick={addAmountKeyword}>追加</button></div></div></section>}
                {settingsSection === "pdf" && <section className="setting-panel"><h3>PDF</h3><label className="check"><input type="checkbox" checked={settings.fastPdfSplitEnabled} onChange={(e)=>setSettings(c=>({...c,fastPdfSplitEnabled:e.target.checked}))}/>個別PDF出力を高速化する</label><small className="setting-note">分割できない場合は自動的に通常方式へ切り替えます。</small></section>}
              </div>
            </div>
            <div className="settings-actions"><button onClick={()=>setSettingsOpen(false)}>キャンセル</button><button className="primary" disabled={settings.rowExcludeMode.startsWith("selected_columns")&&settings.rowExcludeColumns.length===0} onClick={async()=>{setSettingsOpen(false);if(dataPath)await inspect();}}>設定を適用</button></div>
          </section>
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
            {commonFields.length > 0 && <div className="common-preview"><strong>共通項目</strong>{commonFields.map(name=><span key={name}><b>{name}</b><em>{commonValues[name] || "未入力"}</em></span>)}</div>}<div className="exclusion-summary"><strong>現在の除外条件</strong><span>{preview.exclusion_summary || "設定なし"}</span></div><div className="table-help">表の下部にあるスクロールバーで、右側の列まで確認できます。</div>
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
