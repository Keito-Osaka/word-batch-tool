# Wordファイル一括作成 Ver.3.0.0

WordテンプレートとExcel/CSVからWord・PDFを一括作成するWindowsアプリです。

## Ver.3.0.0
- Pythonバックエンド、Native Messagingブリッジ、固定ID版Edge拡張機能を単一のNSISインストーラーへ統合
- PowerShellによる導入操作を廃止
- アプリ内にMicrosoft Edge連携設定、接続確認、修復機能を追加
- 状態表示を「連携設定」「最終応答」に統一
- 接続確認時は対象文書を読み取らず、診断通信だけを実行

## ビルド
GitHub Actionsの `Windows build` を実行します。Artifact名は `WordBatchTool-Windows-v3.0.0` です。
