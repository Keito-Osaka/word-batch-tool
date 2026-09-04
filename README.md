# Wordファイル一括作成ツール

Quiet型UIとPython文書生成エンジンを組み合わせたTauri 2アプリの第1段階版です。

## 現在の対応範囲
- Wordテンプレート選択
- Excel/CSV選択とプレビュー
- 行除外設定
- 個別Word生成
- 通し番号、ファイル名列、金額カンマ
- GitHub ActionsによるWindows NSISビルド

画面にはPDF、結合、ZIPの選択肢がありますが、第1段階では個別Word生成のみ接続済みです。

## GitHub Actionsでのビルド
1. このフォルダの中身を非公開リポジトリ直下へアップロード
2. GitHubの Actions タブで `Windows build` を選択
3. `Run workflow` を実行
4. 完了後、Artifactsの `WordBatchTool-Windows` をダウンロード

## テスト
`test-data`には架空データだけを収録しています。実在する個人情報はリポジトリへ追加しないでください。
