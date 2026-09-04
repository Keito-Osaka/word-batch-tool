# 第2段階版への更新手順

このZIPには変更対象ファイルだけが入っています。GitHubリポジトリの同じ場所へ上書きしてください。

正式アイコンを維持するため `src-tauri/icons` は含めていません。

上書き対象:
- src/App.tsx
- src/styles.css
- src/types.ts
- src-tauri/src/lib.rs
- src-tauri/capabilities/default.json
- python/core.py
- python/bridge.py
- python/bridge.spec
- python/requirements.txt
- .github/workflows/build-windows.yml

上書き後、新しいコミットから Actions の Windows build を手動実行してください。
