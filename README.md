# goobike-auto

グーバイク自動登録・GAS

## プロジェクト構成

- `gas/` Google Apps Script のソースコード
  - `Code.gs` スプレッドシートの保護・差し戻しロジック
  - `Sidebar.html` 自動更新パネル（サイドバー）用 UI
- `pages/` Next.js を利用した補助用のウェブ UI
- `__tests__/` テストコード（Jest）

## 開発コマンド

```bash
npm install
npm run dev
npm run lint
npm test -- --passWithNoTests
```
