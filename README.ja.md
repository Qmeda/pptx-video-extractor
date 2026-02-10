# PPTX Video Extractor

PPTX（PowerPoint）ファイルに埋め込まれた動画を抽出する VS Code 拡張です。

## 機能

- エクスプローラー上で `.pptx` を右クリックして抽出
- GUI で出力先フォルダを選択
- 同名ファイルは ` (1)`, ` (2)` のように連番付与して回避
- `ppt/media` 内の元ファイル名を保持

## 必要要件

- VS Code 1.80 以降
- Node.js（ローカルビルド用）

## 使い方

1. プロジェクトフォルダを VS Code で開く
2. `F5` で拡張機能の開発ホストを起動
3. 開発ホスト内で `.pptx` を右クリック
4. **Extract Videos from PPTX** を選択
5. 出力先フォルダを選択

## ビルド

```bash
npm install
npm run compile
```

## 補足

- PPTX は ZIP 形式です。動画は `ppt/media/` 配下に保存されています。
- 一般的な動画拡張子のみ抽出します。

## ライセンス

ライセンスをここに記載してください。