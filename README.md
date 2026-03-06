# shinsa

`SPEC.md`準拠で組み直した、交付金審査向けのローカル作業ツールです。

## 入口

```powershell
.\shinsa.bat
```

- 管理者権限不要
- `powershell.exe -ExecutionPolicy Bypass -File Main.ps1` で起動
- 日常操作は `shinsa>` の対話シェルから行う

## シェルコマンド

- `gui`
- `sync`
- `status`
- `writeback`
- `help`
- `quit`

## ローカル永続ファイル

`config.local.json` 以外のアプリデータは `paths.json_root` 配下にまとまります。

- `ledger.json`
- `mails.json`
- `folders.json`
- `cache.json`

## 初期設定

1. `config\config.local.sample.json` を `config\config.local.json` にコピー
2. `mail_archive_root` `sharepoint_ledger_path` `sharepoint_case_root` を実環境に合わせて編集
3. Excel台帳を読む場合は `ledger.columns.*` を台帳の実列名に合わせて編集
4. `.\shinsa.bat` を起動し、最初に `sync` を実行

## データモデル

- `sync` は Mail Archive と OneDrive 同期済みローカルから `ledger.json` `mails.json` `folders.json` を再生成します
- `cache.json` は保持され、手動メール紐付けとUI状態を持ちます
- GUI は4つのJSONを読み込み、メモリ上で突合して表示します
- `writeback` は `ledger.json` と正本台帳との差分だけを反映します

## サンプル設定

サンプルの `config.local.sample.json` は次を指します。

- `data/sample/mail`
- `data/sample/onedrive/ledger/ledger.source.json`
- `data/sample/onedrive/cases`

そのまま `sync` して動作確認できます。

## 仕様と補助資産

- 全体仕様: `SPEC.md`
- Outlook VBA エクスポート例: `VBA\ShinsaOutlookExport.bas`
