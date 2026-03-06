# shinsa

`SPEC.md`準拠で組み直した、交付金審査向けのローカル作業ツールです。

## 起動

```
shinsa.bat
```

- 管理者権限不要
- GUI が直接起動します

## ローカル永続ファイル

`config.local.json` 以外のアプリデータは `paths.json_root` 配下にまとまります。

- ソースごとの JSON（`anken.json`, `kenshu.json`, `mails.json`, `folders.json` など）
- `cache.json`（UI 状態）

## 初期設定

1. `config\config.local.sample.json` を `config\config.local.json` にコピー
2. 各ソースの `source_path` を実環境に合わせて編集
3. Excel 台帳を使う場合は `source_table`（構造化テーブル名）と `columns`（列マッピング）を設定
4. `shinsa.bat` を起動（初回は自動 sync）

## データモデル

- GUI 起動時に JSON が無ければ自動で sync します
- sync は各ソースの `source_path` から JSON を再生成します
- `cache.json` は保持され、UI 状態を持ちます
- writeback は編集済みフィールドを正本に反映します

## ソース構成

`config.base.json` の `sources` で定義します。

- **主軸ソース**（`join` なし）: ComboBox で切替。レコード一覧と詳細を表示
- **従属ソース**（`join` あり）: 主軸レコード選択時にタブで自動表示
- `join` の `match_mode` で `exact`（完全一致）/ `domain`（ドメイン一致）を切替可能
- `contact_email` はセミコロン区切りで複数アドレス登録可能

## サンプル設定

サンプルの `config.local.sample.json` は次を指します。

- `data/sample/onedrive/table/anken.source.json`（案件台帳：8件、mail/folder が join）
- `data/sample/onedrive/table/kenshu.source.json`（研修管理台帳：5件、単独）
- `data/sample/mail`（メール：8通）
- `data/sample/onedrive/cases`（案件フォルダ：6案件分）

そのまま起動して動作確認できます。

## 仕様と補助資産

- 全体仕様: `SPEC.md`
- Outlook VBA エクスポート例: `VBA\ShinsaOutlookExport.bas`
