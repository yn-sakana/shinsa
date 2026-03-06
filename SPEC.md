# shinsa 仕様書

## 1. 目的

- Outlook で受信した申請メールと添付をローカルに保存する
- SharePoint 上の案件台帳と案件フォルダは OneDrive 同期済みローカルを通じて扱う
- 正本を直接の作業場にせず、ローカルに JSON 化したデータを GUI で閲覧・編集する
- 正本への反映は bridge 処理で明示的に行う

## 2. 基本原則

- 正本は 3 つだけとする
  - Outlook
  - SharePoint 案件台帳
  - SharePoint 案件フォルダ
- SharePoint へは VBA も PowerShell も直接アクセスしない
- SharePoint の入口は OneDrive 同期済みローカルパスとする
- Outlook の入口は Outlook VBA が作るローカル Mail Archive とする
- `shinsa` は正本を直接編集しない
- `shinsa` は JSON 化したデータと cache だけを扱う
- 業務項目は台帳に存在する列だけを扱う
- GUI が業務列を勝手に増やさない
- アプリ内でしか使わない状態は `cache.json` にだけ持つ

## 3. 全体アーキテクチャ

```text
Outlook
  │
  │ Outlook VBA
  ▼
Mail Archive
  │
  │
  ├──────────────┐
  │              │
  ▼              ▼
SharePoint --OneDrive--> OneDrive同期済みローカル
                     - 案件台帳
                     - 案件フォルダ

Mail Archive + OneDrive同期済みローカル
                    │
                    ▼
                 shinsa
          - JSON化
          - GUI表示
          - cache管理
                    │
                    ▼
      ledger.json / mails.json / folders.json / cache.json
                    │
                    ▼
                   GUI
          - 閲覧
          - 編集
          - 正本を開く
                    │
                    ▼
          bridge (明示操作のみ)
      - Excel VBA: 台帳反映
      - Outlook VBA: 下書き作成
      - PowerShell: 案件フォルダ反映
```

### 3.1 技術スタック

- ランチャー
  - `bat`
  - 正式入口は `shinsa.bat`
- アプリ本体
  - `PowerShell`
- GUI
  - `WinForms`
- Outlook 連携
  - `Outlook VBA`
- 台帳反映
  - `Excel VBA`
- 案件フォルダ反映
  - `PowerShell`
- SharePoint 入口
  - `OneDrive 同期済みローカル`
- データ形式
  - `JSON`
- 正本台帳
  - `Excel`

### 3.2 起動要件

- 日常の正式入口は `shinsa.bat` だけとする
- `shinsa.bat` は管理者権限不要で起動できることを要件とする
- 通常利用で UAC 昇格を要求しない
- PowerShell プロファイルへの追記を前提にしない
- 右クリックメニュー登録やレジストリ変更を前提にしない
- ランチャーは最小に保ち、業務ロジックを持たせない
- `shinsa.bat` は `powershell.exe -ExecutionPolicy Bypass -File ...` で本体を起動する
- 利用者は長い引数列を覚える必要がないことを要件とする

## 4. 操作モデル

### 4.1 入口

- 日常の入口は `shinsa.bat` だけとする
- 起動後は対話型シェルに入る
- ユーザーは長い引数を毎回打たない
- 単語 1 つのコマンドで操作する

### 4.2 対話型シェル

- `shinsa>` プロンプトを表示する
- 想定コマンド
  - `gui`
  - `sync`
  - `status`
  - `writeback`
  - `help`
  - `quit`
- GUI は別プロセスで起動し、GUI 使用中でもシェル側で別コマンドを打てる
- シェルは daily driver として使える軽さを優先する

### 4.3 GUI とシェルの役割分担

- シェル
  - 同期
  - 状態確認
  - bridge 起動
  - 補助操作
- GUI
  - 検索
  - 閲覧
  - 編集
  - 正本を開く

## 5. GUI 方針

### 5.1 見た目

- GUI は Win95 / Win98 系の業務アプリ風を基調とする
- 情報密度を高くする
- 装飾より可読性と操作速度を優先する
- ボタン、一覧、入力欄、分割ペイン中心の構成にする
- フラットすぎる Web 風 UI にはしない

### 5.2 必須体験

- 一覧から案件を素早く選べる
- 右側または下側で詳細をすぐ見られる
- メール、添付、案件フォルダ、台帳を `開く` で直接開ける
- 画面が見切れない
- 低解像度でも最低限操作できる
- キーボード操作しやすい

### 5.3 主画面構成

- 左または上
  - 案件一覧
- 右または下
  - 台帳項目の詳細
  - 紐付いたメール一覧
  - 案件フォルダファイル一覧
  - 編集欄
- 上部
  - 検索
  - 同期
  - 書き戻し
  - ステータス表示

## 6. ローカルで管理するファイル

ローカルで永続化するのは次の 5 つだけとする。

1. `ledger.json`
2. `mails.json`
3. `folders.json`
4. `config.local.json`
5. `cache.json`

補足
- `ledger.json` `mails.json` `folders.json` はそれぞれ正本の JSON 化データ
- `cache.json` はアプリ内状態だけを持つ
- GUI の統合ビューはメモリ上で作る
- `cases.json` のような統合 JSON は持たない
- `index.json` のような専用 index も持たない

## 7. 正本とローカル入口

### 7.1 Outlook

- 正本: Outlook メールボックス
- ローカル入口: Outlook VBA が出力する Mail Archive
- `shinsa` は Outlook を直接読まない

### 7.2 SharePoint 案件台帳

- 正本: SharePoint 上の Excel 台帳
- ローカル入口: OneDrive 同期済みローカルの台帳ファイル
- `shinsa` はこのローカル台帳を読む

### 7.3 SharePoint 案件フォルダ

- 正本: SharePoint 上の案件フォルダ
- ローカル入口: OneDrive 同期済みローカルの案件フォルダ
- `shinsa` はこのローカル案件フォルダを読む

## 8. JSON ごとの責務

### 8.1 ledger.json

役割
- 案件台帳そのものの JSON 化データ
- 台帳列だけを持つ
- GUI から編集してよいのは、台帳にも存在する列だけ

保持するもの
- 台帳の主キー列
- 団体名
- 申請担当者名
- 申請担当者メールアドレス
- 審査ステータス
- 担当者
- 不足書類
- 公開用メモ
- その他、台帳に実在する列

保持しないもの
- メール手動紐付け
- GUI 独自の進捗
- ウィンドウ状態
- 検索条件

主な列例
- `case_id`
- `receipt_no`
- `organization_name`
- `contact_name`
- `contact_email`
- `status`
- `assigned_to`
- `missing_documents`
- `review_note_public`
- `ledger_path`
- `ledger_sheet`
- `ledger_row_id`

### 8.2 mails.json

役割
- Mail Archive の JSON 化データ
- メールそのものの事実だけを持つ
- 1 メール 1 レコード

保持するもの
- 差出人
- 件名
- 受信日時
- 本文パス
- msg パス
- 添付パス
- Outlook フォルダパス

保持しないもの
- 手動で紐付けた案件 ID
- GUI 独自の進捗
- 一時メモ

主な列例
- `mail_id`
- `entry_id`
- `mailbox_address`
- `folder_path`
- `received_at`
- `sender_name`
- `sender_email`
- `subject`
- `body_path`
- `msg_path`
- `attachment_paths`

### 8.3 folders.json

役割
- 案件フォルダの JSON 化データ
- 案件フォルダ内のファイル事実だけを持つ
- 1 ファイル 1 レコードを基本とする

保持するもの
- 案件 ID
- フォルダパス
- ファイルパス
- 相対パス
- 更新日時
- サイズ

保持しないもの
- 採用ファイル選択状態
- GUI 独自のタグ
- publish フラグ

主な列例
- `case_id`
- `folder_path`
- `file_path`
- `relative_path`
- `file_name`
- `extension`
- `modified_at`
- `size`

### 8.4 config.local.json

役割
- ローカル実行設定
- 自分の送信アカウント設定
- OneDrive パス設定
- Mail Archive パス設定
- 台帳列名設定

主な項目
- `paths.mail_archive_root`
- `paths.sharepoint_ledger_path`
- `paths.sharepoint_case_root`
- `paths.json_root`
- `mail.self_address`
- `ledger.key_column`
- `ledger.sheet_name`
- `ledger.columns.*`

### 8.5 cache.json

役割
- アプリ内でしか使わない状態だけを持つ
- 削除しても正本や JSON 化データ自体は壊れない

保持するもの
- メールの手動紐付け
- GUI 独自の進捗
- 検索条件
- ソート順
- 最後に開いた案件
- ウィンドウ位置やペイン幅

保持しないもの
- 台帳に戻す値
- メールそのものの事実
- 案件フォルダそのものの事実

構成例
- `mail_links`
- `mail_progress`
- `ui_state`

## 9. 紐付けルール

### 9.1 強い紐付け

- `ledger.json` と `folders.json` は `case_id` で確実に紐付く
- `case_id` が無い場合は、仕様上の主キー列を `case_id` に正規化して扱う

### 9.2 弱い紐付け

- `mails.json` と `ledger.json` は `contact_email` と `sender_email` で候補を出す
- これは確実ではない
- 自動候補は GUI で表示するだけに留める
- 手動補正結果は `cache.json` に持つ

### 9.3 GUI 上の結合

- `ledger.json` を主テーブルとして扱う
- `folders.json` は `case_id` で結合する
- `mails.json` は `cache.json.mail_links` を優先し、無ければメールアドレス一致で候補表示する
- 統合結果はメモリ上で作るだけで、永続化しない

## 10. sync の仕様

sync は次を行う。

1. Outlook VBA が更新した Mail Archive を読む
2. OneDrive 同期済み台帳を読む
3. OneDrive 同期済み案件フォルダを読む
4. `ledger.json` を再生成する
5. `mails.json` を再生成する
6. `folders.json` を再生成する
7. `cache.json` は保持する

重要
- `cache.json` は sync で消さない
- 手動紐付けや GUI 独自進捗は `cache.json` に残る
- `ledger.json` `mails.json` `folders.json` は source から作り直してよい

## 11. GUI の仕様

### 11.1 基本方針

- GUI は `ledger.json` `mails.json` `folders.json` `cache.json` を読み込む
- GUI はメモリ上で突合して表示する
- GUI は統合 JSON を別保存しない

### 11.2 主画面

- 主テーブルは `ledger.json`
- 台帳の案件行を一覧表示する
- 選択案件に対して以下を表示する
  - 紐付いたメール一覧
  - 案件フォルダ内ファイル一覧
  - 台帳列の編集欄

### 11.3 編集できるもの

- 台帳に実在する列
- `cache.json` に属するアプリ内状態

具体例
- 台帳列の編集
  - `status`
  - `assigned_to`
  - `missing_documents`
  - `review_note_public`
- cache の編集
  - メール手動紐付け
  - GUI 独自進捗

### 11.4 開く操作

GUI から次を直接開けるようにする。

- `mail.msg`
- メール本文ファイル
- 添付ファイル
- 案件フォルダ
- 台帳ファイル

## 12. bridge / writeback の仕様

### 12.1 台帳反映

- 実処理は Excel VBA bridge
- 対象は OneDrive 同期済みローカル台帳
- `ledger.json` と実際の台帳を比較して反映する
- 変更確認は VBA 側で行う
- 反映対象は台帳に存在する列だけ
- GUI 独自項目は反映しない

### 12.2 案件フォルダ反映

- 実処理は PowerShell
- 対象は OneDrive 同期済みローカル案件フォルダ
- どのファイルを配置するかは GUI と cache を見て判断する
- 破壊的上書きは避ける

### 12.3 Outlook 下書き作成

- 実処理は Outlook VBA bridge
- 送信元は `config.local.json` の自分のアドレスを使う
- 宛先は `ledger.json` の相手先アドレスを使う
- 自動送信はしない
- 下書き作成までを責務とする

## 13. 技術責務

### 13.1 Outlook VBA

- Outlook からメールと添付を Mail Archive に保存する
- 必要に応じて返信下書きを作成する
- SharePoint には直接アクセスしない

### 13.2 PowerShell

- Mail Archive と OneDrive 同期済みローカルを読む
- `ledger.json` `mails.json` `folders.json` を生成する
- GUI を起動する
- 案件フォルダ反映を行う

### 13.3 Excel VBA bridge

- `ledger.json` と台帳を比較する
- 変更確認を出す
- 台帳列だけ反映する
- SharePoint には直接アクセスせず、OneDrive ローカル台帳を扱う

### 13.4 GUI

- JSON を読み込んで統合表示する
- 台帳列を編集する
- cache を更新する
- 正本ファイルを開く

## 14. 代表ユースケース

### 14.1 新規申請の確認

1. Outlook VBA が申請メールを Mail Archive に保存する
2. OneDrive が台帳と案件フォルダを同期する
3. ユーザーが `shinsa` を起動する
4. `sync` で JSON を更新する
5. GUI で案件一覧を確認する
6. 該当案件のメール、添付、案件フォルダ、台帳を開いて確認する

### 14.2 メールの手動紐付け

1. 自動候補だけでは案件が確定しないメールを GUI で開く
2. 対象案件をユーザーが選ぶ
3. 紐付け結果を `cache.json` に保存する
4. 次回以降の GUI 表示でもその紐付けを優先する

### 14.3 台帳更新

1. GUI で台帳列を編集する
2. `writeback` を実行する
3. Excel VBA bridge が OneDrive ローカル台帳との差分を確認する
4. ユーザーが確認後、限定列だけ反映する
5. OneDrive が SharePoint に同期する

### 14.4 返信下書き作成

1. GUI で案件を開く
2. 必要なメールや添付を確認する
3. bridge を通じて Outlook VBA で下書きを作成する
4. ユーザーが Outlook で内容を確認する

## 15. 未決事項

- 台帳の主キーを `case_id` と `receipt_no` のどちらに揃えるか
- 申請担当者アドレス帳が台帳内包か別シートか別ファイルか
- 案件フォルダ反映時の退避ルールをどうするか
- GUI 独自進捗の状態一覧をどう定義するか
- Mail Archive の保存階層をどこまで Outlook フォルダ構造に寄せるか
