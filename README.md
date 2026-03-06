# shinsa

交付金審査用のローカル作業ツールです。

## 入口

```powershell
.\shinsa.bat
```

- 管理者権限不要
- 起動後は `shinsa>` の対話モードに入ります
- GUI は `gui` コマンドで別プロセス起動します

## 主要コマンド

- `gui`
- `sync`
- `index`
- `writeback`
- `status`
- `config`
- `help`
- `quit`

## 初期設定

1. `config\config.local.sample.json` を `config\config.local.json` にコピー
2. OneDrive と Mail Archive のパスを環境に合わせて編集
3. `shinsa.bat` を起動

## 仕様

- 全体仕様: `SPEC.md`
- Outlook 連携: `VBA\ShinsaOutlookExport.bas`
