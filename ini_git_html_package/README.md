# INI 設定更新ツール（BAT + 共通VBS v3）

**2026-10-20 更新** / Windows Script Host・ADODB.Stream を使用。PowerShell不要。

## GitHub Pages での閲覧

- `index.html`：使い方とBAT・VBSのコード全文
- `set_test_7777DB.html`：BAT全文
- `set_ini_test.html`：VBS全文

HTMLは閲覧・コピー・ダウンロード用です。INIの更新はWindowsでBATを実行してください。

## 実行ファイル

- `set_test_7777DB.bat`：INIパス、対象セクション、キー名、置換値、文字コードを設定
- `set_ini_test.vbs`：BATの指定に従って処理する共通VBS。単体で再利用可能

現在のBAT設定：

```bat
set "INI_FILE=%~dp0test.ini"
set "TARGET_SECTION=\u30BB\u30C3\u30C8"
set "TARGET_KEY=case"
set "NEW_VALUE=99DB"
set "INI_CHARSET=utf-8"
```

上記の `TARGET_SECTION` は **[セット]**。このセクションでコメントアウトされていない行の `case=` の値を **99DB** にします。別のINIや値に変更するときはBATの設定を編集するだけです。

BATのファイル名は旧名称のままですが、実際の置換値は `NEW_VALUE` に従います。

既存のBATから2～3引数で呼び出した場合は、互換動作として **[設定1] の test=** を更新します。新形式は5引数です。

## 注意

- バックアップファイルは**自動作成しません**。本番適用前にINIをコピーして検証してください。
- 日本語のセクション名は `\uXXXX` を組み合わせてBATに指定できます。これによりBATの文字コード差による文字化けを避けます。
- `INI_CHARSET` は元ファイルと一致させてください（`utf-8`、`shift_jis`、`unicode`）。
- BATとVBSは同じフォルダーに配置します。対象INIには書込権限が必要です。
- 新コードは静的に照合済みですが、職場のAWS/Windowsでの実行試験は未実施です。
- 公開リポジトリに会社固有のパス・認証情報などを載せないでください。
