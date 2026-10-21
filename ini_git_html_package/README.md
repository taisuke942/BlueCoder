# INI updater v4 — ODBC Database更新

2026-10-21 更新。先に配布したv4の共通VBSを反映しています。

## 今回の設定

```bat
set "INI_FILE=%~dp0test.ini"
set "TARGET_SECTION=ODBC\u63A5\u7D9A\u6587\u5B57\u5217"
set "TARGET_KEY=Database"
set "NEW_VALUE=99DB"
set "INI_CHARSET=shift_jis"
```

`INI_FILE` は実際のINIパス、`NEW_VALUE` は希望の共通Database名に変更してください。
`TARGET_SECTION` は [ODBC接続文字列] を指定します（角括弧は引数に含めません）。
ファイル名は従来の `set_test_7777DB.bat` のままですが、実際の値は `NEW_VALUE` に従います。

## 2行の接続情報

指定セクション内で、コメントアウトされていない各行の最後の項目が `Database=` なら、その値だけを共通の `NEW_VALUE` に変更します。2行ある場合は2行とも処理します。
2行限定ではなく、同じ条件の有効行が3行あれば3行とも対象です。
`;` または `#` で始まるコメント行（先頭の空白・タブも許容）、別セクション、Database以外の項目は保持する処理です。
2か所とも書き換われば `Updated: 2 value(s).` と表示します。既に同じ値の箇所は変更件数に含めません。

## 閲覧と実行

- [使い方・BAT・VBS全文](index.html)
- [BATソース](set_test_7777DB.html)
- [VBSソース](set_ini_test.html)

詳細ページの旧URLは、一覧ページ内の最新ソースへ移動します。コードはHTML内に直接収録しているので、閲覧に外部ライブラリやソース取得処理は不要です。

実行用のBATと `set_ini_test.vbs` は同じフォルダーに置きます。PowerShellは使用しません。
開始時の表示は `INI updater v4 (ODBC Database tail)` です。
GitHubの更新だけでは職場のローカルファイルは更新されません。BAT・VBSの両方を反映してください。

## 対応範囲と注意

接続情報が1行で、Databaseが最後の項目である形式向けです。末尾セミコロンの有無に対応します。
接続文字列全体を引用符で囲む形式、複数行形式、Databaseより後に別項目がある形式は対象外です。
新しいDatabase名に空白・制御文字・引用符・波括弧・セミコロン・#は使用できません。
Database以外のキーは従来処理を維持します。2～3引数で呼ぶ旧BATは [設定1] の test= を対象とする互換動作になります。

**バックアップは自動生成しません。最初はINIのコピーで試してください。**
Shift_JISは `INI_CHARSET=shift_jis` で読み書きします。文字コード往復変換と更新直前・書き込み後のバイト検証を維持しています。
WindowsのVBS本体・ADODB.Stream・職場AWS環境での通し試験は未実施です。
スクリプトはサーバーへ接続せず、SQLも実行しません。
職場のINI・実際の接続先・認証情報はこのフォルダーに含めていません。

## エラー識別子

E004: セクション名のUnicodeエスケープ
E101: INIパス
E102: 文字コード・元バイト保持
E201: 指定セクションなし
E202: 対象となる末尾Databaseなし
E203: 接続情報の括弧・引用符の不整合

識別子が表示された場合は、実際の接続情報を送らずに切り分けできます。
