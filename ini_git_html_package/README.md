# INI updater v5 — 全角Database名を直接書けるBAT雛形

2026-10-22。BATと共通VBSを両方とも差し替えてください。職場で重ねた部分修正・診断用MsgBox・WScript.Quitの追加は不要です。

## ファイル

- `set_A2.bat`: `NEW_VALUE=Ａ２テストＤＢ`
- `set_B7.bat`: `NEW_VALUE=Ｂ７テストＤＢ`
- `set_test_7777DB.bat`: 既存名を維持。設定値は従来どおり99DB。
- `set_ini_test.vbs`: 全BAT共通のv5。
- `index.html`: 使い方とソース閲覧ページ。
- `.txt`: UTF-8の閲覧用ソース。実行用BATの代用ではありません。

実行用BATはShift_JIS（CP932）・CRLFで保存済み。編集後も同じ形式で保存します。VBSはASCII・CRLF。BATとVBSを同じフォルダーに置き、INI_FILEを実際のパスに変更してください。

## パターンの増やし方

BATをコピーして `set "NEW_VALUE=Ｂ７テストＤＢ"` のように名前をそのまま変更します。値のUnicodeエスケープ化は不要です。INIの場所などが同じなら、パターン別に変えるのはNEW_VALUEだけです。

## 修正内容

v5は `--env` 指定で、プロセス環境変数を名前で受け取ります。INI_FILE・NEW_VALUE・INI_CHARSET・TARGET_SECTION・TARGET_KEYが対象です。５番目の位置引数への依存を解消し、必須値の空チェックは維持しました。旧BATの2・3・5引数形式との互換性も残しています。

今回の対象は `[ODBC接続文字列]` 内のコメントアウトされていない２行で、各行最後のDatabaseの値です。両方を同じNEW_VALUEにします。EXPECTED_MATCHES=2が既定で、２件そろわなければ書き込みません。意図的に件数を変える場合だけ変更してください（0は制限なし）。

`set_A2.bat /check` は読取・解析だけを行い、書き込みません。`/nopause` を付けると停止待ちを省略します。２件とも変更すると `Matched: 2 value(s).` と `Updated: 2 value(s).` を表示します。

## 注意

バックアップは自動生成しません。最初はINIのコピーで確認してください。検証用の一時ファイルは処理後に削除しますが、書き込み途中の障害時に元へ戻す保証はありません。

１設定１行・Databaseが最後の属性という形式が対象です。接続文字列全体の引用符囲み、複数行、Database以降の別属性・行末コメントは対象外です。更新値の空白・制御文字・引用符・波括弧・セミコロン・#・BATで展開される%を使う例は対象外です。

短いエラー識別子: E011=INI_FILE、E012=NEW_VALUE、E013=TARGET_SECTION、E014=TARGET_KEY、E102=文字コード、E201/E202=対象なし、E204=件数不一致。

`tests/test_windows.py` はWindowsで実際のBATとVBSを実行する15件の試験です。架空のINIを一時フォルダーへ作り、全角名・2行更新・コメント保持・チェックモード・失敗時の未更新を確認します。最新結果はGitHub Actionsの `INI updater Windows tests` を参照してください。職場AWS環境の実行確認とは別です。
