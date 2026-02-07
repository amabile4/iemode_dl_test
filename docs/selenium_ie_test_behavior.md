Selenium IE Mode テストの挙動まとめ
================================

対象: `automation/selenium_ie_test.py`

目的
----
Edge IEモード + IEDriver + pywinauto で、
ログイン → ダウンロード → 保存 → 上書き確認 → 保存完了確認を一連で実行する。
移植時に必要な「待ち条件」「ダイアログの見分け」「操作の分岐」を明示する。

前提
----
- Flaskサーバーが `http://localhost:5000` で稼働
- IEDriverServer.exe のパスが有効
- Edge が IEモード起動可能
- pywinauto を使用（UIA / Win32）

起動〜終了の流れ
--------------
1. 既存の IEモード Edge を終了
2. IEDriver + Edge IEモードを起動
3. ログインページで ID/Password 入力 → Enter でログイン
4. ダウンロードリンク実行 → confirm ダイアログを Win32 で OK
5. ダウンロードバーで「名前を付けて保存」を選択
6. 保存ダイアログでファイル名入力・保存
7. 上書き確認が出たら「はい」を押下
8. ダウンロード完了待ち（partial消失 / 安定化）
9. 保存ファイルの更新時刻を確認
10. 起動した IEモード Edge のみ終了

状態遷移 / 待ち条件（重要）
-----------------------
- ログイン後:
  - `document.readyState == complete`
  - `LOC_DOWNLOAD_LINK` が存在するまで待機
- confirm ダイアログ:
  - Win32で `TITLE_CONFIRM` を検出
  - ボタン一覧から `OK` を選択
  - ダイアログの消失を待つ
- ダウンロードバー:
  - UIAで `IENotificationBar` (ToolBar) の可視化待ち
  - キーボード操作で「名前を付けて保存」を起動
  - 保存ダイアログの出現で成功判定
- 保存ダイアログ:
  - `TITLE_SAVE_DIALOG` の可視化待ち
  - `FileNameControlHost` → Edit に保存パスを設定
  - 保存ボタン `auto_id=1` を押下
  - 上書き確認が出たら「はい(&Y) / はい(Y)」を厳密に選択
  - 保存ダイアログの消失を待つ
- ダウンロード完了:
  - `sample.csv.partial` の有無を監視
  - `sample.csv` の更新時刻が `download_start` 以降
  - ファイルサイズが `WAIT_STABLE_SEC` 以上変化しないこと

定数/識別子（移植ポイント）
------------------------
- ダイアログ:
  - `WIN32_CLASS_DIALOG = "#32770"`
  - `TITLE_CONFIRM = "Web ページからのメッセージ"`
  - `TITLE_SAVE_DIALOG = "名前を付けて保存"`
  - `TITLE_OVERWRITE_DIALOG = "名前を付けて保存の確認"`
- IEモードウィンドウ:
  - `IE_WINDOW_TITLE_RE = r"^ダウンロード.*"`
- UIA:
  - `UIA_NOTIFICATION_BAR_ID = "IENotificationBar"`
  - `SAVE_FILENAME_CONTROL_ID = "FileNameControlHost"`
  - `SAVE_BUTTON_AUTO_ID = "1"`
- Selenium locator:
  - `LOC_USER_ID`, `LOC_PASSWORD`, `LOC_DOWNLOAD_LINK`

ログ出力
--------
- `log/run_YYYYMMDDHHMMSS.log` に出力
- ログ行にタイムスタンプとレベルが入る
- 実行時に出力先パスをログに出す

移植の注意点
-----------
- ログイン後の入力確認は `USER_ID` の一致のみ厳密確認
  - パスワードは値取得不可の場合があるため WARN のみ
- ダウンロードバーは UIA で直接クリックせずキーボード操作
  - UIAで SplitButton を触ると不安定になりやすい環境があるため
- 上書き確認は Win32 でボタン一覧を取得し、完全一致で選択

想定される調整ポイント
--------------------
- タイトル/ボタン名のローカライズ差
- IEモードウィンドウのタイトルプレフィックス
- 保存先パス・ファイル名
- ダウンロードバーのキーボード操作順序
