# IEモード ダウンロードバー自動操作 実装ガイド

## 本資料の目的

IEモード（またはIE）のダウンロード通知バーを pywinauto で自動操作し、
「名前を付けて保存」までを完了させる手順を、再現可能な形でまとめたものである。

会社環境では Selenium + IEDriver による Webページ操作（ログイン・画面遷移・リンククリック）は
既に動作している前提であり、本資料の核心は **Step 8〜10（confirmダイアログ・ダウンロードバー・
保存ダイアログ）の pywinauto による操作方法** にある。

---

## 全体の処理フロー

```
Step 1-7:  Selenium + IEDriver (会社環境で動作済み)
           ログイン → 画面遷移 → リンククリック

Step 8:    pywinauto (Win32)  ← JavaScript confirm ダイアログ処理
Step 9:    pywinauto (UIA)    ← ダウンロード通知バー操作
Step 10:   pywinauto (UIA)    ← 「名前を付けて保存」ダイアログ操作
```

**重要**: pywinauto には `"win32"` と `"uia"` の2つのバックエンドがあり、
操作対象によって使い分ける必要がある。間違えると要素が見つからない。

| 操作対象 | バックエンド | 理由 |
|----------|-------------|------|
| confirm ダイアログ | `win32` | Win32標準ダイアログ (class `#32770`)。UIAでは検出不可 |
| ダウンロード通知バー | `uia` | IE の通知バーは UIA ツリーでのみ構造が見える |
| 名前を付けて保存ダイアログ | `uia` | Windows共通ファイルダイアログ。UIAで操作可能 |

---

## 検証済み環境

```
OS:        Windows 11 24H2 (Build 10.0.26100.7306)
Edge:      144.0.3719.104
IE:        11.1882.26100.0 (COM インフラとして残存)
Python:    3.13.11 (64-bit)
pywinauto: pip install pywinauto で取得した最新版
comtypes:  pip install comtypes で取得した最新版 (※自宅検証用。会社ではSelenium使用)
```

---

## 必要なライブラリ

```
pip install pywinauto
```

会社環境では selenium は既にあるはず。
自宅検証用に COM 経由で IE を操作する場合は追加で:

```
pip install comtypes
```

---

## Step 8: JavaScript confirm ダイアログの処理

### 背景

ダウンロードリンクに `onclick="return confirm('ダウンロードしますか？');"` が
設定されている場合、クリック後に Windows 標準の confirm ダイアログが表示される。

**Selenium の場合**: `driver.switch_to.alert.accept()` で処理できる可能性がある。
ただし IEモードでは Selenium の Alert API が効かないケースがあるため、
pywinauto での処理方法も把握しておく必要がある。

### confirm ダイアログの正体

- **ウィンドウクラス**: `#32770` (Win32 標準ダイアログ)
- **タイトル**: `Web ページからのメッセージ` (環境によって異なる。`".*Web.*"` で正規表現マッチ)
- **ボタン**: `OK` / `キャンセル`
- **UIA では検出できない**。必ず `Desktop(backend="win32")` を使用すること

### ダイアログ構造 (Spy++/inspect.exe で確認済み)

```
Dialog '#32770' - 'Web ページからのメッセージ'
  ├── Static (メッセージテキスト)
  ├── Button 'OK'
  └── Button 'キャンセル'
```

### 実装コード

```python
from pywinauto import Desktop

def handle_confirm_dialog():
    """confirm ダイアログで OK を押す"""
    desktop = Desktop(backend="win32")  # ← 必ず win32
    dialog = desktop.window(class_name="#32770", title_re=".*Web.*")
    dialog.wait("visible", timeout=15)
    dialog.set_focus()
    dialog["OK"].click()
```

### 重要な落とし穴: COM の click() はブロックする

COM (IWebBrowser2) 経由で `link.click()` を実行した場合、
JavaScript の `confirm()` ダイアログが閉じるまで **click() が戻ってこない**。
つまり、click() の後にダイアログ処理コードを書いても、永遠に到達しない。

**解決策**: ダイアログ処理を **別スレッド** で先に起動してから、リンクをクリックする。

```python
import threading

def _handle_confirm_dialog_thread():
    """別スレッドで confirm ダイアログを待機して OK を押す"""
    desktop = Desktop(backend="win32")
    dialog = desktop.window(class_name="#32770", title_re=".*Web.*")
    dialog.wait("visible", timeout=15)
    dialog.set_focus()
    dialog["OK"].click()

# 使い方:
# 1. 先にスレッドを起動
t = threading.Thread(target=_handle_confirm_dialog_thread, daemon=True)
t.start()

# 2. リンクをクリック (confirm が出てブロック → 別スレッドが OK → ブロック解除)
link.click()

# 3. スレッドの完了を待つ
t.join(timeout=15)
```

**Selenium の場合は click() がブロックしない可能性がある**ため、
スレッド化が不要な場合もある。動作を見て判断すること。

---

## Step 9: ダウンロード通知バーの操作 (核心部分)

### 背景

IEモードでファイルダウンロードが開始されると、ウィンドウ下部に通知バーが表示される。

```
┌──────────────────────────────────────────────────────┐
│  (ページコンテンツ)                                      │
│                                                      │
├──────────────────────────────────────────────────────┤
│ sample.csv を開くか、または保存しますか？                   │
│           [ファイルを開く(O)]  [保存(S) ▼]  [キャンセル]  │
└──────────────────────────────────────────────────────┘
                                    ↑
                               この ▼ を押して
                               「名前を付けて保存(A)」を選ぶ
```

### 通知バーの UI ツリー構造 (UIA で確認済み)

```
Window 'ページタイトル - Internet Explorer'
  └── ToolBar 'IENotificationBar'  ← auto_id="IENotificationBar"
        ├── Text (Static)  '〜を開くか、または保存しますか？'
        ├── Button 'ファイルを開く(O)'
        ├── SplitButton '保存'      ← 親SplitButton
        │     └── SplitButton '6'   ← ★ドロップダウン矢印 (これをクリック)
        ├── Button 'キャンセル'
        └── Button '閉じる'
```

**重要ポイント**:
- 「保存」は `SplitButton` であり、`Button` ではない
- ドロップダウン矢印は「保存」SplitButton の **子要素** で、`title="6"` の `SplitButton`
- `Button` として探すと見つからない。`control_type="SplitButton"` を指定すること

### ドロップダウンメニューの構造

ドロップダウン矢印をクリックすると、以下のメニューが表示される:

```
┌─────────────────────┐
│ 保存(S)              │
│ 名前を付けて保存(A)   │  ← これを選択
│ 保存して開く(O)       │
└─────────────────────┘
```

- メニューは `control_type="Menu"` のトップレベルウィンドウとして出現
- 各項目は `control_type="MenuItem"`

### 実装コード

```python
from pywinauto import Desktop
import time

def handle_download_bar():
    """ダウンロード通知バーで「名前を付けて保存」を選択する"""
    desktop = Desktop(backend="uia")  # ← 必ず uia
    time.sleep(3)  # 通知バー表示の待機 (短すぎると検出失敗)

    # 1. IEウィンドウを検出して最前面化
    ie_window = desktop.window(title_re=".*Internet Explorer.*")
    ie_window.wait("visible", timeout=15)
    ie_window.set_focus()

    # 2. 通知バーを検出
    notification_bar = ie_window.child_window(
        auto_id="IENotificationBar",
        control_type="ToolBar"
    )
    notification_bar.wait("visible", timeout=10)

    # 3. 「保存」SplitButton を検出
    save_button = notification_bar.child_window(
        title="保存",
        control_type="SplitButton"
    )
    save_button.wait("visible", timeout=10)

    # 4. ★ドロップダウン矢印 (子SplitButton, title="6") をクリック
    #    ※ save_button 自体をクリックすると「保存」が直接実行されてしまう
    #    ※ ドロップダウンを開くには、子要素の title="6" をクリックする
    dropdown = save_button.child_window(
        title="6",
        control_type="SplitButton"
    )
    dropdown.click_input()
    time.sleep(1)  # メニュー表示待ち

    # 5. 「名前を付けて保存(A)」メニュー項目をクリック
    save_as_item = desktop.window(control_type="Menu").child_window(
        title="名前を付けて保存(A)",
        control_type="MenuItem"
    )
    save_as_item.click_input()
```

### トラブルシューティング

| 症状 | 原因 | 対処 |
|------|------|------|
| `IENotificationBar` が見つからない | 通知バーの表示が遅い | `time.sleep()` を増やす (3〜5秒) |
| `SplitButton '保存'` が見つからない | バックエンドが違う | `backend="uia"` を確認 |
| `SplitButton '6'` が見つからない | IE のバージョン差 | `save_button.print_control_identifiers()` で子要素を確認 |
| メニューが閉じてしまう | `click()` と `click_input()` の違い | `click_input()` を使用 (実マウス操作) |
| 「名前を付けて保存(A)」が見つからない | メニュータイトルの差異 | `title_re=".*名前.*保存.*"` で正規表現マッチ |

### 要素の調べ方 (デバッグ用)

環境によって title や auto_id が異なる場合がある。以下のコードで実際の構造をダンプできる:

```python
# 通知バーの全子要素を表示
notification_bar.print_control_identifiers(depth=4)

# 保存ボタンの子要素を表示
save_button.print_control_identifiers(depth=2)
```

出力例:
```
SplitButton - '保存'    (L889, T707, R941, B731)
   | SplitButton - '6'    (L923, T707, R941, B731)
```

---

## Step 10: 「名前を付けて保存」ダイアログの操作

### ダイアログの構造

Windows 標準のファイル保存ダイアログ (`#32770`) が開く。

```
┌─ 名前を付けて保存 ──────────────────────────┐
│                                            │
│  (フォルダツリー / ファイル一覧)               │
│                                            │
│  ファイル名(N): [sample.csv           ]  ▼  │
│  ファイルの種類(T): [すべてのファイル]    ▼  │
│                                            │
│                    [保存(S)]  [キャンセル]   │
└────────────────────────────────────────────┘
```

### UIA 要素の特定方法

| 要素 | 検出方法 (優先順) |
|------|------------------|
| ファイル名欄 | `auto_id="FileNameControlHost"` 内の `Edit` |
| 保存ボタン | `auto_id="1"` (Win32 標準の IDOK) |

### 実装コード

```python
from pywinauto import Desktop, keyboard as kbd
import os
import time

SAVE_PATH = r"C:\path\to\save\directory"  # 環境に合わせて変更

def handle_save_dialog():
    """「名前を付けて保存」ダイアログでファイルを保存する"""
    desktop = Desktop(backend="uia")  # ← 必ず uia

    save_dialog = desktop.window(title="名前を付けて保存")
    save_dialog.wait("visible", timeout=15)
    save_dialog.set_focus()
    time.sleep(1)

    # 保存先フルパス
    save_file_path = os.path.join(SAVE_PATH, "sample.csv")
    os.makedirs(SAVE_PATH, exist_ok=True)

    # --- ファイル名入力 ---
    file_name_set = False

    # 方法1: auto_id="FileNameControlHost" 経由 (推奨)
    try:
        fn_host = save_dialog.child_window(auto_id="FileNameControlHost")
        fn_edit = fn_host.child_window(control_type="Edit")
        fn_edit.set_edit_text(save_file_path)
        file_name_set = True
    except Exception:
        pass

    # 方法2: title で探す
    if not file_name_set:
        try:
            fn_edit = save_dialog.child_window(
                title_re=".*ファイル名.*", control_type="Edit"
            )
            fn_edit.set_edit_text(save_file_path)
            file_name_set = True
        except Exception:
            pass

    # 方法3: キーボード操作 (Alt+N でファイル名欄にフォーカス)
    if not file_name_set:
        save_dialog.set_focus()
        time.sleep(0.5)
        kbd.send_keys("%n")        # Alt+N → ファイル名欄にフォーカス
        time.sleep(0.5)
        kbd.send_keys("^a")        # Ctrl+A → 全選択
        kbd.send_keys(save_file_path, with_spaces=True)
        file_name_set = True

    time.sleep(0.5)

    # --- 保存ボタン押下 ---
    save_clicked = False

    # 方法1: auto_id="1" (IDOK) (推奨)
    try:
        save_btn = save_dialog.child_window(auto_id="1", control_type="Button")
        save_btn.click_input()
        save_clicked = True
    except Exception:
        pass

    # 方法2: title で探す
    if not save_clicked:
        try:
            save_btn = save_dialog.child_window(
                title_re=".*保存.*", control_type="Button"
            )
            save_btn.click_input()
            save_clicked = True
        except Exception:
            pass

    # 方法3: Alt+S
    if not save_clicked:
        save_dialog.set_focus()
        time.sleep(0.3)
        kbd.send_keys("%s")  # Alt+S
        save_clicked = True

    # 方法4: Enter (最終手段)
    if not save_clicked:
        kbd.send_keys("{ENTER}")

    time.sleep(2)

    # --- 上書き確認ダイアログ対応 ---
    try:
        confirm = desktop.window(title="名前を付けて保存の確認")
        if confirm.exists(timeout=2):
            confirm.child_window(
                title="はい(&Y)", control_type="Button"
            ).click_input()
            time.sleep(1)
    except Exception:
        pass  # 上書き確認が出なければスキップ
```

### 検証で判明した成功パターン

自宅 Windows 11 24H2 環境での実行結果:

```
ファイル名設定: FileNameControlHost経由  ← 方法1で成功
保存ボタン押下: auto_id=1経由          ← 方法1で成功
```

---

## 会社環境への適用方法

### 前提

会社環境では Selenium + IEDriver でページ操作ができているため、
Step 1〜7 は既存の Selenium コードをそのまま使う。

### 組み込みイメージ

```python
from selenium import webdriver
from selenium.webdriver.ie.options import Options
from selenium.webdriver.common.by import By
import threading
import time

# ===== 既存の Selenium コード (Step 1〜7) =====

options = Options()
options.attach_to_edge_chrome = True
options.edge_executable_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
# options.ignore_protected_mode_settings = True  # 必要に応じて

driver = webdriver.Ie(options=options)

# ログイン
driver.get("http://対象システムのURL/login")
driver.find_element(By.CLASS_NAME, "txtUserID").send_keys("ユーザーID")
driver.find_element(By.CLASS_NAME, "txtPassWord").send_keys("パスワード")
driver.find_element(By.TAG_NAME, "button").click()
time.sleep(2)

# ===== ここから pywinauto (Step 8〜10) =====

# --- Step 8: confirm ダイアログ ---
# Selenium の alert API を試す
try:
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    # ダウンロードリンクをクリック
    driver.find_element(By.PARTIAL_LINK_TEXT, "ダウンロード").click()

    # Selenium の Alert API で処理を試みる
    WebDriverWait(driver, 5).until(EC.alert_is_present())
    driver.switch_to.alert.accept()

except Exception:
    # Selenium で処理できない場合は pywinauto にフォールバック
    # (COM の場合はスレッド化が必要だが、Selenium の場合は不要な可能性あり)
    from pywinauto import Desktop
    desktop = Desktop(backend="win32")
    dialog = desktop.window(class_name="#32770", title_re=".*Web.*")
    dialog.wait("visible", timeout=15)
    dialog.set_focus()
    dialog["OK"].click()

# --- Step 9: ダウンロード通知バー ---
handle_download_bar()   # 上記の関数をそのまま使う

# --- Step 10: 名前を付けて保存 ---
handle_save_dialog()    # 上記の関数をそのまま使う

# 完了
driver.quit()
```

### 会社環境で調整が必要になりうる箇所

| 箇所 | 確認方法 | 調整例 |
|------|---------|--------|
| IE ウィンドウタイトル | `title_re=".*Internet Explorer.*"` | 業務システム名が入る場合は正規表現を調整 |
| 通知バーの auto_id | `print_control_identifiers()` で確認 | `IENotificationBar` 以外の場合あり |
| ドロップダウン矢印の title | `save_button.print_control_identifiers()` | `"6"` 以外の場合あり |
| メニュー項目名 | 目視確認 | `"名前を付けて保存(A)"` のカッコが全角/半角 |
| 保存ダイアログのタイトル | 目視確認 | `"名前を付けて保存"` 以外の場合あり |
| SAVE_PATH | 業務要件に合わせる | `r"C:\Users\xxx\Downloads"` 等 |

---

## 調査用コード: UI要素のダンプ

環境が異なると要素の title や auto_id が変わることがある。
以下のコードで、その環境の実際の UI 構造を調べることができる。

### 通知バーの構造を調べる

ダウンロードを開始して通知バーが表示された状態で以下を実行:

```python
from pywinauto import Desktop
import time

desktop = Desktop(backend="uia")
ie_window = desktop.window(title_re=".*Internet Explorer.*")
ie_window.set_focus()

# 通知バー全体の構造をダンプ (depth=5 で十分深く)
bar = ie_window.child_window(auto_id="IENotificationBar", control_type="ToolBar")
bar.print_control_identifiers(depth=5)
```

### 保存ダイアログの構造を調べる

「名前を付けて保存」ダイアログが表示された状態で以下を実行:

```python
from pywinauto import Desktop

desktop = Desktop(backend="uia")
dlg = desktop.window(title="名前を付けて保存")
dlg.print_control_identifiers(depth=3)
```

---

## click() と click_input() の違い

pywinauto には2種類のクリック方法がある。

| メソッド | 動作 | 用途 |
|---------|------|------|
| `click()` | Win32 メッセージ送信 (WM_CLICK) | バックグラウンドでも動作。ただしUI要素によっては反応しない |
| `click_input()` | 実際のマウスカーソル移動 + クリック | 確実に動作する。ただしマウスが動く。ウィンドウが前面にある必要あり |

**本実装では `click_input()` を使用している。**
理由: ダウンロード通知バーやメニュー項目は `click()` では反応しないケースがあったため。

`click_input()` を使う場合は事前に `set_focus()` でウィンドウを最前面にすること。

---

## pywinauto バックエンド早見表

```python
# Win32 バックエンド - 古い Win32 ダイアログ用
desktop = Desktop(backend="win32")
dialog = desktop.window(class_name="#32770", title="...")
dialog["OK"].click()       # ボタン名で直接アクセス可能

# UIA バックエンド - モダンな UI 用
desktop = Desktop(backend="uia")
window = desktop.window(title_re=".*Internet Explorer.*")
child = window.child_window(auto_id="xxx", control_type="Button")
child.click_input()
```

**使い分けの原則**:
- `#32770` クラスの標準ダイアログ (alert, confirm) → `win32`
- IE の通知バー、ファイルダイアログ等 → `uia`
- 迷ったら両方試す。`print_control_identifiers()` で要素が見えるかどうかで判断

---

## 自宅検証用: テスト環境の再現手順

自宅で全フローを再現するための手順。

### 1. フォルダ構成

```
任意のフォルダ\
├── app.py                  # Flask テストサーバー
├── requirements.txt        # pip install -r requirements.txt
├── templates\
│   ├── login.html          # ログインページ
│   └── download.html       # ダウンロードページ
├── static\
│   └── sample.csv          # ダウンロード用データ
├── automation\
│   └── ie_mode_test.py     # 自動化スクリプト (COM版)
└── download\               # ファイル保存先 (自動作成)
```

### 2. 環境構築

```cmd
python -m venv .venv
.venv\Scripts\activate
pip install flask pywinauto comtypes
```

### 3. ファイル作成

#### app.py
```python
from flask import Flask, render_template, redirect, url_for, request, send_file
import os

app = Flask(__name__)

@app.route("/")
@app.route("/login", methods=["GET"])
def login_page():
    return render_template("login.html")

@app.route("/login", methods=["POST"])
def login_submit():
    return redirect(url_for("download_page"))

@app.route("/download")
def download_page():
    return render_template("download.html")

@app.route("/download/csv")
def download_csv():
    csv_path = os.path.join(app.static_folder, "sample.csv")
    return send_file(csv_path, as_attachment=True, download_name="sample.csv")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
```

#### templates/login.html
```html
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>ログイン</title>
</head>
<body>
    <h1>ログイン</h1>
    <form method="POST" action="/login">
        <div>
            <label>ユーザーID</label><br>
            <input type="text" class="txtUserID" name="userid">
        </div>
        <div>
            <label>パスワード</label><br>
            <input type="password" class="txtPassWord" name="password">
        </div>
        <div>
            <button type="submit">ログイン</button>
        </div>
    </form>
</body>
</html>
```

#### templates/download.html
```html
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>ダウンロード</title>
</head>
<body>
    <h1>ダウンロード</h1>
    <p>
        <a href="/download/csv"
           onclick="return confirm('ダウンロードしますか？');">
           CSVファイルをダウンロード
        </a>
    </p>
</body>
</html>
```

#### static/sample.csv
```csv
ID,名前,メールアドレス,部署
1,山田太郎,yamada@example.com,営業部
2,鈴木花子,suzuki@example.com,開発部
3,田中一郎,tanaka@example.com,総務部
```

### 4. 実行

ターミナル1 (サーバー起動):
```cmd
python app.py
```

ターミナル2 (自動化実行):
```cmd
python automation\ie_mode_test.py
```

### 5. 期待される出力

```
  [DEBUG] 起動前のIE PID数: 0
IE (IWebBrowser2 COM) を起動中...
  [DEBUG] 新規IEプロセスPID: {xxxxx}
[OK] IE起動完了
[OK] ログイン完了 → ダウンロードページへ遷移
[OK] ダウンロードリンクをクリック
[OK] confirmダイアログでOKを選択
[OK] ダウンロードバーで「名前を付けて保存」を選択
  [DEBUG] 保存先: D:\...\download\sample.csv
  [DEBUG] ファイル名設定: FileNameControlHost経由
  [DEBUG] 保存ボタン押下: auto_id=1経由
[OK] 「名前を付けて保存」ダイアログで保存を実行
[OK] ファイル保存確認: D:\...\download\sample.csv (171 bytes)

===== テスト完了 =====
```

---

## 補足: 自宅環境での IEDriver 問題について

自宅環境 (Windows 11 24H2 + Edge 144) では Selenium + IEDriver のセッション確立が
できなかった。これは **自宅環境固有の問題** であり、会社環境には影響しない。

原因: Edge 144 の `--ie-mode-test` で起動されたIEモードタブが、
IEDriver が期待する `IEFrame` ウィンドウクラスを持たず、
Active Accessibility 経由の IWebBrowser2 検出が無限ループする。

そのため自宅では COM (`InternetExplorer.Application`) で IE を直接起動する方式で
pywinauto のコードを検証した。pywinauto 部分（Step 8〜10）のコードは
IEDriver 方式でも COM 方式でも **同じものがそのまま使える**。
