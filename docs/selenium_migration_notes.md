# Selenium + IEDriver 版への移行ノート

## 目的

現在の `automation/ie_mode_test.py` は COM (IWebBrowser2) でページ操作を行っているが、
会社環境では Selenium + IEDriver でページ操作が動作している。
Selenium 版のスクリプトを作成し、pywinauto 部分と組み合わせて動作させる。

---

## 作業ブランチ

`feature/selenium-iedriver` ブランチで作業すること。
master は COM 版（動作確認済み）を保持する。

---

## アーキテクチャの違い

```
【COM版 (現在の master)】
  comtypes.CreateObject("InternetExplorer.Application")
    → ie.Navigate(), ie.Document.getElementsByClassName() 等
    → プロセス: iexplore.exe

【Selenium版 (これから作る)】
  webdriver.Ie(options=options)
    → driver.get(), driver.find_element() 等
    → プロセス: IEDriverServer.exe + msedge.exe
```

---

## Selenium + IEDriver の基本設定

```python
from selenium import webdriver
from selenium.webdriver.ie.service import Service
from selenium.webdriver.ie.options import Options

options = Options()
options.attach_to_edge_chrome = True
options.edge_executable_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
# ↑ Edge の実行パスは環境に合わせる。(x86) の場合と Program Files 直下の場合がある

# 必要に応じて以下を追加
# options.ignore_protected_mode_settings = True
# options.initial_browser_url = "http://localhost:5000"

service = Service(executable_path=r"path\to\IEDriverServer.exe")
driver = webdriver.Ie(service=service, options=options)
```

---

## Step 別の移行ポイント

### Step 1〜7: ページ操作 (Selenium に置き換え)

| COM版 | Selenium版 |
|-------|-----------|
| `ie = comtypes.client.CreateObject(...)` | `driver = webdriver.Ie(...)` |
| `ie.Navigate(url)` | `driver.get(url)` |
| `ie.Document.getElementsByClassName("txtUserID")` | `driver.find_element(By.CLASS_NAME, "txtUserID")` |
| `element.value = "xxx"` | `element.send_keys("xxx")` |
| `button.click()` | `button.click()` |
| `ie.ReadyState == 4` で待機 | `WebDriverWait` + `expected_conditions` |

### Step 8: confirm ダイアログ — 重要な違いあり

**Selenium の場合、まず Alert API を試す:**
```python
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver.find_element(By.PARTIAL_LINK_TEXT, "ダウンロード").click()
WebDriverWait(driver, 10).until(EC.alert_is_present())
driver.switch_to.alert.accept()
```

**Alert API が動かない場合のみ、pywinauto にフォールバック:**
```python
from pywinauto import Desktop
desktop = Desktop(backend="win32")
dialog = desktop.window(class_name="#32770", title_re=".*Web.*")
dialog.wait("visible", timeout=15)
dialog.set_focus()
dialog["OK"].click()
```

**COM版との重要な違い:**
- COM の `link.click()` は `confirm()` が閉じるまでブロックする → スレッド化が必要だった
- Selenium の `click()` はブロックしない可能性が高い → スレッド化が不要かもしれない
- ただし環境依存なので、動作を見て判断すること

### Step 9: ダウンロード通知バー — そのまま使える

pywinauto (UIA) のコードはそのまま流用可能。変更不要。

**唯一の注意点**: IEウィンドウの検出タイトル
- COM版: `title_re=".*Internet Explorer.*"` (IEFrame ウィンドウ)
- Selenium+IEDriver版: Edge IEモードのウィンドウタイトルが異なる可能性あり
- `title_re=".*Internet Explorer.*"` で検出できるか要確認
- 検出できない場合は `title_re=".*ダウンロード.*"` 等、ページタイトルで探す

### Step 10: 保存ダイアログ — そのまま使える

pywinauto (UIA) のコードはそのまま流用可能。変更不要。

---

## 自宅環境での制約

IEDriver 4.14.0 + Edge 144 + Windows 11 24H2 ではセッション確立が失敗する。
詳細は `docs/iedriver_investigation.txt` 参照。

```
停止箇所: BrowserFactory.cpp(507)
  "Using Active Accessibility to find IWebBrowser2 interface"
  → 無限ループ → 120秒でタイムアウト
```

**この問題は会社環境では発生しない** (会社では IEDriver で動作確認済み)。
自宅で Selenium 版を開発する場合は、ページ操作部分の動作確認ができないことを理解しておく。

対処案:
- ページ操作部分は会社環境に合わせてコードを書き、pywinauto 部分のみ自宅で検証
- または、Selenium のページ操作をモック/スキップして pywinauto 部分だけテストする仕組みを入れる

---

## PID管理の違い

| | COM版 | Selenium版 |
|---|-------|-----------|
| 追跡対象プロセス | `iexplore.exe` | `msedge.exe` (または不要) |
| 終了方法 | `ie.Quit()` + `taskkill` | `driver.quit()` で済む可能性大 |

Selenium の `driver.quit()` が正常に動けば、PID追跡は不要かもしれない。

---

## ファイル構成案

```
automation/
├── ie_mode_test.py          # COM版 (master で動作確認済み)
└── selenium_ie_test.py      # Selenium版 (新規作成)
```

または `ie_mode_test.py` を Selenium 版に書き換えてもよい。
master に COM 版が残っているので、ブランチ上では自由に変更して問題ない。

---

## 成功基準

COM版と同じく、以下がすべて動作すること:

1. Selenium + IEDriver で Edge IEモードが起動する
2. ログインページでフォーム入力 → ログインボタンクリック
3. ダウンロードリンクのクリック
4. confirm ダイアログで OK
5. ダウンロード通知バーで「名前を付けて保存」選択
6. 保存ダイアログでファイルパス指定 → 保存
7. 指定パスに CSV ファイルが存在する

---

## 調査スクリプトの記録（削除済み）

複数の IE モードウィンドウやダイアログ判定のため、以下の PowerShell スクリプトで調査を実施した。
スクリプト本体はリポジトリから削除し、調査内容のみをここに記録する。

- `Save-ProcessList.ps1`
  - 目的: 実行中プロセスの一覧を出力
  - 出力例: `log/process_list.txt`
- `Save-WindowList.ps1`
  - 目的: 画面上のウィンドウタイトル一覧を出力（IEモード検出用）
  - 出力例: `log/window_list.txt`
- `Check-IEModeProcesses.ps1`
  - 目的: `msedge.exe` の `CommandLine` を取得し、`--ie-mode-force` など IE モード起動引数を確認
  - 出力例: `log/iemode_processes.txt`
- `Save-DialogClassInfo.ps1`
  - 目的: 「名前を付けて保存」/「上書き確認」ダイアログの Class 名と Title を列挙
  - 出力例: `log/dialog_class_info.txt`

調査結果（要点）:
- 保存ダイアログ: `ClassName=#32770`, `Title=名前を付けて保存`
- 上書き確認ダイアログ: `ClassName=#32770`, `Title=名前を付けて保存の確認`
