# Edge IEモード ダウンロードテスト自動化プロジェクト

## 1. プロジェクトの目的

本プロジェクトは、**Edge の IEモード上でファイルダウンロード操作を自動化する**ための検証環境である。

業務システム等でIEモードを使用している環境において、Selenium + IEDriver によるブラウザ自動操作が実現可能かを検証し、ダウンロードバーの「名前を付けて保存」操作までを含めた一連のフローを自動化することを目標とする。

## 2. ゴール

以下の一連の操作をプログラムで自動実行し、指定パスにCSVファイルが保存されること。

1. IEモードのブラウザを起動し、テスト用Webページを開く
2. ログインページでユーザーID・パスワードを入力してログインする
3. ダウンロードページのリンクをクリックする
4. JavaScript `confirm` ダイアログで「OK」を選択する
5. IEモードのダウンロード通知バーから「名前を付けて保存」を実行する
6. 保存ダイアログでファイルパスを指定して保存する

## 3. 構成要素

### 3-1. テスト用Webページ（Flask サーバー）

テスト対象となるWebページを Flaskで配信する。認証ロジックは持たず、画面遷移とダウンロード動作のみを提供するテスト専用環境。

**ルート一覧:**

| ルート | メソッド | 動作 |
|--------|----------|------|
| `/` `/login` | GET | ログインページを表示 |
| `/login` | POST | ダウンロードページへリダイレクト（認証チェックなし） |
| `/download` | GET | ダウンロードページを表示 |
| `/download/csv` | GET | CSVファイルをダウンロード応答 |

### 3-2. ログインページ (`templates/login.html`)

```
+------------------------------------------+
|  ログイン                                 |
|                                          |
|  ユーザーID                               |
|  [_________________________] ← class="txtUserID"
|                                          |
|  パスワード                               |
|  [*************************] ← class="txtPassWord" (文字マスク)
|                                          |
|  [ログイン]                               |
+------------------------------------------+
```

**HTML要素仕様:**

| 要素 | タグ | class / type | 備考 |
|------|------|-------------|------|
| ユーザーID | `<input>` | `class="txtUserID"` `type="text"` | |
| パスワード | `<input>` | `class="txtPassWord"` `type="password"` | 文字マスク表示 |
| ログインボタン | `<button>` | `type="submit"` | form POST → `/login` |

- フォームは `method="POST"` `action="/login"` で送信
- サーバー側は認証チェックなし、POST受信で即座に `/download` へリダイレクト

### 3-3. ダウンロードページ (`templates/download.html`)

```
+------------------------------------------+
|  ダウンロード                              |
|                                          |
|  CSVファイルをダウンロード  ← <a>リンク     |
+------------------------------------------+
```

**HTML要素仕様:**

```html
<a href="/download/csv" onclick="return confirm('ダウンロードしますか？');">
  CSVファイルをダウンロード
</a>
```

**動作フロー:**

```
リンクをクリック
    │
    ▼
onclick → window.confirm('ダウンロードしますか？')
    │
    ├─ OK (true)  → href="/download/csv" へ遷移 → CSVダウンロード開始
    │                  │
    │                  ▼
    │              IEモードのダウンロード通知バーが表示
    │              [ファイルを開く] [保存 ▼] [キャンセル]
    │                       │
    │                       ▼ (▼をクリック)
    │              [保存]
    │              [名前を付けて保存(A)]  ← これを選択
    │              [保存して開く(O)]
    │                       │
    │                       ▼
    │              「名前を付けて保存」ダイアログ
    │              ファイルパスを指定 → [保存] → 完了
    │
    └─ キャンセル (false) → 何も起こらない
```

### 3-4. CSVファイル (`static/sample.csv`)

ダウンロードされるテスト用データ。
`Content-Disposition: attachment; filename="sample.csv"` ヘッダ付きで配信。

```csv
ID,名前,メールアドレス,部署
1,山田太郎,yamada@example.com,営業部
2,鈴木花子,suzuki@example.com,開発部
3,田中一郎,tanaka@example.com,総務部
```

## 4. 自動化プログラムの想定動作

### 4-1. 当初の構成（Selenium + IEDriver）

```
Python スクリプト
    │
    ├── Selenium WebDriver API
    │     └── IEDriverServer.exe (32-bit)
    │           └── Edge IEモード (msedge.exe --ie-mode-test)
    │
    └── pywinauto
          └── ダウンロード通知バー / 保存ダイアログ操作
```

**操作フロー:**

| Step | 操作 | 使用ツール |
|------|------|-----------|
| 1 | Edge IEモードでブラウザを起動 | Selenium + IEDriver |
| 2 | `http://localhost:5000/login` へ遷移 | Selenium `driver.get()` |
| 3 | `class="txtUserID"` にユーザーIDを入力 | Selenium `find_element()` |
| 4 | `class="txtPassWord"` にパスワードを入力 | Selenium `find_element()` |
| 5 | ログインボタンをクリック | Selenium `click()` |
| 6 | ダウンロードページへの遷移を待機 | Selenium `WebDriverWait` |
| 7 | ダウンロードリンクをクリック | Selenium `click()` |
| 8 | confirm ダイアログで OK を押す | Selenium `switch_to.alert.accept()` |
| 9 | ダウンロード通知バーで「名前を付けて保存」を選択 | pywinauto |
| 10 | 保存ダイアログでファイルパスを指定して保存 | pywinauto |

### 4-2. 判明した問題と方針転換

IEDriver 4.14.0 + Edge 144 + Windows 11 24H2 の組み合わせで、
IEDriverのセッション確立（Step 1）が完了しないことが判明した。

詳細は `docs/iedriver_investigation.txt` を参照。

**現在の構成（COM直接操作 + pywinauto）:**

```
Python スクリプト
    │
    ├── comtypes (IWebBrowser2 COM)
    │     └── InternetExplorer.Application
    │           └── IE (IEFrame ウィンドウ)
    │
    └── pywinauto
          └── confirm ダイアログ / ダウンロード通知バー / 保存ダイアログ操作
```

| Step | 操作 | 使用ツール |
|------|------|-----------|
| 1 | IE (COM) を起動 | comtypes `CreateObject()` |
| 2 | `http://localhost:5000/login` へ遷移 | COM `ie.Navigate()` |
| 3 | `class="txtUserID"` にユーザーIDを入力 | COM DOM API `getElementsByClassName()` |
| 4 | `class="txtPassWord"` にパスワードを入力 | COM DOM API |
| 5 | ログインボタンをクリック | COM DOM API `click()` |
| 6 | ダウンロードページへの遷移を待機 | COM `ReadyState` ポーリング |
| 7 | ダウンロードリンクをクリック | COM DOM API |
| 8 | confirm ダイアログで OK を押す | pywinauto |
| 9 | ダウンロード通知バーで「名前を付けて保存」を選択 | pywinauto |
| 10 | 保存ダイアログでファイルパスを指定して保存 | pywinauto |

## 5. 環境構築手順

### 5-1. 前提条件

- Windows 11 (IE COMインフラが存在すること)
- Python 3.x
- Edge がインストールされていること

### 5-2. セットアップ

```cmd
cd D:\Git\iemode_dl_test
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### 5-3. 実行

**ターミナル1: Flaskサーバー起動**
```cmd
python app.py
```

**ターミナル2: 自動化スクリプト実行**
```cmd
python automation\ie_mode_test.py
```

## 6. 達成基準

以下がすべて満たされた場合、テスト成功とする。

- [ ] Flaskサーバーが起動し、`http://localhost:5000` でログインページが表示される
- [ ] 自動化スクリプトがIEモードブラウザを起動できる
- [ ] ログインページでユーザーID・パスワードが自動入力され、ログインボタンが押される
- [ ] ダウンロードページへ遷移する
- [ ] ダウンロードリンクがクリックされ、confirmダイアログが表示される
- [ ] confirmダイアログで「OK」が自動選択される
- [ ] ダウンロード通知バーが表示され、「名前を付けて保存」が自動選択される
- [ ] 保存ダイアログでファイルパスが指定され、CSVファイルが保存される
- [ ] 指定パス (`D:\Git\iemode_dl_test\download`) にCSVファイルが存在する

## 7. プロジェクト構成

```
D:\Git\iemode_dl_test\
├── app.py                          # Flaskサーバー
├── requirements.txt                # 依存パッケージ (flask, selenium, pywinauto, comtypes)
├── templates/
│   ├── login.html                  # ログインページ
│   └── download.html               # ダウンロードページ
├── static/
│   └── sample.csv                  # ダウンロード用CSVファイル
├── bin/
│   └── IEDriverServer.exe          # IEDriver (※現在は未使用)
├── automation/
│   ├── ie_mode_test.py             # メイン自動化スクリプト (COM方式)
│   ├── diag.py                     # 環境診断スクリプト
│   ├── fix_registry.py             # レジストリ修正スクリプト
│   ├── check_ie_infra.py           # IE COMインフラ確認スクリプト
│   ├── inspect_iemode.py           # IEモードウィンドウ構造調査スクリプト
│   └── minimal_test.py             # IEDriver最小構成テスト
├── docs/
│   ├── project_overview.md         # 本ドキュメント
│   └── iedriver_investigation.txt  # IEDriver調査報告書
└── .vscode/
    └── launch.json                 # デバッグ設定
```
