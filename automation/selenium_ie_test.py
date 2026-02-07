"""
Selenium + IEDriver + pywinauto によるIEモードダウンロードテスト自動化スクリプト

前提条件:
- Flaskサーバー (app.py) が起動していること (http://localhost:5000)
- selenium / pywinauto がインストール済みであること
- IEDriverServer.exe のパスが有効であること
- Edge の実行パスが環境に合っていること
"""

import os
import subprocess
import threading
import time
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.ie.options import Options
from selenium.webdriver.ie.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pywinauto import Desktop

# ===== 設定 =====
BASE_URL = "http://localhost:5000"
SAVE_PATH = r"D:\Git\iemode_dl_test\download"
USER_ID = "testuser"
PASSWORD = "testpass"

IEDRIVER_PATH = r"G:\git\iemode_dl_test\bin\IEDriverServer.exe"
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
STARTUP_TIMEOUT_SEC = 60
IEDRIVER_LOG_PATH = r"G:\git\iemode_dl_test\log\iedriver.log"
IEDRIVER_LOG_LEVEL = "TRACE"
EDGE_USER_DATA_DIR = r"G:\git\iemode_dl_test\edge_profile"
EDGE_PROFILE_DIR = "Profile 1"


def _kill_iedriver_server():
    """IEDriverServer.exe を強制終了する"""
    subprocess.run(
        ["taskkill", "/F", "/IM", "IEDriverServer.exe"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def _kill_existing_ie_mode_edges():
    """既存のIEモードEdge(msedge.exe)を終了する"""
    cmd = (
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.Name -eq 'msedge.exe' -and $_.CommandLine -match '--ie-mode-force' } | "
        "Select-Object -ExpandProperty ProcessId"
    )
    try:
        result = subprocess.run(
            ["powershell", "-Command", cmd],
            capture_output=True, text=True,
        )
        pids = []
        for line in result.stdout.strip().splitlines():
            line = line.strip()
            if line.isdigit():
                pids.append(line)
        for pid in pids:
            subprocess.run(
                ["taskkill", "/F", "/PID", pid],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        if pids:
            print(f"  [CLEANUP] 既存IEモードEdgeを終了: {', '.join(pids)}")
    except Exception as e:
        print(f"  [WARN] 既存IEモードEdgeの終了に失敗: {e}")


def create_driver(timeout_sec=STARTUP_TIMEOUT_SEC):
    """IEDriver + Edge IEモードでWebDriverを起動する (起動タイムアウト付き)"""
    options = Options()
    options.attach_to_edge_chrome = True
    options.edge_executable_path = EDGE_PATH
    options.ignore_protected_mode_settings = True
    options.ignore_zoom_level = True
    # options.ensure_clean_session = True
    options.browser_attach_timeout = 120000
    options.force_create_process_api = True
    options.force_shell_windows_api = False
    options.initial_browser_url = "http://localhost:5000/login"
    os.makedirs(EDGE_USER_DATA_DIR, exist_ok=True)
    options.add_additional_option(
        "ie.browserCommandLineSwitches",
        f'--user-data-dir="{EDGE_USER_DATA_DIR}" --profile-directory="{EDGE_PROFILE_DIR}"',
    )
    options.add_additional_option("ie.ignoreprocessmatch", True)
    options.add_additional_option("nativeEvents", False)
    options.add_additional_option("requireWindowFocus", False)
    options.add_additional_option("enablePersistentHover", False)

    os.makedirs(os.path.dirname(IEDRIVER_LOG_PATH), exist_ok=True)
    service = Service(
        executable_path=IEDRIVER_PATH,
        log_output=IEDRIVER_LOG_PATH,
        service_args=[
            f"/log-level={IEDRIVER_LOG_LEVEL}",
        ],
    )
    result = {"driver": None, "error": None}

    def _worker():
        try:
            result["driver"] = webdriver.Ie(service=service, options=options)
        except Exception as e:
            result["error"] = e

    t = threading.Thread(target=_worker, daemon=True)
    t.start()
    t.join(timeout=timeout_sec)

    if t.is_alive():
        try:
            service.stop()
        except Exception:
            pass
        _kill_iedriver_server()
        raise TimeoutError(
            f"WebDriver起動が{timeout_sec}秒を超えました。"
            " IEDriverServerがセッション確立でハングしている可能性があります。"
        )

    if result["error"] is not None:
        raise result["error"]
    return result["driver"]


def wait_for_ready(driver, timeout=15):
    """document.readyState == complete を待機する"""
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def set_value_with_fallback(element, value):
    """入力欄に値を入れる。テスト用にOSレベルのキーボード入力で確実性を優先する。"""
    from pywinauto import keyboard as kbd
    try:
        element._parent.execute_script("arguments[0].focus();", element)
    except Exception:
        pass
    time.sleep(0.2)
    kbd.send_keys("^a")
    kbd.send_keys(value, with_spaces=True)


def step_login(driver):
    """ログインページでユーザーID・パスワードを入力し、ログインボタンを押す"""
    driver.get(f"{BASE_URL}/login")
    wait_for_ready(driver)

    userid_input = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CLASS_NAME, "txtUserID"))
    )
    password_input = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CLASS_NAME, "txtPassWord"))
    )

    set_value_with_fallback(userid_input, USER_ID)
    set_value_with_fallback(password_input, PASSWORD)

    try:
        uid_val = userid_input.get_attribute("value")
        if uid_val != USER_ID:
            raise RuntimeError(f"ユーザーIDが一致しません: '{uid_val}'")
        print(f"  [DEBUG] ユーザーID確認OK: '{uid_val}'")
    except Exception as e:
        print(f"  [WARN] ユーザーIDの確認に失敗: {e}")

    try:
        pwd_val = password_input.get_attribute("value")
        if not pwd_val:
            print("  [WARN] パスワードの値を取得できない/空です")
    except Exception as e:
        print(f"  [WARN] パスワードの確認に失敗: {e}")

    # IEモードでは click が失敗しやすいので OSレベルの Enter で送信
    from pywinauto import keyboard as kbd
    kbd.send_keys("{ENTER}")

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "a[href*='/download/csv']"))
    )
    print("[OK] ログイン完了 → ダウンロードページへ遷移")


def _handle_confirm_dialog_pywinauto():
    """confirmダイアログをpywinautoで閉じる"""
    desktop = Desktop(backend="win32")
    dialog = desktop.window(class_name="#32770", title="Web ページからのメッセージ")
    dialog.wait("visible", timeout=15)
    try:
        title = dialog.window_text()
    except Exception:
        title = "(unknown)"
    try:
        texts = [w.window_text() for w in dialog.descendants()
                 if w.window_text()]
    except Exception:
        texts = []
    print(f"  [DEBUG] confirmダイアログ検出: title={title}")
    if texts:
        print(f"  [DEBUG] confirmダイアログ内テキスト: {texts}")
    dialog.set_focus()
    try:
        buttons = [b for b in dialog.descendants()
                   if b.friendly_class_name() == "Button" and b.window_text()]
        btn_texts = [b.window_text() for b in buttons]
    except Exception:
        buttons = []
        btn_texts = []
    if btn_texts:
        print(f"  [DEBUG] confirmダイアログのボタン一覧: {btn_texts}")

    target = "OK"
    target_btn = None
    for b in buttons:
        if b.window_text() == target:
            target_btn = b
            break
    if target_btn is None:
        raise RuntimeError(f"confirmダイアログのOKボタンが見つかりません: {btn_texts}")
    try:
        target_btn.click()
    except Exception:
        target_btn.click_input()
    print("  [DEBUG] confirmダイアログで「OK」を選択")
    try:
        dialog.wait_not("visible", timeout=10)
    except Exception:
        pass
    print("[OK] confirmダイアログでOKを選択 (pywinauto)")


def step_click_download_and_confirm(driver):
    """
    ダウンロードリンクをクリックし、confirmダイアログでOKを押す

    Selenium Alert API は環境によってハングするため、Win32ダイアログ操作を優先する。
    """
    t = threading.Thread(target=_handle_confirm_dialog_pywinauto, daemon=True)
    t.start()

    link = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "a[href*='/download/csv']"))
    )
    # IEモードでは click が失敗しやすいので Enter でリンクを起動
    try:
        link._parent.execute_script("arguments[0].focus();", link)
    except Exception:
        pass
    time.sleep(0.2)
    from pywinauto import keyboard as kbd
    kbd.send_keys("{ENTER}")
    print("[OK] ダウンロードリンクを実行 (Enter)")
    t.join(timeout=20)


def _find_ie_window(desktop, timeout=15):
    """IEモードのウィンドウを探して返す（先頭が「ダウンロード」）"""
    patterns = [
        "^ダウンロード.*",
    ]
    end = time.time() + timeout
    while time.time() < end:
        for pattern in patterns:
            try:
                wins = desktop.windows(title_re=pattern)
            except Exception:
                wins = []
            if len(wins) == 1:
                win = wins[0]
                try:
                    win.wait("visible", timeout=2)
                except Exception:
                    pass
                return win
            if len(wins) > 1:
                # アクティブ/フォーカス中を優先して選ぶ
                for w in wins:
                    try:
                        if w.is_active() or w.has_focus():
                            return w
                    except Exception:
                        pass
                # それでも決められなければ先頭を返す
                return wins[0]
        time.sleep(0.5)
    raise RuntimeError("IEモードのウィンドウが見つかりません")


def step_handle_download_bar():
    """
    IEのダウンロード通知バーで「名前を付けて保存」を実行する
    """
    desktop = Desktop(backend="uia")
    time.sleep(3)

    from pywinauto import keyboard as kbd

    # UIAのCOMError対策として通知バー取得をリトライ
    for retry in range(3):
        try:
            ie_window = _find_ie_window(desktop)
            try:
                ie_window.set_focus()
            except Exception:
                pass

            # UIAWrapper の場合は handle から WindowSpecification を作り直す
            try:
                ie_spec = desktop.window(handle=ie_window.handle)
            except Exception:
                ie_spec = ie_window

            notification_bar = ie_spec.child_window(auto_id="IENotificationBar",
                                                    control_type="ToolBar")
            notification_bar.wait("visible", timeout=10)

            # キーボード操作のみで「名前を付けて保存」を選択する
            save_dialog = desktop.window(title="名前を付けて保存")
            for attempt in range(3):
                kbd.send_keys("%n")
                time.sleep(0.3)
                kbd.send_keys("{TAB}")
                time.sleep(0.3)
                kbd.send_keys("{DOWN}")
                time.sleep(0.3)
                kbd.send_keys("{DOWN}")
                time.sleep(0.3)
                kbd.send_keys("{ENTER}")
                if save_dialog.exists(timeout=5):
                    print("[OK] ダウンロードバーで「名前を付けて保存」を選択 (キーボード)")
                    return
                # 反応しなかった場合はメニューを閉じて再試行
                kbd.send_keys("{ESC}")
                time.sleep(0.5)
                print(f"  [WARN] 保存ダイアログ未表示。再試行 {attempt + 1}/3")
            raise RuntimeError("ダウンロードバーで「名前を付けて保存」を起動できませんでした")
        except Exception as e:
            print(f"  [WARN] ダウンロードバー取得に失敗。再試行 {retry + 1}/3: {e}")
            time.sleep(1.0)

    raise RuntimeError("ダウンロードバーの取得に繰り返し失敗しました")


def step_handle_save_dialog():
    """
    「名前を付けて保存」ダイアログでファイルパスを指定して保存する
    """
    from pywinauto import keyboard as kbd

    desktop = Desktop(backend="uia")

    save_dialog = desktop.window(title="名前を付けて保存")
    save_dialog.wait("visible", timeout=20)
    save_dialog.set_focus()
    time.sleep(1)

    save_file_path = os.path.join(SAVE_PATH, "sample.csv")
    before_mtime = os.path.getmtime(save_file_path) if os.path.exists(save_file_path) else None
    os.makedirs(SAVE_PATH, exist_ok=True)
    print(f"  [DEBUG] 保存先: {save_file_path}")

    try:
        fn_host = save_dialog.child_window(auto_id="FileNameControlHost")
        fn_edit = fn_host.child_window(control_type="Edit")
        fn_edit.set_edit_text(save_file_path)
        print("  [DEBUG] ファイル名設定: FileNameControlHost経由")
    except Exception as e1:
        raise RuntimeError(f"ファイル名フィールドへの入力に失敗しました: {e1}")

    time.sleep(0.5)

    try:
        save_btn = save_dialog.child_window(auto_id="1", control_type="Button")
        download_start = time.time()
        save_btn.click_input()
        print("  [DEBUG] 保存ボタン押下: auto_id=1経由")
    except Exception as e1:
        raise RuntimeError(f"保存ボタンの押下に失敗しました: {e1}")

    # 上書き確認ダイアログの検出と確実なクリック
    try:
        win32 = Desktop(backend="win32")
        confirm_overwrite = win32.window(class_name="#32770", title="名前を付けて保存の確認")
        if confirm_overwrite.exists(timeout=5):
            try:
                title = confirm_overwrite.window_text()
            except Exception:
                title = "(unknown)"
            try:
                texts = [w.window_text() for w in confirm_overwrite.descendants()
                         if w.window_text()]
            except Exception:
                texts = []
            print(f"  [DEBUG] 上書き確認ダイアログ検出: title={title}")
            if texts:
                print(f"  [DEBUG] 上書き確認ダイアログ内テキスト: {texts}")

            # ボタン候補を収集して完全一致で押す
            try:
                buttons = [b for b in confirm_overwrite.descendants()
                           if b.friendly_class_name() == "Button" and b.window_text()]
                btn_texts = [b.window_text() for b in buttons]
            except Exception:
                buttons = []
                btn_texts = []
            if btn_texts:
                print(f"  [DEBUG] 上書き確認ダイアログのボタン一覧: {btn_texts}")

            def _norm_button_text(s):
                # すべての空白系と不可視制御を除去
                out = []
                for ch in s:
                    if ch.isspace():
                        continue
                    # Unicode category: Cc/Cf を除外
                    if ord(ch) < 32 or ord(ch) == 127 or ch == "\u200e" or ch == "\u200f":
                        continue
                    out.append(ch)
                return "".join(out)

            targets = ["はい(&Y)", "はい(Y)"]
            norm_targets = {_norm_button_text(t): t for t in targets}
            matched = []
            for t in btn_texts:
                norm_t = _norm_button_text(t)
                if norm_t in norm_targets:
                    matched.append(t)

            if len(matched) == 1:
                target = matched[0]
                target_btn = None
                for b in buttons:
                    if b.window_text() == target:
                        target_btn = b
                        break
                if target_btn is None:
                    raise RuntimeError(f"上書き確認ダイアログのボタンが取得できません: {target}")
                try:
                    target_btn.click()
                except Exception:
                    target_btn.click_input()
                print(f"  [DEBUG] 上書き確認ダイアログで「{target}」を選択")
                # 上書き確認の押下後をダウンロード開始時刻とみなす
                download_start = time.time()
            elif len(matched) == 0:
                raise RuntimeError("上書き確認ダイアログのボタンが見つかりません: はい(&Y) / はい(Y)")
            else:
                raise RuntimeError(f"上書き確認ダイアログのボタンが複数一致: {matched}")
            try:
                confirm_overwrite.wait_not("visible", timeout=10)
            except Exception:
                pass
    except Exception as e:
        print(f"  [WARN] 上書き確認処理で例外: {e}")

    # ダイアログが閉じるのを待機
    try:
        save_dialog.wait_not("visible", timeout=10)
    except Exception:
        pass
    time.sleep(1)

    if download_start is None:
        download_start = time.time()
    print("[OK] 「名前を付けて保存」ダイアログで保存を実行")
    return save_file_path, before_mtime, download_start


def wait_for_download_complete(save_file_path, start_time, timeout=60, stable_sec=3):
    """partialファイル消滅と本体の更新を待つ（開始時刻以降のもののみ対象）"""
    partial_path = f"{save_file_path}.partial"
    end = time.time() + timeout
    try:
        start_time_local = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(start_time))
    except Exception:
        start_time_local = str(start_time)
    print(
        f"  [DEBUG] ダウンロード監視: file={save_file_path}, partial={partial_path}, "
        f"start_time={start_time_local}"
    )
    last_size = None
    last_mtime = None
    stable_since = None
    while time.time() < end:
        partial_exists = os.path.exists(partial_path)
        file_exists = os.path.exists(save_file_path)
        partial_is_new = False
        partial_mtime = None
        if partial_exists:
            try:
                partial_mtime = os.path.getmtime(partial_path)
                partial_is_new = partial_mtime >= start_time
            except Exception:
                partial_is_new = True
        file_is_new = False
        file_mtime = None
        file_size = None
        if file_exists:
            try:
                file_mtime = os.path.getmtime(save_file_path)
                file_size = os.path.getsize(save_file_path)
                file_is_new = file_mtime >= start_time
            except Exception:
                file_is_new = True

        if file_exists and file_is_new and not partial_is_new:
            return True

        # .partial が残っていても、本体が更新されて安定していれば完了とみなす
        if file_exists:
            if file_size == last_size and file_mtime == last_mtime:
                if stable_since is None:
                    stable_since = time.time()
                elif time.time() - stable_since >= stable_sec:
                    return True
            else:
                stable_since = None
            last_size = file_size
            last_mtime = file_mtime

        time.sleep(0.5)
    raise TimeoutError(
        f"ダウンロード完了待ちがタイムアウト: {save_file_path} / {partial_path} "
        f"(file_mtime={file_mtime}, file_size={file_size}, "
        f"partial_mtime={partial_mtime})"
    )


def main():
    driver = None
    try:
        print("IEDriver + Edge IEモードを起動中...")
        _kill_existing_ie_mode_edges()
        driver = create_driver()
        time.sleep(2)
        print("[OK] WebDriver起動完了")

        step_login(driver)
        step_click_download_and_confirm(driver)
        step_handle_download_bar()
        save_file_path, before_mtime, download_start = step_handle_save_dialog()

        # ダウンロード完了待ち (.partial が消えるまで)
        wait_for_download_complete(save_file_path, download_start, timeout=90)

        if os.path.exists(save_file_path):
            file_size = os.path.getsize(save_file_path)
            after_mtime = os.path.getmtime(save_file_path)
            if before_mtime is None:
                print(f"[OK] ファイル保存確認: {save_file_path} ({file_size} bytes)")
            else:
                if after_mtime > before_mtime:
                    print(f"[OK] ファイル更新確認: {save_file_path} ({file_size} bytes)")
                else:
                    print(f"[WARN] ファイルが更新されていない可能性: {save_file_path} ({file_size} bytes)")
        else:
            print(f"[WARN] ファイルが見つかりません: {save_file_path}")

        print("\n===== テスト完了 =====")
        time.sleep(2)
    except Exception as e:
        print(f"\n[ERROR] テスト失敗: {e}")
        raise
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        print("ブラウザを終了しました")


if __name__ == "__main__":
    main()
