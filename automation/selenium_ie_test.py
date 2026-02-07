"""
Selenium + IEDriver + pywinauto によるIEモードダウンロードテスト自動化スクリプト

前提条件:
- Flaskサーバー (app.py) が起動していること (http://localhost:5000)
- selenium / pywinauto がインストール済みであること
- IEDriverServer.exe のパスが有効であること
- Edge の実行パスが環境に合っていること
"""

import logging
import os
import subprocess
import threading
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.ie.options import Options
from selenium.webdriver.ie.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pywinauto import Desktop

# ===== 設定 =====
BASE_URL = "http://localhost:5000"
SAVE_PATH = r"D:\Git\iemode_dl_test\download"
SAVE_FILENAME = "sample.csv"
USER_ID = "testuser"
PASSWORD = "testpass"

IEDRIVER_PATH = r"G:\git\iemode_dl_test\bin\IEDriverServer.exe"
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
STARTUP_TIMEOUT_SEC = 60
IEDRIVER_LOG_PATH = r"G:\git\iemode_dl_test\log\iedriver.log"
IEDRIVER_LOG_LEVEL = "TRACE"

# ===== 画面・要素定義（移植用） =====
WIN32_CLASS_DIALOG = "#32770"
TITLE_CONFIRM = "Web ページからのメッセージ"
TITLE_SAVE_DIALOG = "名前を付けて保存"
TITLE_OVERWRITE_DIALOG = "名前を付けて保存の確認"
IE_WINDOW_TITLE_RE = r"^ダウンロード.*"

LOC_USER_ID = (By.CLASS_NAME, "txtUserID")
LOC_PASSWORD = (By.CLASS_NAME, "txtPassWord")
LOC_DOWNLOAD_LINK = (By.CSS_SELECTOR, "a[href*='/download/csv']")

UIA_NOTIFICATION_BAR_ID = "IENotificationBar"
SAVE_FILENAME_CONTROL_ID = "FileNameControlHost"
SAVE_BUTTON_AUTO_ID = "1"

WAIT_LOGIN_PAGE = 20
WAIT_POST_LOGIN = 10
WAIT_CONFIRM_DIALOG = 15
WAIT_DIALOG_CLOSE = 10
WAIT_SAVE_DIALOG = 20
WAIT_NOTIFICATION_BAR = 10
WAIT_DOWNLOAD_TIMEOUT = 90
WAIT_STABLE_SEC = 3

_tracked_edge_pids = set()
_logger = logging.getLogger("iemode_dl_test")


def init_logging():
    log_dir = os.path.dirname(IEDRIVER_LOG_PATH)
    os.makedirs(log_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    run_log_path = os.path.join(log_dir, f"run_{ts}.log")
    _logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")

    file_handler = logging.FileHandler(run_log_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(fmt)

    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.INFO)
    stream_handler.setFormatter(fmt)

    _logger.addHandler(file_handler)
    _logger.addHandler(stream_handler)
    _logger.propagate = False
    _logger.info(f"[OK] run.log 出力先: {run_log_path}")


def log(message, level=logging.INFO):
    _logger.log(level, message)


def _kill_iedriver_server():
    """IEDriverServer.exe を強制終了する"""
    subprocess.run(
        ["taskkill", "/F", "/IM", "IEDriverServer.exe"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def _kill_existing_ie_mode_edges():
    """既存のIEモードEdge(msedge.exe)を終了する"""
    pids = _get_ie_mode_edge_pids()
    if not pids:
        return
    for pid in pids:
        subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    log(f"  [CLEANUP] 既存IEモードEdgeを終了: {', '.join(map(str, pids))}")


def _get_ie_mode_edge_pids():
    """IEモード起動中のmsedge.exe PID一覧を取得する"""
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
        pids = set()
        for line in result.stdout.strip().splitlines():
            line = line.strip()
            if line.isdigit():
                pids.add(int(line))
        return pids
    except Exception as e:
        log(f"  [WARN] IEモードEdge PID取得に失敗: {e}")
        return set()


def _snapshot_ie_mode_edges():
    """起動前のIEモードEdge PIDを記録する"""
    return _get_ie_mode_edge_pids()


def _track_new_ie_mode_edges(before_pids):
    """起動前後を比較し、新しく生まれたPIDを記録する"""
    global _tracked_edge_pids
    after_pids = _get_ie_mode_edge_pids()
    new_pids = after_pids - before_pids
    _tracked_edge_pids.update(new_pids)
    log(f"  [DEBUG] 新規IEモードEdge PID: {new_pids if new_pids else 'なし'}")


def _cleanup_tracked_ie_mode_edges():
    """このプログラムが起動したIEモードEdgeだけを終了する"""
    if not _tracked_edge_pids:
        return
    log(f"  [CLEANUP] 終了対象IEモードEdge PID: {_tracked_edge_pids}")
    for pid in _tracked_edge_pids:
        subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    time.sleep(1)
    log("[OK] プログラム起動分のIEモードEdgeをクリーンアップ")


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
    options.initial_browser_url = f"{BASE_URL}/login"
    # ie.browserCommandLineSwitches は環境差が大きく、今回は使用しない
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


def wait_win32_dialog(title, class_name=WIN32_CLASS_DIALOG, timeout=15):
    """Win32ダイアログを待って返す"""
    desktop = Desktop(backend="win32")
    dialog = desktop.window(class_name=class_name, title=title)
    dialog.wait("visible", timeout=timeout)
    return dialog


def log_dialog_info(dialog, label):
    """ダイアログのタイトル/テキスト/ボタン一覧をログ出力する"""
    try:
        title = dialog.window_text()
    except Exception:
        title = "(unknown)"
    try:
        texts = [w.window_text() for w in dialog.descendants() if w.window_text()]
    except Exception:
        texts = []
    log(f"  [DEBUG] {label}検出: title={title}")
    if texts:
        log(f"  [DEBUG] {label}内テキスト: {texts}")
    try:
        buttons = [b for b in dialog.descendants()
                   if b.friendly_class_name() == "Button" and b.window_text()]
        btn_texts = [b.window_text() for b in buttons]
    except Exception:
        buttons = []
        btn_texts = []
    if btn_texts:
        log(f"  [DEBUG] {label}のボタン一覧: {btn_texts}")
    return buttons, btn_texts


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

    userid_input = WebDriverWait(driver, WAIT_LOGIN_PAGE).until(
        EC.visibility_of_element_located(LOC_USER_ID)
    )
    password_input = WebDriverWait(driver, WAIT_LOGIN_PAGE).until(
        EC.visibility_of_element_located(LOC_PASSWORD)
    )

    set_value_with_fallback(userid_input, USER_ID)
    set_value_with_fallback(password_input, PASSWORD)

    try:
        uid_val = userid_input.get_attribute("value")
        if uid_val != USER_ID:
            raise RuntimeError(f"ユーザーIDが一致しません: '{uid_val}'")
        log(f"  [DEBUG] ユーザーID確認OK: '{uid_val}'")
    except Exception as e:
        log(f"  [WARN] ユーザーIDの確認に失敗: {e}")

    try:
        pwd_val = password_input.get_attribute("value")
        if not pwd_val:
            log("  [WARN] パスワードの値を取得できない/空です")
    except Exception as e:
        log(f"  [WARN] パスワードの確認に失敗: {e}")

    # IEモードでは click が失敗しやすいので OSレベルの Enter で送信
    from pywinauto import keyboard as kbd
    kbd.send_keys("{ENTER}")

    # 遷移完了待ち（readyState + 要素出現の両方）
    try:
        wait_for_ready(driver, timeout=WAIT_POST_LOGIN)
    except Exception:
        pass
    WebDriverWait(driver, WAIT_POST_LOGIN).until(
        EC.presence_of_element_located(LOC_DOWNLOAD_LINK)
    )
    log("[OK] ログイン完了 → ダウンロードページへ遷移")


def _handle_confirm_dialog_pywinauto():
    """confirmダイアログをpywinautoで閉じる"""
    dialog = wait_win32_dialog(TITLE_CONFIRM, timeout=WAIT_CONFIRM_DIALOG)
    buttons, btn_texts = log_dialog_info(dialog, "confirmダイアログ")
    dialog.set_focus()
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
    log("  [DEBUG] confirmダイアログで「OK」を選択")
    dialog.wait_not("visible", timeout=WAIT_DIALOG_CLOSE)
    log("[OK] confirmダイアログでOKを選択 (pywinauto)")


def step_click_download_and_confirm(driver):
    """
    ダウンロードリンクをクリックし、confirmダイアログでOKを押す

    Selenium Alert API は環境によってハングするため、Win32ダイアログ操作を優先する。
    """
    t = threading.Thread(target=_handle_confirm_dialog_pywinauto, daemon=True)
    t.start()

    link = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located(LOC_DOWNLOAD_LINK)
    )
    # IEモードでは click が失敗しやすいので Enter でリンクを起動
    try:
        link._parent.execute_script("arguments[0].focus();", link)
    except Exception:
        pass
    time.sleep(0.2)
    from pywinauto import keyboard as kbd
    kbd.send_keys("{ENTER}")
    log("[OK] ダウンロードリンクを実行 (Enter)")
    t.join(timeout=WAIT_CONFIRM_DIALOG + 5)
    if t.is_alive():
        raise TimeoutError("confirmダイアログ処理が完了しませんでした")


def _find_ie_window(desktop, timeout=15):
    """IEモードのウィンドウを探して返す（先頭が「ダウンロード」）"""
    patterns = [
        IE_WINDOW_TITLE_RE,
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

            notification_bar = ie_spec.child_window(
                auto_id=UIA_NOTIFICATION_BAR_ID,
                control_type="ToolBar",
            )
            notification_bar.wait("visible", timeout=WAIT_NOTIFICATION_BAR)

            # キーボード操作のみで「名前を付けて保存」を選択する
            save_dialog = desktop.window(title=TITLE_SAVE_DIALOG)
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
                if save_dialog.exists(timeout=WAIT_SAVE_DIALOG):
                    log("[OK] ダウンロードバーで「名前を付けて保存」を選択 (キーボード)")
                    return
                # 反応しなかった場合はメニューを閉じて再試行
                kbd.send_keys("{ESC}")
                time.sleep(0.5)
                log(f"  [WARN] 保存ダイアログ未表示。再試行 {attempt + 1}/3")
            raise RuntimeError("ダウンロードバーで「名前を付けて保存」を起動できませんでした")
        except Exception as e:
            log(f"  [WARN] ダウンロードバー取得に失敗。再試行 {retry + 1}/3: {e}")
            time.sleep(1.0)

    raise RuntimeError("ダウンロードバーの取得に繰り返し失敗しました")


def step_handle_save_dialog():
    """
    「名前を付けて保存」ダイアログでファイルパスを指定して保存する
    """
    from pywinauto import keyboard as kbd

    desktop = Desktop(backend="uia")

    save_dialog = desktop.window(title=TITLE_SAVE_DIALOG)
    save_dialog.wait("visible", timeout=WAIT_SAVE_DIALOG)
    save_dialog.set_focus()
    time.sleep(0.5)

    save_file_path = os.path.join(SAVE_PATH, SAVE_FILENAME)
    before_mtime = os.path.getmtime(save_file_path) if os.path.exists(save_file_path) else None
    os.makedirs(SAVE_PATH, exist_ok=True)
    log(f"  [DEBUG] 保存先: {save_file_path}")

    try:
        fn_host = save_dialog.child_window(auto_id=SAVE_FILENAME_CONTROL_ID)
        fn_edit = fn_host.child_window(control_type="Edit")
        fn_edit.set_edit_text(save_file_path)
        log("  [DEBUG] ファイル名設定: FileNameControlHost経由")
    except Exception as e1:
        raise RuntimeError(f"ファイル名フィールドへの入力に失敗しました: {e1}")

    time.sleep(0.5)

    try:
        save_btn = save_dialog.child_window(auto_id=SAVE_BUTTON_AUTO_ID, control_type="Button")
        download_start = time.time()
        save_btn.click_input()
        log("  [DEBUG] 保存ボタン押下: auto_id=1経由")
    except Exception as e1:
        raise RuntimeError(f"保存ボタンの押下に失敗しました: {e1}")

    # 上書き確認ダイアログの検出と確実なクリック
    overwrite_clicked = False
    try:
        win32 = Desktop(backend="win32")
        confirm_overwrite = win32.window(class_name=WIN32_CLASS_DIALOG, title=TITLE_OVERWRITE_DIALOG)
        if confirm_overwrite.exists(timeout=WAIT_DIALOG_CLOSE):
            buttons, btn_texts = log_dialog_info(confirm_overwrite, "上書き確認ダイアログ")

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
                log(f"  [DEBUG] 上書き確認ダイアログで「{target}」を選択")
                # 上書き確認の押下後をダウンロード開始時刻とみなす
                download_start = time.time()
                overwrite_clicked = True
            elif len(matched) == 0:
                raise RuntimeError("上書き確認ダイアログのボタンが見つかりません: はい(&Y) / はい(Y)")
            else:
                raise RuntimeError(f"上書き確認ダイアログのボタンが複数一致: {matched}")
            confirm_overwrite.wait_not("visible", timeout=WAIT_DIALOG_CLOSE)
    except Exception as e:
        log(f"  [WARN] 上書き確認処理で例外: {e}")

    # ダイアログが閉じるのを待機
    save_dialog.wait_not("visible", timeout=WAIT_DIALOG_CLOSE)
    time.sleep(0.5)

    if download_start is None:
        download_start = time.time()
    if overwrite_clicked:
        log("  [DEBUG] 上書き確認完了 → 保存完了待ちへ移行")
    log("[OK] 「名前を付けて保存」ダイアログで保存を実行")
    return save_file_path, before_mtime, download_start


def wait_for_download_complete(save_file_path, start_time, timeout=60, stable_sec=3):
    """partialファイル消滅と本体の更新を待つ（開始時刻以降のもののみ対象）"""
    partial_path = f"{save_file_path}.partial"
    end = time.time() + timeout
    try:
        start_time_local = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(start_time))
    except Exception:
        start_time_local = str(start_time)
    log(
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

        elapsed = time.time() - start_time
        if file_exists and not partial_exists and elapsed >= 1:
            return True
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
    before_edge_pids = set()
    try:
        init_logging()
        log("IEDriver + Edge IEモードを起動中...")
        _kill_existing_ie_mode_edges()
        before_edge_pids = _snapshot_ie_mode_edges()
        driver = create_driver()
        time.sleep(2)
        _track_new_ie_mode_edges(before_edge_pids)
        log("[OK] WebDriver起動完了")

        step_login(driver)
        step_click_download_and_confirm(driver)
        step_handle_download_bar()
        save_file_path, before_mtime, download_start = step_handle_save_dialog()

        # ダウンロード完了待ち (.partial が消えるまで)
        wait_for_download_complete(
            save_file_path,
            download_start,
            timeout=WAIT_DOWNLOAD_TIMEOUT,
            stable_sec=WAIT_STABLE_SEC,
        )

        if os.path.exists(save_file_path):
            file_size = os.path.getsize(save_file_path)
            after_mtime = os.path.getmtime(save_file_path)
            if before_mtime is None:
                log(f"[OK] ファイル保存確認: {save_file_path} ({file_size} bytes)")
            else:
                if after_mtime > before_mtime:
                    log(f"[OK] ファイル更新確認: {save_file_path} ({file_size} bytes)")
                else:
                    log(f"[WARN] ファイルが更新されていない可能性: {save_file_path} ({file_size} bytes)")
        else:
            log(f"[WARN] ファイルが見つかりません: {save_file_path}")

        log("\n===== テスト完了 =====")
        time.sleep(2)
    except Exception as e:
        log(f"\n[ERROR] テスト失敗: {e}")
        raise
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        _cleanup_tracked_ie_mode_edges()
        log("ブラウザを終了しました")


if __name__ == "__main__":
    main()
