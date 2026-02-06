"""
IWebBrowser2 COM + pywinauto によるIEモードダウンロードテスト自動化スクリプト

IEDriver経由ではEdge 144のIEモードウィンドウを制御できないため、
COM直接操作（InternetExplorer.Application）でIEを起動・操作する。

前提条件:
- Flaskサーバー (app.py) が起動していること (http://localhost:5000)
- comtypes がインストールされていること (pip install comtypes)
"""

import os
import subprocess
import threading
import time
import comtypes.client
from pywinauto import Desktop

# ===== 設定 =====
BASE_URL = "http://localhost:5000"
SAVE_PATH = r"D:\Git\iemode_dl_test\download"
USER_ID = "testuser"
PASSWORD = "testpass"

# このプログラムが起動したプロセスのPIDを記録する
_tracked_pids = set()


def get_pids(process_name):
    """指定プロセス名の現在のPID一覧を取得する"""
    result = subprocess.run(
        ["tasklist", "/FI", f"IMAGENAME eq {process_name}", "/FO", "CSV", "/NH"],
        capture_output=True, text=True,
    )
    pids = set()
    for line in result.stdout.strip().splitlines():
        parts = line.replace('"', '').split(',')
        if len(parts) >= 2 and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def kill_pids(pids):
    """指定したPIDリストのプロセスを強制終了する"""
    for pid in pids:
        subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )


def snapshot_before():
    """ブラウザ起動前のiexplore.exe PIDを記録する"""
    return get_pids("iexplore.exe")


def track_new_pids(before_pids):
    """起動前後を比較し、新しく生まれたPIDを記録する"""
    global _tracked_pids
    after_pids = get_pids("iexplore.exe")
    new_pids = after_pids - before_pids
    _tracked_pids.update(new_pids)
    print(f"  [DEBUG] 新規IEプロセスPID: {new_pids if new_pids else 'なし'}")


def cleanup_tracked():
    """このプログラムが起動したプロセスだけを終了する"""
    if _tracked_pids:
        print(f"  [CLEANUP] 終了対象PID: {_tracked_pids}")
        kill_pids(_tracked_pids)
    time.sleep(1)
    print("[OK] プログラム起動分のプロセスをクリーンアップ")


def create_ie():
    """IWebBrowser2 COMオブジェクトを生成し、ブラウザを表示する"""
    ie = comtypes.client.CreateObject("InternetExplorer.Application")
    ie.Visible = True
    return ie


def wait_for_ready(ie, timeout=15):
    """IEのReadyState==4 (READYSTATE_COMPLETE) を待機する"""
    start = time.time()
    while time.time() - start < timeout:
        if ie.ReadyState == 4:
            return True
        time.sleep(0.5)
    raise TimeoutError(f"ページ読み込みが{timeout}秒以内に完了しませんでした")


def navigate(ie, url):
    """指定URLに遷移し、読み込み完了を待つ"""
    ie.Navigate(url)
    time.sleep(1)
    wait_for_ready(ie)


def get_document(ie):
    """IEのHTMLDocumentオブジェクトを取得する"""
    return ie.Document


def step_login(ie):
    """ログインページでユーザーID・パスワードを入力し、ログインボタンを押す"""
    navigate(ie, f"{BASE_URL}/login")

    doc = get_document(ie)
    # class名で要素を取得
    userid_inputs = doc.getElementsByClassName("txtUserID")
    password_inputs = doc.getElementsByClassName("txtPassWord")

    if userid_inputs.length == 0 or password_inputs.length == 0:
        raise RuntimeError("ログインフォームの要素が見つかりません")

    userid_inputs.item(0).value = USER_ID
    password_inputs.item(0).value = PASSWORD

    # ログインボタンをクリック
    buttons = doc.getElementsByTagName("button")
    buttons.item(0).click()

    # ダウンロードページへの遷移を待機
    time.sleep(2)
    wait_for_ready(ie)
    print("[OK] ログイン完了 → ダウンロードページへ遷移")


def _handle_confirm_dialog_thread():
    """別スレッドでconfirmダイアログを待機してOKを押す"""
    desktop = Desktop(backend="win32")
    dialog = desktop.window(class_name="#32770", title_re=".*Web.*")
    dialog.wait("visible", timeout=15)
    dialog.set_focus()
    dialog["OK"].click()
    print("[OK] confirmダイアログでOKを選択")


def step_click_download_and_confirm(ie):
    """
    ダウンロードリンクをクリックし、confirmダイアログでOKを押す

    COM経由の link.click() は confirm() が閉じるまでブロックするため、
    ダイアログ処理を別スレッドで先に起動してからクリックする。
    """
    # 先にダイアログ処理スレッドを起動
    t = threading.Thread(target=_handle_confirm_dialog_thread, daemon=True)
    t.start()

    # リンクをクリック (confirmダイアログが出てブロックされる → 別スレッドがOKを押す → 戻る)
    doc = get_document(ie)
    links = doc.getElementsByTagName("a")
    for i in range(links.length):
        link = links.item(i)
        href = link.getAttribute("href")
        if href and "/download/csv" in str(href):
            link.click()
            print("[OK] ダウンロードリンクをクリック")
            t.join(timeout=15)
            return
    raise RuntimeError("ダウンロードリンクが見つかりません")


def step_handle_download_bar():
    """
    IEのダウンロード通知バーで「名前を付けて保存」を実行する

    通知バーの構造:
      Toolbar '通知' (IENotificationBar)
        ├── Static '通知バーのテキスト'
        ├── Button 'ファイルを開く'
        ├── SplitButton '保存'
        │     └── SplitButton '6'  ← ドロップダウン矢印
        ├── Button 'キャンセル'
        └── Button '閉じる'
    """
    desktop = Desktop(backend="uia")
    time.sleep(3)  # ダウンロードバーの表示を待機

    # IEウィンドウを検出・最前面化
    ie_window = desktop.window(title_re=".*Internet Explorer.*")
    ie_window.wait("visible", timeout=15)
    ie_window.set_focus()

    # 通知バー検出
    notification_bar = ie_window.child_window(auto_id="IENotificationBar",
                                               control_type="ToolBar")
    notification_bar.wait("visible", timeout=10)

    # 「保存」SplitButton内のドロップダウン矢印（子SplitButton title="6"）をクリック
    save_button = notification_bar.child_window(title="保存", control_type="SplitButton")
    save_button.wait("visible", timeout=10)
    dropdown = save_button.child_window(title="6", control_type="SplitButton")
    dropdown.click_input()
    time.sleep(1)

    # 「名前を付けて保存」メニュー項目をクリック
    save_as_item = desktop.window(control_type="Menu").child_window(
        title="名前を付けて保存(A)", control_type="MenuItem"
    )
    save_as_item.click_input()
    print("[OK] ダウンロードバーで「名前を付けて保存」を選択")


def step_handle_save_dialog():
    """
    「名前を付けて保存」ダイアログでファイルパスを指定して保存する

    Windows標準のファイル保存ダイアログ (#32770) を操作する。
    ダイアログ要素のタイトルやauto_idはOS・言語設定で異なるため、
    複数のフォールバック戦略を使用する。
    """
    from pywinauto import keyboard as kbd

    desktop = Desktop(backend="uia")

    save_dialog = desktop.window(title="名前を付けて保存")
    save_dialog.wait("visible", timeout=15)
    save_dialog.set_focus()
    time.sleep(1)

    # 保存先ファイルパス
    save_file_path = os.path.join(SAVE_PATH, "sample.csv")
    os.makedirs(SAVE_PATH, exist_ok=True)
    print(f"  [DEBUG] 保存先: {save_file_path}")

    # --- ファイル名入力 ---
    # 方法1: auto_id="FileNameControlHost" 内の Edit を探す
    # 方法2: title に "ファイル名" を含む ComboBox/Edit を探す
    # 方法3: キーボード操作 (Alt+N でファイル名フィールドにフォーカス)
    file_name_set = False

    # 方法1: FileNameControlHost
    try:
        fn_host = save_dialog.child_window(auto_id="FileNameControlHost")
        fn_edit = fn_host.child_window(control_type="Edit")
        fn_edit.set_edit_text(save_file_path)
        file_name_set = True
        print("  [DEBUG] ファイル名設定: FileNameControlHost経由")
    except Exception as e1:
        print(f"  [DEBUG] FileNameControlHost失敗: {e1}")

    # 方法2: title_re でファイル名フィールドを探す
    if not file_name_set:
        try:
            fn_edit = save_dialog.child_window(title_re=".*ファイル名.*",
                                                control_type="Edit")
            fn_edit.set_edit_text(save_file_path)
            file_name_set = True
            print("  [DEBUG] ファイル名設定: title_re経由")
        except Exception as e2:
            print(f"  [DEBUG] title_re(Edit)失敗: {e2}")

    # 方法3: キーボードで直接入力 (Alt+N → ファイル名欄にフォーカス)
    if not file_name_set:
        try:
            save_dialog.set_focus()
            time.sleep(0.5)
            kbd.send_keys("%n")  # Alt+N
            time.sleep(0.5)
            kbd.send_keys("^a")  # Ctrl+A (全選択)
            kbd.send_keys(save_file_path, with_spaces=True)
            file_name_set = True
            print("  [DEBUG] ファイル名設定: キーボード(Alt+N)経由")
        except Exception as e3:
            print(f"  [DEBUG] キーボード入力失敗: {e3}")

    if not file_name_set:
        raise RuntimeError("ファイル名フィールドへの入力に失敗しました")

    time.sleep(0.5)

    # --- 保存ボタン押下 ---
    # 方法1: auto_id="1" (標準ダイアログのIDOK)
    # 方法2: title_re で保存ボタンを探す
    # 方法3: Alt+S キーボードショートカット
    # 方法4: Enter キー
    save_clicked = False

    # 方法1: auto_id="1"
    try:
        save_btn = save_dialog.child_window(auto_id="1", control_type="Button")
        save_btn.click_input()
        save_clicked = True
        print("  [DEBUG] 保存ボタン押下: auto_id=1経由")
    except Exception as e1:
        print(f"  [DEBUG] auto_id=1失敗: {e1}")

    # 方法2: title_re
    if not save_clicked:
        try:
            save_btn = save_dialog.child_window(title_re=".*保存.*",
                                                 control_type="Button")
            save_btn.click_input()
            save_clicked = True
            print("  [DEBUG] 保存ボタン押下: title_re経由")
        except Exception as e2:
            print(f"  [DEBUG] title_re(Button)失敗: {e2}")

    # 方法3: Alt+S キーボードショートカット
    if not save_clicked:
        try:
            save_dialog.set_focus()
            time.sleep(0.3)
            kbd.send_keys("%s")  # Alt+S
            save_clicked = True
            print("  [DEBUG] 保存ボタン押下: Alt+S経由")
        except Exception as e3:
            print(f"  [DEBUG] Alt+Sキー送信失敗: {e3}")

    # 方法4: Enter キー
    if not save_clicked:
        try:
            kbd.send_keys("{ENTER}")
            save_clicked = True
            print("  [DEBUG] 保存ボタン押下: Enterキー経由")
        except Exception as e4:
            print(f"  [DEBUG] Enterキー送信失敗: {e4}")

    if not save_clicked:
        raise RuntimeError("保存ボタンの押下に失敗しました")

    time.sleep(2)

    # 「上書き確認」ダイアログが出た場合に対応
    try:
        confirm_overwrite = desktop.window(title="名前を付けて保存の確認")
        if confirm_overwrite.exists(timeout=2):
            confirm_overwrite.child_window(title="はい(&Y)",
                                            control_type="Button").click_input()
            print("  [DEBUG] 上書き確認ダイアログで「はい」を選択")
            time.sleep(1)
    except Exception:
        pass  # 上書き確認が出なければスキップ

    print("[OK] 「名前を付けて保存」ダイアログで保存を実行")


def main():
    ie = None
    try:
        # Step 0: PID記録
        before_pids = snapshot_before()
        print(f"  [DEBUG] 起動前のIE PID数: {len(before_pids)}")

        # Step 1: IE起動 (COM経由)
        print("IE (IWebBrowser2 COM) を起動中...")
        ie = create_ie()
        time.sleep(2)
        track_new_pids(before_pids)
        print("[OK] IE起動完了")

        # Step 2: ログイン
        step_login(ie)

        # Step 3-4: ダウンロードリンククリック + confirmダイアログ処理
        step_click_download_and_confirm(ie)

        # Step 5: ダウンロードバー処理
        step_handle_download_bar()

        # Step 6: 名前を付けて保存
        step_handle_save_dialog()

        # Step 7: ファイル存在確認
        save_file_path = os.path.join(SAVE_PATH, "sample.csv")
        if os.path.exists(save_file_path):
            file_size = os.path.getsize(save_file_path)
            print(f"[OK] ファイル保存確認: {save_file_path} ({file_size} bytes)")
        else:
            print(f"[WARN] ファイルが見つかりません: {save_file_path}")

        print("\n===== テスト完了 =====")
        time.sleep(2)

    except Exception as e:
        print(f"\n[ERROR] テスト失敗: {e}")
        raise
    finally:
        if ie:
            try:
                ie.Quit()
            except Exception:
                pass
        cleanup_tracked()
        print("ブラウザを終了しました")


if __name__ == "__main__":
    main()
