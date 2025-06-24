import tkinter as tk
from tkinter import messagebox
import pygetwindow as gw
import win32process
import win32gui
import psutil

# --- 設定 ---
APP_TITLE = "視窗佈局管理員"

# 全域變數，用來在記憶體中儲存佈局
saved_layout_in_memory = []

def get_window_pid(win):
    """透過視窗物件獲取其處理程序ID (PID)"""
    try:
        hwnd = win._hWnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid
    except Exception:
        return None

def save_window_positions(silent=False):
    """掃描所有視窗並將佈局（包含標題）儲存到記憶體中。"""
    global saved_layout_in_memory
    
    all_windows = gw.getAllWindows()
    current_layout = []

    for win in all_windows:
        if win.isMinimized or not win.visible or not win.title or win.title == APP_TITLE:
            continue

        pid = get_window_pid(win)
        if pid:
            layout_info = {
                "pid": pid,
                "title_at_save": win.title,
                "left": win.left,
                "top": win.top,
                "width": win.width,
                "height": win.height,
            }
            if not any(item['pid'] == pid for item in current_layout):
                current_layout.append(layout_info)

    saved_layout_in_memory = current_layout
    
    if not silent:
        messagebox.showinfo("成功", f"成功記錄 {len(saved_layout_in_memory)} 個視窗佈局！")
    else:
        print(f"初始佈局已自動記錄，共 {len(saved_layout_in_memory)} 個視窗。")

def reset_window_positions():
    """從記憶體讀取佈局並恢復視窗位置，錯誤訊息會顯示標題。"""
    global saved_layout_in_memory

    if not saved_layout_in_memory:
        messagebox.showwarning("警告", "記憶體中沒有佈局紀錄。")
        return

    pid_to_window_map = {}
    for win in gw.getAllWindows():
        pid = get_window_pid(win)
        if pid and pid not in pid_to_window_map and win.visible and not win.isMinimized:
            pid_to_window_map[pid] = win

    restored_count = 0
    permission_denied_titles = []

    for saved_win_data in saved_layout_in_memory:
        pid = saved_win_data["pid"]
        title = saved_win_data.get("title_at_save", "未知標題")

        if pid in pid_to_window_map:
            target_win = pid_to_window_map[pid]
            try:
                if target_win.isMaximized: target_win.restore()
                target_win.resizeTo(saved_win_data["width"], saved_win_data["height"])
                target_win.moveTo(saved_win_data["left"], saved_win_data["top"])
                restored_count += 1
            except gw.PyGetWindowException as e:
                if "Error code 5" in str(e) or "存取被拒" in str(e):
                    permission_denied_titles.append(title)
                    print(f"權限不足，跳過 PID {pid} (標題: '{title}')。")
                else:
                    print(f"恢復 PID {pid} (標題: '{title}') 時發生視窗控制錯誤: {e}")
            except Exception as e:
                print(f"恢復 PID {pid} (標題: '{title}') 時發生未知錯誤: {e}")

    message = f"成功恢復 {restored_count} / {len(saved_layout_in_memory)} 個視窗！"
    if permission_denied_titles:
        message += f"\n\n由於權限不足，以下 {len(permission_denied_titles)} 個視窗無法移動：\n"
        formatted_titles = "\n- ".join(permission_denied_titles)
        message += f"- {formatted_titles}"
        message += "\n\n(建議：以系統管理員身分執行此工具)"
        
    messagebox.showinfo("完成", message)

def create_and_run_gui():
    """建立圖形介面並執行"""
    root = tk.Tk()
    root.title(APP_TITLE)

    # --- 設定應用程式圖示 ---
    try:
        root.iconbitmap('favicon.ico')
    except tk.TclError:
        print("警告：找不到 'favicon.ico' 檔案，將使用預設圖示。")
    # -------------------------

    root.geometry("350x200")
    root.resizable(False, False)

    window_width, window_height = 350, 200
    screen_width, screen_height = root.winfo_screenwidth(), root.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    root.attributes('-topmost', True)

    frame = tk.Frame(root, padx=20, pady=10)
    frame.pack(expand=True)
    
    info_label = tk.Label(frame, text="佈局僅存於記憶體，關閉後將遺失", fg="red", font=("Arial", 9))
    info_label.pack(pady=(0, 10))

    # --- 更新按鈕文字 ---
    save_button = tk.Button(frame, text="儲存佈局", command=save_window_positions, font=("Arial", 12), width=25, height=2)
    save_button.pack(pady=5)

    reset_button = tk.Button(frame, text="恢復佈局", command=reset_window_positions, font=("Arial", 12), width=25, height=2)
    reset_button.pack(pady=5)
    # -------------------

    root.after(100, lambda: save_window_positions(silent=True))

    root.mainloop()

if __name__ == "__main__":
    create_and_run_gui()