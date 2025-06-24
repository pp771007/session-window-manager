import tkinter as tk
from tkinter import messagebox
import pygetwindow as gw
import win32process
import win32gui

# --- 設定 ---
APP_TITLE = "視窗佈局管理員"
# psutil 不再是核心必需品，但保留它用於未來可能的除錯
# from collections import defaultdict 也不再需要

saved_layout_in_memory = []

def get_window_hwnd(win):
    """安全地獲取視窗的句柄 (HWND)"""
    try:
        # _hWnd 是 pygetwindow 儲存視窗句柄的內部屬性
        return win._hWnd
    except Exception:
        return None

def save_window_positions(silent=False):
    """掃描所有視窗並基於 HWND 儲存它們的佈局。"""
    global saved_layout_in_memory
    
    current_layout = []
    for win in gw.getAllWindows():
        if win.isMinimized or not win.visible or not win.title or win.title == APP_TITLE:
            continue

        hwnd = get_window_hwnd(win)
        if hwnd:
            current_layout.append({
                "hwnd": hwnd,
                "title_at_save": win.title, # 僅供錯誤報告使用
                "left": win.left,
                "top": win.top,
                "width": win.width,
                "height": win.height,
            })

    saved_layout_in_memory = current_layout
    
    if not silent:
        messagebox.showinfo("成功", f"成功記錄 {len(saved_layout_in_memory)} 個視窗佈局！")
    else:
        print(f"初始佈局已自動記錄，共 {len(saved_layout_in_memory)} 個視窗。")

def reset_window_positions():
    """從記憶體讀取佈局並基於 HWND 恢復視窗位置。"""
    global saved_layout_in_memory

    if not saved_layout_in_memory:
        messagebox.showwarning("警告", "記憶體中沒有佈局紀錄。")
        return

    # 建立一個從當前 HWND 到視窗物件的快速查找地圖
    current_window_map = {get_window_hwnd(win): win for win in gw.getAllWindows()}
    
    restored_count = 0
    permission_denied_titles = []
    not_found_titles = []

    for saved_win_data in saved_layout_in_memory:
        hwnd = saved_win_data["hwnd"]
        title = saved_win_data["title_at_save"]

        if hwnd in current_window_map:
            # 視窗仍然存在，執行恢復
            target_win = current_window_map[hwnd]
            try:
                if target_win.isMaximized: target_win.restore()
                target_win.resizeTo(saved_win_data["width"], saved_win_data["height"])
                target_win.moveTo(saved_win_data["left"], saved_win_data["top"])
                restored_count += 1
            except gw.PyGetWindowException as e:
                if "Error code 5" in str(e) or "存取被拒" in str(e):
                    permission_denied_titles.append(title)
                else:
                    print(f"恢復 '{title}' 時發生視窗控制錯誤: {e}")
            except Exception as e:
                print(f"恢復 '{title}' 時發生未知錯誤: {e}")
        else:
            # HWND 找不到了，說明這個視窗已經被關閉
            not_found_titles.append(title)
            
    # 產生最終報告
    message = f"成功恢復 {restored_count} / {len(saved_layout_in_memory)} 個視窗！"
    if permission_denied_titles:
        message += f"\n\n權限不足，無法移動以下視窗：\n- " + "\n- ".join(set(permission_denied_titles))
    if not_found_titles:
        message += f"\n\n找不到以下已關閉的視窗：\n- " + "\n- ".join(set(not_found_titles))
    if permission_denied_titles:
         message += "\n\n(提示：以系統管理員身分執行可解決權限問題)"
        
    messagebox.showinfo("完成", message)

def create_and_run_gui():
    """建立圖形介面並執行"""
    root = tk.Tk()
    root.title(APP_TITLE)

    try:
        root.iconbitmap('favicon.ico')
    except tk.TclError:
        print("警告：找不到 'favicon.ico' 檔案，將使用預設圖示。")

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

    save_button = tk.Button(frame, text="儲存佈局", command=save_window_positions, font=("Arial", 12), width=25, height=2)
    save_button.pack(pady=5)

    reset_button = tk.Button(frame, text="恢復佈局", command=reset_window_positions, font=("Arial", 12), width=25, height=2)
    reset_button.pack(pady=5)

    root.after(100, lambda: save_window_positions(silent=True))

    root.mainloop()

if __name__ == "__main__":
    create_and_run_gui()