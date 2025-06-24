import tkinter as tk
from tkinter import messagebox
import win32gui
import win32con

# --- 全域變數 ---
APP_TITLE = "視窗佈局管理員"
saved_layout_in_memory = {}
saved_z_order = []

def save_window_positions(silent=False):
    """儲存所有可見視窗的位置、大小和Z順序。"""
    global saved_layout_in_memory, saved_z_order

    saved_layout_in_memory.clear()
    saved_z_order.clear()

    win = win32gui.GetTopWindow(0)
    while win:
        # 使用 GetWindowPlacement 檢查視窗是否最小化
        placement = win32gui.GetWindowPlacement(win)
        is_minimized = placement[1] == win32con.SW_SHOWMINIMIZED

        if win32gui.IsWindowVisible(win) and win32gui.GetWindowText(win) and not is_minimized:
            title = win32gui.GetWindowText(win)
            if title != APP_TITLE:
                rect = win32gui.GetWindowRect(win)
                left, top, right, bottom = rect
                saved_layout_in_memory[win] = {
                    "title_at_save": title,
                    "left": left, "top": top,
                    "width": right - left, "height": bottom - top
                }
                saved_z_order.append(win)
        
        win = win32gui.GetWindow(win, win32con.GW_HWNDNEXT)
    
    if not silent:
        messagebox.showinfo("成功", f"成功記錄 {len(saved_z_order)} 個視窗的佈局與順序！")
    else:
        print(f"初始佈局與順序已自動記錄，共 {len(saved_z_order)} 個視窗。")

def reset_window_positions():
    """恢復所有已記錄視窗的位置、大小和Z順序。"""
    global saved_layout_in_memory, saved_z_order

    if not saved_layout_in_memory or not saved_z_order:
        messagebox.showwarning("警告", "記憶體中沒有佈局紀錄。")
        return

    restored_count = 0
    permission_denied_titles = []
    
    current_hwnds = {hwnd for hwnd, layout in saved_layout_in_memory.items() if win32gui.IsWindow(hwnd)}
    not_found_hwnds = set(saved_layout_in_memory.keys()) - current_hwnds

    # 1. 恢復位置和大小
    for hwnd, layout in saved_layout_in_memory.items():
        if hwnd in not_found_hwnds:
            continue
        try:
            # --- 核心修改點 ---
            # 使用 GetWindowPlacement 來檢查視窗狀態 (最小化或最大化)
            placement = win32gui.GetWindowPlacement(hwnd)
            if placement[1] == win32con.SW_SHOWMINIMIZED or placement[1] == win32con.SW_SHOWMAXIMIZED:
                # 如果是最小化或最大化，先恢復到正常狀態
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 
                                 layout['left'], layout['top'], 
                                 layout['width'], layout['height'], 0)
            restored_count += 1
        except Exception as e:
            if hasattr(e, 'winerror') and e.winerror == 5:
                permission_denied_titles.append(layout['title_at_save'])
            else:
                print(f"恢復 '{layout['title_at_save']}' 位置時出錯: {e}")

    # 2. 恢復 Z-Order
    for i in range(len(saved_z_order) - 2, -1, -1):
        hwnd_to_place = saved_z_order[i]
        hwnd_after = saved_z_order[i+1]
        
        if hwnd_to_place not in current_hwnds or hwnd_after not in current_hwnds:
            continue
        try:
            win32gui.SetWindowPos(hwnd_to_place, hwnd_after, 0, 0, 0, 0,
                                 win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        except Exception:
            pass 

    # 3. 恢復焦點
    if saved_z_order and saved_z_order[0] in current_hwnds:
        try:
            top_hwnd = saved_z_order[0]
            win32gui.SetForegroundWindow(root.winfo_id()) 
            win32gui.SetForegroundWindow(top_hwnd)
        except Exception:
            pass
            
    # 產生報告
    not_found_titles = [saved_layout_in_memory[hwnd]['title_at_save'] for hwnd in not_found_hwnds]
    message = f"成功恢復 {restored_count} / {len(saved_layout_in_memory)} 個視窗！"
    if permission_denied_titles:
        message += f"\n\n權限不足，無法移動以下視窗：\n- " + "\n- ".join(set(permission_denied_titles))
    if not_found_titles:
        message += f"\n\n找不到以下已關閉的視窗：\n- " + "\n- ".join(set(not_found_titles))
    if permission_denied_titles:
         message += "\n\n(提示：以系統管理員身分執行可解決權限問題)"
        
    messagebox.showinfo("完成", message)

def create_and_run_gui():
    global root
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