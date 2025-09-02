import tkinter as tk
from tkinter import messagebox
import win32gui
import win32con
import time
import threading
import sys
import os

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class WindowLayoutManager:
    def __init__(self, root):
        self.root = root
        self.APP_TITLE = "視窗佈局管理員"
        
        self.saved_layout = {}
        self.saved_z_order = []

        self._setup_gui()
        
        self.root.after(100, lambda: self.save_window_positions(silent=True))
        self.root.after(60000, self.auto_update_layout)

    def _setup_gui(self):
        self.root.title(self.APP_TITLE)
        try:
            self.root.iconbitmap(resource_path('favicon.ico'))
        except tk.TclError:
            print("警告：找不到 'favicon.ico' 檔案，將使用預設圖示。")

        window_width, window_height = 350, 160
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.root.resizable(False, False)

        frame = tk.Frame(self.root, padx=20, pady=10)
        frame.pack(expand=True, fill='both')

        save_button = tk.Button(frame, text="手動儲存佈局 (覆蓋)", command=self.save_window_positions, 
                                font=("Arial", 12), width=25, height=2, 
                                bg="#cce5ff", activebackground="#b8daff", relief=tk.FLAT)
        save_button.pack(pady=5)
        
        self.restore_button = tk.Button(frame, text="恢復佈局", command=self.restore_window_positions_threaded, 
                                        font=("Arial", 12), width=25, height=2, 
                                        bg="#d4edda", activebackground="#c3e6cb", relief=tk.FLAT)
        self.restore_button.pack(pady=5)
        
        self.status_var = tk.StringVar()
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W, padx=5)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self._set_status("正在初始化...")

    def _set_status(self, message):
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        self.status_var.set(f"[{timestamp}] {message}")

    def _get_all_windows_with_zorder(self):
        """獲取所有視窗並按正確的 Z-order 排序（由上到下）"""
        windows_info = []
        
        def enum_windows_proc(hwnd, lParam):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:  # 只處理有標題的視窗
                    placement = win32gui.GetWindowPlacement(hwnd)
                    is_minimized = placement[1] == win32con.SW_SHOWMINIMIZED
                    if not is_minimized:
                        windows_info.append(hwnd)
            return True
        
        # 使用 EnumWindows 來獲取正確的 Z-order
        win32gui.EnumWindows(enum_windows_proc, 0)
        return windows_info

    def _get_all_windows(self):
        """保留原有方法作為備用"""
        return self._get_all_windows_with_zorder()

    def save_window_positions(self, silent=False):
        self.saved_layout.clear()
        self.saved_z_order.clear()

        # 使用新的方法獲取視窗
        all_windows = self._get_all_windows_with_zorder()
        self.saved_z_order = all_windows.copy()

        for hwnd in all_windows:
            title = win32gui.GetWindowText(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            left, top, right, bottom = rect
            self.saved_layout[hwnd] = {
                "title_at_save": title,
                "left": left, "top": top,
                "width": right - left, "height": bottom - top
            }
        
        status_msg = f"佈局已儲存，共記錄 {len(self.saved_z_order)} 個視窗。"
        self._set_status(status_msg)
        
        if silent:
            print(f"初始佈局與順序已自動記錄，共 {len(self.saved_z_order)} 個視窗。")
        else:
            print(status_msg + " (舊紀錄已被覆蓋)")
            # 印出 Z-order 供除錯
            print("Z-order (由上到下):")
            for i, hwnd in enumerate(self.saved_z_order):
                title = win32gui.GetWindowText(hwnd)
                print(f"  {i+1}. {title}")

    def restore_window_positions_threaded(self):
        self._set_status("正在恢復佈局...")
        
        thread = threading.Thread(target=self.restore_window_positions)
        thread.start()

    def _restore_zorder(self, current_hwnds):
        """恢復 Z-order，由底層往上層處理"""
        print("開始恢復 Z-order...")
        
        # 過濾出仍存在的視窗
        valid_zorder = [hwnd for hwnd in self.saved_z_order if hwnd in current_hwnds]
        
        if not valid_zorder:
            print("  沒有有效的視窗需要恢復 Z-order")
            return
        
        try:
            # 從最底層開始，將每個視窗設到最上層
            # 這樣最後處理的視窗會在最上層
            for i, hwnd in enumerate(reversed(valid_zorder)):
                try:
                    title = win32gui.GetWindowText(hwnd)
                    print(f"  設定 Z-order ({len(valid_zorder)-i}/{len(valid_zorder)}): {title}")
                    
                    # 先確保視窗可見且未最小化
                    placement = win32gui.GetWindowPlacement(hwnd)
                    if placement[1] == win32con.SW_SHOWMINIMIZED:
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    
                    # 將視窗移到最上層
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 
                                        0, 0, 0, 0,
                                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
                    
                    time.sleep(0.02)  # 短暫延遲讓系統處理
                    
                except Exception as e:
                    print(f"    設定 {title} Z-order 失敗: {e}")
            
            print("Z-order 恢復完成")
            
        except Exception as e:
            print(f"恢復 Z-order 時發生錯誤: {e}")

    def _finalize_restore(self, restored_count, num_closed_windows, permission_denied_titles, current_hwnds_len):
        has_issues = bool(permission_denied_titles)

        status_message = f"成功恢復 {restored_count} 個視窗"
        if num_closed_windows > 0:
            status_message += f" ({num_closed_windows} 個已不存在)"

        if not has_issues:
            self._set_status(status_message + "。")
        else:
            status_message += f"，另有 {len(permission_denied_titles)} 個權限問題。"
            self._set_status(status_message)

            message = f"成功恢復 {restored_count} / {current_hwnds_len} 個視窗！"
            if permission_denied_titles:
                message += f"\n\n權限不足，無法移動以下視窗：\n- " + "\n- ".join(sorted(set(permission_denied_titles)))
                message += "\n\n(提示：以系統管理員身分執行可解決權限問題)"
            
            messagebox.showinfo("恢復報告", message)

    def restore_window_positions(self):
        restored_count = 0
        permission_denied_titles = []

        # 清理已關閉的視窗
        closed_hwnds = [hwnd for hwnd in self.saved_layout if not win32gui.IsWindow(hwnd)]
        num_closed_windows = len(closed_hwnds)
        for hwnd in closed_hwnds:
            del self.saved_layout[hwnd]

        current_hwnds = set(self.saved_layout.keys())
        
        print(f"開始恢復 {len(current_hwnds)} 個視窗的位置...")
        
        # 先恢復位置和大小
        for hwnd, layout in self.saved_layout.items():
            try:
                title = layout.get('title_at_save', '未知視窗')
                self.root.after(0, self._set_status, f"正在恢復位置: {title}")
                time.sleep(0.05)

                placement = win32gui.GetWindowPlacement(hwnd)
                if placement[1] in (win32con.SW_SHOWMINIMIZED, win32con.SW_SHOWMAXIMIZED):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

                rect = win32gui.GetWindowRect(hwnd)
                current_left, current_top, current_right, current_bottom = rect
                current_width = current_right - current_left
                current_height = current_bottom - current_top

                # 判斷是否需要調整位置或大小
                if not (current_left == layout['left'] and
                        current_top == layout['top'] and
                        current_width == layout['width'] and
                        current_height == layout['height']):
                    
                    print(f"  調整視窗位置: {title}")
                    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST,
                                        layout['left'], layout['top'],
                                        layout['width'], layout['height'], 
                                        win32con.SWP_NOZORDER)  # 先不改變 Z-order
                
                restored_count += 1

            except Exception as e:
                if hasattr(e, 'winerror') and e.winerror == 5:
                    permission_denied_titles.append(layout['title_at_save'])
                else:
                    print(f"恢復 '{layout['title_at_save']}' 位置時出錯: {e}")

        # 位置調整完成後，恢復 Z-order
        print("位置恢復完成，開始恢復視窗層級...")
        self.root.after(0, self._set_status, "正在恢復視窗層級...")
        time.sleep(0.1)  # 讓位置調整完成
        
        self._restore_zorder(current_hwnds)
        
        self.root.after(0, self._finalize_restore, restored_count, num_closed_windows, permission_denied_titles, len(current_hwnds))

    def auto_update_layout(self):
        newly_added_titles = []
        all_windows = self._get_all_windows()
        
        for hwnd in all_windows:
            if hwnd not in self.saved_layout:
                title = win32gui.GetWindowText(hwnd)
                rect = win32gui.GetWindowRect(hwnd)
                left, top, right, bottom = rect
                self.saved_layout[hwnd] = {
                    "title_at_save": title,
                    "left": left, "top": top,
                    "width": right - left, "height": bottom - top
                }
                newly_added_titles.append(title)
                
                # 同時更新 Z-order 列表
                if hwnd not in self.saved_z_order:
                    # 新視窗加入到最上層
                    self.saved_z_order.insert(0, hwnd)

        if newly_added_titles:
            status_msg = f"自動偵測到 {len(newly_added_titles)} 個新視窗。"
            self._set_status(status_msg)
            print(self.status_var.get())
            for title in newly_added_titles:
                print(f"  - {title}")
            print(f"目前共記錄 {len(self.saved_layout)} 個視窗。")

        self.root.after(60000, self.auto_update_layout)

if __name__ == "__main__":
    root = tk.Tk()
    app = WindowLayoutManager(root)
    print("程式已啟動，將在一分鐘後開始自動偵測新視窗...")
    root.mainloop()