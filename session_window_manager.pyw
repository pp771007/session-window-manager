import tkinter as tk
from tkinter import messagebox
import win32gui
import win32con
import win32process
import win32api
import time
import threading
import sys
import os

from version import __version__

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
        self.APP_TITLE = f"視窗佈局管理員 {__version__}"
        
        self.saved_layout = {}

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
        bottom_y = int(screen_height - window_height - 78)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{bottom_y}')
        self.root.resizable(False, False)

        frame = tk.Frame(self.root, padx=20, pady=5)
        frame.pack(fill='both', expand=False)

        # 第一行:儲存按鈕和齒輪按鈕
        button_row1 = tk.Frame(frame)
        button_row1.pack(pady=3, fill=tk.X)
        
        save_button = tk.Button(button_row1, text="💾 手動儲存佈局 (覆蓋)", command=self.save_window_positions, 
                                font=("Arial", 12), height=2, 
                                bg="#cce5ff", activebackground="#b8daff", relief=tk.FLAT)
        save_button.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        view_button = tk.Button(button_row1, text="⚙", command=self.open_layout_editor, 
                                font=("Arial", 14), width=4, height=2, 
                                bg="#ffe5b4", activebackground="#ffd699", relief=tk.FLAT)
        view_button.pack(side=tk.LEFT, padx=(5, 0))
        
        self.restore_button = tk.Button(frame, text="🔄 恢復佈局", command=self.restore_window_positions_threaded, 
                                        font=("Arial", 12), height=2, 
                                        bg="#d4edda", activebackground="#c3e6cb", relief=tk.FLAT)
        self.restore_button.pack(pady=3, fill=tk.X)
        
        self.status_var = tk.StringVar()
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W, padx=5)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X, ipady=2)
        self._set_status("正在初始化...")

    def _set_status(self, message):
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        self.status_var.set(f"[{timestamp}] {message}")

    def _should_exclude_window(self, hwnd):
        """檢查視窗是否應該被排除（系統程式）"""
        try:
            # 獲取進程ID
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            
            # 獲取進程路徑
            handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
            exe_path = win32process.GetModuleFileNameEx(handle, 0)
            win32api.CloseHandle(handle)
            
            # 轉換為小寫便於比較
            exe_path_lower = exe_path.lower()
            
            # 排除 Program Manager（桌面背景視窗）：標題為 "Program Manager" 且進程為 explorer.exe
            title = win32gui.GetWindowText(hwnd)
            if title == "Program Manager" and 'explorer.exe' in exe_path_lower:
                return True
            
            # 保留檔案總管視窗
            if 'explorer.exe' in exe_path_lower:
                return False

            # 保留常見終端機（即使位於系統路徑下）
            terminal_exes = ('conhost.exe', 'cmd.exe', 'powershell.exe', 'pwsh.exe',
                             'wt.exe', 'windowsterminal.exe', 'openconsole.exe')
            if os.path.basename(exe_path_lower) in terminal_exes:
                return False

            # 排除系統資料夾下的程式
            system_paths = [
                'c:\\windows\\system32\\',
                'c:\\windows\\systemapps\\',
                'c:\\windows\\immersivecontrolpanel\\',
                'c:\\program files\\windowsapps\\'
            ]
            
            for sys_path in system_paths:
                if exe_path_lower.startswith(sys_path):
                    return True
            
            return False
            
        except Exception:
            # 如果無法獲取進程資訊，預設不排除
            return False

    def _get_all_windows(self):
        """獲取所有可見的有標題視窗（排除系統程式）"""
        windows_info = []
        
        def enum_windows_proc(hwnd, lParam):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:  # 只處理有標題的視窗
                    placement = win32gui.GetWindowPlacement(hwnd)
                    is_minimized = placement[1] == win32con.SW_SHOWMINIMIZED
                    if not is_minimized:
                        # 排除系統程式
                        if not self._should_exclude_window(hwnd):
                            windows_info.append(hwnd)
            return True
        
        win32gui.EnumWindows(enum_windows_proc, 0)
        return windows_info

    def save_window_positions(self, silent=False):
        self.saved_layout.clear()

        all_windows = self._get_all_windows()

        for hwnd in all_windows:
            title = win32gui.GetWindowText(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            left, top, right, bottom = rect
            
            # 獲取執行檔路徑和名稱
            exe_path = ""
            exe_name = ""
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                exe_path = win32process.GetModuleFileNameEx(handle, 0)
                exe_name = os.path.basename(exe_path)
                win32api.CloseHandle(handle)
            except Exception:
                pass
            
            self.saved_layout[hwnd] = {
                "title_at_save": title,
                "left": left, "top": top,
                "width": right - left, "height": bottom - top,
                "exe_path": exe_path,
                "exe_name": exe_name
            }
        
        status_msg = f"佈局已儲存，共記錄 {len(self.saved_layout)} 個視窗。"
        self._set_status(status_msg)
        
        if silent:
            print(f"初始佈局已自動記錄，共 {len(self.saved_layout)} 個視窗。")
        else:
            print(status_msg + " (舊紀錄已被覆蓋)")

    def restore_window_positions_threaded(self):
        self._set_status("正在恢復佈局...")
        
        thread = threading.Thread(target=self.restore_window_positions)
        thread.start()

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
        
        # 恢復位置和大小
        for hwnd, layout in self.saved_layout.items():
            try:
                title = layout.get('title_at_save', '未知視窗')
                self.root.after(0, self._set_status, f"正在恢復位置: {title}")
                time.sleep(0.03)

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
                                        0)
                
                restored_count += 1

            except Exception as e:
                if hasattr(e, 'winerror') and e.winerror == 5:
                    permission_denied_titles.append(layout['title_at_save'])
                else:
                    print(f"恢復 '{layout['title_at_save']}' 位置時出錯: {e}")
        
        self.root.after(0, self._finalize_restore, restored_count, num_closed_windows, permission_denied_titles, len(current_hwnds))

    def auto_update_layout(self):
        newly_added_titles = []
        all_windows = self._get_all_windows()
        
        for hwnd in all_windows:
            if hwnd not in self.saved_layout:
                title = win32gui.GetWindowText(hwnd)
                rect = win32gui.GetWindowRect(hwnd)
                left, top, right, bottom = rect
                
                # 獲取執行檔路徑和名稱
                exe_path = ""
                exe_name = ""
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                    exe_path = win32process.GetModuleFileNameEx(handle, 0)
                    exe_name = os.path.basename(exe_path)
                    win32api.CloseHandle(handle)
                except Exception:
                    pass
                
                self.saved_layout[hwnd] = {
                    "title_at_save": title,
                    "left": left, "top": top,
                    "width": right - left, "height": bottom - top,
                    "exe_path": exe_path,
                    "exe_name": exe_name
                }
                newly_added_titles.append(title)

        if newly_added_titles:
            status_msg = f"自動偵測到 {len(newly_added_titles)} 個新視窗。"
            self._set_status(status_msg)
            print(self.status_var.get())
            for title in newly_added_titles:
                print(f"  - {title}")
            print(f"目前共記錄 {len(self.saved_layout)} 個視窗。")

        self.root.after(60000, self.auto_update_layout)

    def open_layout_editor(self):
        """開啟佈局編輯視窗"""
        if not self.saved_layout:
            messagebox.showwarning("提醒", "目前沒有記錄任何視窗。")
            return
        
        editor_window = tk.Toplevel(self.root)
        editor_window.title("編輯視窗佈局")
        editor_window.geometry("900x500")
        editor_window.transient(self.root)  # 建立父子關係
        
        try:
            editor_window.iconbitmap(resource_path('favicon.ico'))
        except tk.TclError:
            pass
        
        # 窗口置中
        editor_window.update_idletasks()
        screen_width = editor_window.winfo_screenwidth()
        screen_height = editor_window.winfo_screenheight()
        window_width = editor_window.winfo_width()
        window_height = editor_window.winfo_height()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        editor_window.geometry(f"900x500+{center_x}+{center_y}")
        
        # 標題欄
        header_frame = tk.Frame(editor_window, bg="#f0f0f0", padx=10, pady=10)
        header_frame.pack(fill=tk.X)
        header_label = tk.Label(header_frame, text=f"共有 {len(self.saved_layout)} 個記錄的視窗 (雙擊直接編輯)", 
                               font=("Arial", 12, "bold"), bg="#f0f0f0")
        header_label.pack()
        
        # 樹狀視圖框架
        tree_frame = tk.Frame(editor_window)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 垂直滾動條
        vsb = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 水平滾動條
        hsb = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 樹狀視圖
        from tkinter import ttk
        columns = ("x", "y", "width", "height")
        tree = ttk.Treeview(tree_frame, columns=columns, height=15, yscrollcommand=vsb.set, xscrollcommand=hsb.set, selectmode="extended")
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)
        
        tree.column("#0", anchor=tk.W, width=350, minwidth=200)
        tree.column("x", anchor=tk.CENTER, width=80)
        tree.column("y", anchor=tk.CENTER, width=80)
        tree.column("width", anchor=tk.CENTER, width=80)
        tree.column("height", anchor=tk.CENTER, width=80)
        
        tree.heading("#0", text="視窗標題", anchor=tk.W)
        tree.heading("x", text="X")
        tree.heading("y", text="Y")
        tree.heading("width", text="Width")
        tree.heading("height", text="Height")
        
        # 加載數據到樹狀視圖 - 按 x 軸排序，x 相同時按 y 軸排序（由小到大）
        sorted_windows = sorted(self.saved_layout.items(), 
                               key=lambda item: (item[1]['left'], item[1]['top']))
        
        for hwnd, layout in sorted_windows:
            title = layout.get('title_at_save', '未知視窗')
            x = layout['left']
            y = layout['top']
            width = layout['width']
            height = layout['height']
            tree.insert("", "end", iid=str(hwnd), text=title,
                       values=(x, y, width, height))
        
        tree.pack(fill=tk.BOTH, expand=True)
        
        # 雙擊編輯事件
        def on_double_click(event):
            item = tree.identify('item', event.x, event.y)
            if not item:
                return
            
            hwnd = int(item)
            
            if hwnd in self.saved_layout:
                self.open_full_edit_dialog(hwnd, item, tree, editor_window)
        
        tree.bind("<Double-1>", on_double_click)
        
        # 強制關閉選中視窗的功能
        def force_close_selected():
            """強制關閉選中的視窗"""
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("提示", "請先選擇要關閉的視窗。", parent=editor_window)
                return
            
            # 收集要關閉的視窗資訊
            windows_to_close = []
            windows_not_exist = []
            
            for item in selected:
                hwnd = int(item)
                layout = self.saved_layout.get(hwnd)
                if layout:
                    title = layout.get('title_at_save', '未知視窗')
                    if win32gui.IsWindow(hwnd):
                        windows_to_close.append((hwnd, item, title))
                    else:
                        windows_not_exist.append((hwnd, item, title))
            
            # 處理已不存在的視窗
            if windows_not_exist:
                titles = "\n- ".join([t for _, _, t in windows_not_exist])
                result = messagebox.askyesno("視窗已不存在", 
                                            f"以下 {len(windows_not_exist)} 個視窗已不存在:\n\n- {titles}\n\n是否從記錄中移除?",
                                            parent=editor_window)
                if result:
                    for hwnd, item, _ in windows_not_exist:
                        if hwnd in self.saved_layout:
                            del self.saved_layout[hwnd]
                        tree.delete(item)
            
            # 處理要關閉的視窗
            if not windows_to_close:
                return
            
            # 確認對話框
            if len(windows_to_close) == 1:
                title = windows_to_close[0][2]
                msg = f"確定要強制關閉視窗:\n\n{title}\n\n此操作無法復原。"
            else:
                titles = "\n- ".join([t for _, _, t in windows_to_close])
                msg = f"確定要強制關閉以下 {len(windows_to_close)} 個視窗:\n\n- {titles}\n\n此操作無法復原。"
            
            result = messagebox.askyesno("確認關閉", msg, parent=editor_window)
            if not result:
                return
            
            # 執行關閉
            success_count = 0
            failed_list = []
            
            for hwnd, item, title in windows_to_close:
                try:
                    # 嘗試正常關閉
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    time.sleep(0.05)
                    
                    # 如果視窗還在,強制終止
                    if win32gui.IsWindow(hwnd):
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        handle = win32api.OpenProcess(1, False, pid)
                        win32api.TerminateProcess(handle, 0)
                        win32api.CloseHandle(handle)
                    
                    # 從記錄中移除
                    if hwnd in self.saved_layout:
                        del self.saved_layout[hwnd]
                    tree.delete(item)
                    success_count += 1
                    
                except Exception as e:
                    failed_list.append(f"{title}: {str(e)}")
            
            # 顯示結果
            if failed_list:
                fail_msg = "\n- ".join(failed_list)
                messagebox.showwarning("部分失敗", 
                                      f"成功關閉 {success_count} 個視窗\n失敗 {len(failed_list)} 個:\n\n- {fail_msg}",
                                      parent=editor_window)
            else:
                messagebox.showinfo("成功", f"已成功關閉 {success_count} 個視窗", parent=editor_window)
        
        # 關閉按鈕框架
        button_frame = tk.Frame(editor_window, padx=10, pady=10)
        button_frame.pack(fill=tk.X)
        
        close_btn = tk.Button(button_frame, text="關閉", command=editor_window.destroy, 
                             font=("Arial", 10), width=15)
        close_btn.pack(side=tk.RIGHT, padx=5)
        
        force_close_btn = tk.Button(button_frame, text="強制關閉選中視窗", command=force_close_selected,
                                   font=("Arial", 10), width=18,
                                   bg="#f8d7da", activebackground="#f5c6cb", relief=tk.FLAT)
        force_close_btn.pack(side=tk.RIGHT, padx=5)

    def open_full_edit_dialog(self, hwnd, hwnd_str, tree, parent_window=None):
        """編輯所有參數的對話框"""
        layout = self.saved_layout[hwnd]
        title = layout.get('title_at_save', '未知視窗')
        exe_name = layout.get('exe_name', '')
        exe_path = layout.get('exe_path', '')
        
        parent = parent_window if parent_window else self.root
        dialog = tk.Toplevel(parent)
        dialog.title("編輯視窗位置")
        dialog.resizable(False, False)
        dialog.transient(parent)  # 建立父子關係
        dialog.grab_set()  # 設定為模態對話框，父視窗無法操作
        
        try:
            dialog.iconbitmap(resource_path('favicon.ico'))
        except tk.TclError:
            pass
        
        # 標題框架 - 使用 Frame 包裹以便更好控制
        title_frame = tk.Frame(dialog, padx=10, pady=8)
        title_frame.pack(fill=tk.X)
        
        # 標題 - 不設固定高度,讓它自動調整
        title_label = tk.Label(title_frame, text=f"視窗: {title}", 
                              font=("Arial", 10, "bold"), wraplength=330, justify=tk.LEFT,
                              anchor=tk.W)
        title_label.pack(fill=tk.X)
        
        # 執行檔資訊
        if exe_name:
            exe_info_frame = tk.Frame(dialog, padx=10, pady=5, bg="#f8f9fa")
            exe_info_frame.pack(fill=tk.X)
            
            exe_name_label = tk.Label(exe_info_frame, text=f"執行檔: {exe_name}", 
                                     font=("Arial", 8), fg="#495057", bg="#f8f9fa",
                                     anchor=tk.W)
            exe_name_label.pack(fill=tk.X)
            
            if exe_path:
                exe_path_label = tk.Label(exe_info_frame, text=f"路徑: {exe_path}", 
                                         font=("Arial", 8), fg="#6c757d", bg="#f8f9fa",
                                         wraplength=330, justify=tk.LEFT, anchor=tk.W)
                exe_path_label.pack(fill=tk.X)
        
        # 狀態列
        status_var = tk.StringVar()
        status_var.set("請輸入新的座標和尺寸")
        status_label = tk.Label(dialog, textvariable=status_var, 
                               font=("Arial", 9), fg="#666666", wraplength=330)
        status_label.pack(pady=(5, 10), padx=10)
        
        # 表單框架
        form_frame = tk.Frame(dialog, padx=20, pady=5)
        form_frame.pack(fill=tk.X)
        
        # 讓表單框架的列置中
        form_frame.grid_columnconfigure(0, weight=1)
        form_frame.grid_columnconfigure(1, weight=1)
        
        # X 座標
        tk.Label(form_frame, text="X:", font=("Arial", 10)).grid(row=0, column=0, pady=8, sticky=tk.E, padx=(0, 5))
        x_entry = tk.Entry(form_frame, font=("Arial", 10), width=15, justify='center')
        x_entry.insert(0, str(layout['left']))
        x_entry.grid(row=0, column=1, pady=8, sticky=tk.W, padx=(5, 0))
        
        # Y 座標
        tk.Label(form_frame, text="Y:", font=("Arial", 10)).grid(row=1, column=0, pady=8, sticky=tk.E, padx=(0, 5))
        y_entry = tk.Entry(form_frame, font=("Arial", 10), width=15, justify='center')
        y_entry.insert(0, str(layout['top']))
        y_entry.grid(row=1, column=1, pady=8, sticky=tk.W, padx=(5, 0))
        
        # 寬度
        tk.Label(form_frame, text="寬度:", font=("Arial", 10)).grid(row=2, column=0, pady=8, sticky=tk.E, padx=(0, 5))
        width_entry = tk.Entry(form_frame, font=("Arial", 10), width=15, justify='center')
        width_entry.insert(0, str(layout['width']))
        width_entry.grid(row=2, column=1, pady=8, sticky=tk.W, padx=(5, 0))
        
        # 高度
        tk.Label(form_frame, text="高度:", font=("Arial", 10)).grid(row=3, column=0, pady=8, sticky=tk.E, padx=(0, 5))
        height_entry = tk.Entry(form_frame, font=("Arial", 10), width=15, justify='center')
        height_entry.insert(0, str(layout['height']))
        height_entry.grid(row=3, column=1, pady=8, sticky=tk.W, padx=(5, 0))
        
        # 按鈕框架
        button_frame = tk.Frame(dialog, padx=20, pady=15)
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 更新視窗以獲取實際尺寸
        dialog.update_idletasks()
        
        # 根據標題實際高度動態調整對話框大小
        title_height = title_label.winfo_reqheight()
        dialog_width = 350
        # 基礎高度 + 標題實際高度
        base_height = 270  # 狀態列 + 表單 + 按鈕的基礎高度
        # 如果有執行檔資訊，增加額外高度（考慮執行檔名稱和路徑兩行）
        exe_info_height = 70 if exe_name else 0
        dialog_height = base_height + title_height + 16 + exe_info_height  # 16 是 title_frame 的 pady
        
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        center_x = int(screen_width / 2 - dialog_width / 2)
        center_y = int(screen_height / 2 - dialog_height / 2)
        dialog.geometry(f"{dialog_width}x{dialog_height}+{center_x}+{center_y}")
        
        def save_all():
            try:
                x = int(x_entry.get())
                y = int(y_entry.get())
                width = int(width_entry.get())
                height = int(height_entry.get())
                
                if width <= 0 or height <= 0:
                    status_var.set("❌ 錯誤：寬度和高度必須大於 0")
                    status_label.config(fg="red")
                    return
                
                # 更新記錄
                layout['left'] = x
                layout['top'] = y
                layout['width'] = width
                layout['height'] = height
                
                # 更新樹狀視圖
                tree.item(hwnd_str, values=(x, y, width, height))
                
                # 立即應用到視窗
                if win32gui.IsWindow(hwnd):
                    try:
                        placement = win32gui.GetWindowPlacement(hwnd)
                        if placement[1] in (win32con.SW_SHOWMINIMIZED, win32con.SW_SHOWMAXIMIZED):
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        
                        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST,
                                            layout['left'], layout['top'],
                                            layout['width'], layout['height'], 0)
                        status_var.set("✓ 位置已成功更新並套用")
                        status_label.config(fg="green")
                    except Exception as e:
                        status_var.set(f"❌ 套用失敗：{str(e)}")
                        status_label.config(fg="red")
                else:
                    status_var.set("⚠ 視窗已不存在，但記錄已更新")
                    status_label.config(fg="orange")
            
            except ValueError:
                status_var.set("❌ 錯誤：請輸入有效的數字")
                status_label.config(fg="red")
        
        # 按鍵綁定
        dialog.bind("<Return>", lambda e: save_all())
        dialog.bind("<Escape>", lambda e: dialog.destroy())
        
        save_btn = tk.Button(button_frame, text="確定", command=save_all, 
                            font=("Arial", 10), bg="#d4edda", activebackground="#c3e6cb", relief=tk.FLAT)
        save_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        cancel_btn = tk.Button(button_frame, text="關閉", command=dialog.destroy, 
                              font=("Arial", 10))
        cancel_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

    def open_quick_edit_dialog(self, hwnd, hwnd_str, tree, col_index):
        """快速編輯對話框"""
        layout = self.saved_layout[hwnd]
        title = layout.get('title_at_save', '未知視窗')
        
        col_names = ["X", "Y", "寬度", "高度"]
        col_keys = ["left", "top", "width", "height"]
        col_current_values = [layout['left'], layout['top'], layout['width'], layout['height']]
        
        if col_index >= len(col_names):
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title(f"編輯 {col_names[col_index]}")
        dialog.geometry("300x150")
        dialog.resizable(False, False)
        dialog.transient(self.root)  # 建立父子關係
        
        try:
            dialog.iconbitmap(resource_path('favicon.ico'))
        except tk.TclError:
            pass
        
        # 窗口置中
        dialog.update_idletasks()
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        window_width = dialog.winfo_width()
        window_height = dialog.winfo_height()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        dialog.geometry(f"300x150+{center_x}+{center_y}")
        
        # 標題
        title_label = tk.Label(dialog, text=f"視窗: {title}", 
                              font=("Arial", 10, "bold"), wraplength=280)
        title_label.pack(pady=10, padx=10)
        
        # 輸入框
        tk.Label(dialog, text=f"{col_names[col_index]}:", font=("Arial", 10)).pack(pady=5)
        entry = tk.Entry(dialog, font=("Arial", 12), width=20)
        entry.insert(0, str(col_current_values[col_index]))
        entry.pack(pady=5)
        entry.select_range(0, tk.END)
        entry.focus()
        
        def save_value():
            try:
                new_value = int(entry.get())
                
                if col_index >= 2 and new_value <= 0:  # 寬度和高度必須大於 0
                    messagebox.showerror("錯誤", f"{col_names[col_index]} 必須大於 0。")
                    return
                
                # 更新記錄
                layout[col_keys[col_index]] = new_value
                
                # 更新樹狀視圖
                new_values = [layout['left'], layout['top'], layout['width'], layout['height']]
                tree.item(hwnd_str, values=new_values)
                
                # 立即應用到視窗
                if win32gui.IsWindow(hwnd):
                    try:
                        placement = win32gui.GetWindowPlacement(hwnd)
                        if placement[1] in (win32con.SW_SHOWMINIMIZED, win32con.SW_SHOWMAXIMIZED):
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        
                        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST,
                                            layout['left'], layout['top'],
                                            layout['width'], layout['height'], 0)
                    except Exception as e:
                        print(f"應用視窗位置時出錯: {e}")
                
                dialog.destroy()
            
            except ValueError:
                messagebox.showerror("錯誤", "請輸入有效的數字。")
        
        # 按鍵綁定
        entry.bind("<Return>", lambda e: save_value())
        entry.bind("<Escape>", lambda e: dialog.destroy())
        
        save_btn = tk.Button(dialog, text="確定", command=save_value, 
                            font=("Arial", 10), bg="#d4edda", activebackground="#c3e6cb", relief=tk.FLAT, width=15)
        save_btn.pack(pady=10)

    def open_edit_dialog(self, hwnd_str, tree):
        """打開編輯對話框"""
        hwnd = int(hwnd_str)
        layout = self.saved_layout[hwnd]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("編輯視窗位置")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.transient(self.root)  # 建立父子關係
        
        # 標題顯示
        title_label = tk.Label(dialog, text=f"視窗: {layout.get('title_at_save', '未知視窗')}", 
                              font=("Arial", 11, "bold"), wraplength=380)
        title_label.pack(pady=10, padx=10)
        
        # 表單框架
        form_frame = tk.Frame(dialog, padx=20, pady=10)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # X 座標
        tk.Label(form_frame, text="X 座標:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=8)
        x_entry = tk.Entry(form_frame, font=("Arial", 10), width=20)
        x_entry.insert(0, str(layout['left']))
        x_entry.grid(row=0, column=1, sticky=tk.W, padx=10)
        
        # Y 座標
        tk.Label(form_frame, text="Y 座標:", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=8)
        y_entry = tk.Entry(form_frame, font=("Arial", 10), width=20)
        y_entry.insert(0, str(layout['top']))
        y_entry.grid(row=1, column=1, sticky=tk.W, padx=10)
        
        # 寬度
        tk.Label(form_frame, text="寬度:", font=("Arial", 10)).grid(row=2, column=0, sticky=tk.W, pady=8)
        width_entry = tk.Entry(form_frame, font=("Arial", 10), width=20)
        width_entry.insert(0, str(layout['width']))
        width_entry.grid(row=2, column=1, sticky=tk.W, padx=10)
        
        # 高度
        tk.Label(form_frame, text="高度:", font=("Arial", 10)).grid(row=3, column=0, sticky=tk.W, pady=8)
        height_entry = tk.Entry(form_frame, font=("Arial", 10), width=20)
        height_entry.insert(0, str(layout['height']))
        height_entry.grid(row=3, column=1, sticky=tk.W, padx=10)
        
        # 按鈕框架
        button_frame = tk.Frame(dialog, padx=20, pady=10)
        button_frame.pack(fill=tk.X)
        
        def save_changes():
            try:
                x = int(x_entry.get())
                y = int(y_entry.get())
                width = int(width_entry.get())
                height = int(height_entry.get())
                
                if width <= 0 or height <= 0:
                    messagebox.showerror("錯誤", "寬度和高度必須大於 0。")
                    return
                
                # 更新記錄
                self.saved_layout[hwnd]['left'] = x
                self.saved_layout[hwnd]['top'] = y
                self.saved_layout[hwnd]['width'] = width
                self.saved_layout[hwnd]['height'] = height
                
                # 更新樹狀視圖
                tree.item(hwnd_str, values=(layout.get('title_at_save', '未知視窗'), x, y, width, height))
                
                messagebox.showinfo("成功", "位置已更新！")
                dialog.destroy()
            
            except ValueError:
                messagebox.showerror("錯誤", "請輸入有效的數字。")
        
        save_btn = tk.Button(button_frame, text="儲存", command=save_changes, 
                            font=("Arial", 10), bg="#d4edda", activebackground="#c3e6cb", relief=tk.FLAT, width=10)
        save_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = tk.Button(button_frame, text="取消", command=dialog.destroy, 
                              font=("Arial", 10), width=10)
        cancel_btn.pack(side=tk.LEFT, padx=5)

if __name__ == "__main__":
    root = tk.Tk()
    app = WindowLayoutManager(root)
    print("程式已啟動，將在一分鐘後開始自動偵測新視窗...")
    root.mainloop()