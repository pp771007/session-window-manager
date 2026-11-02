# 視窗佈局管理員 (Session Window Manager)

螢幕關掉再打開後，Windows 常會把視窗擠到別的位置。這個工具記住你排好的視窗位置，螢幕重新連接時自動放回去。

When a monitor is turned off and back on, Windows often shuffles windows around. This tool remembers your window layout and puts every window back automatically when the display reconnects.

> Windows 內建的視窗位置記憶，有些程式會失效（例如 LibreOffice），有些會有些微偏差（例如 Chrome）。這個工具直接記錄每個視窗的座標與大小，恢復時逐一設定回去。

## 📥 下載

到 [Releases](https://github.com/pp771007/session-window-manager/releases/latest) 下載 `session_window_manager.exe`，雙擊即可執行，不需安裝 Python。

程式會要求系統管理員權限，這樣才能移動以系統管理員身分執行的視窗（例如工作管理員）。

系統需求：Windows 10 / 11

## ✨ 功能

*   **螢幕重新連接時自動恢復**：螢幕關閉再打開後，等螢幕穩定 3 秒，自動恢復到最後一次儲存的佈局。螢幕配置與儲存時不同（例如多接了投影機）時不動作。
*   **手動儲存 / 恢復**：「💾 手動儲存佈局」把目前的位置設為基準；「🔄 恢復佈局」隨時一鍵放回。適合臨時把視窗拉給別人看，看完再跳回原位。
*   **自動記錄新視窗**：每分鐘掃描一次，新開的視窗以當下位置加入紀錄，已記錄的視窗不受影響。
*   **編輯紀錄（⚙）**：列出所有已記錄的視窗，雙擊可直接修改座標與大小並立即套用，也可以強制關閉選取的視窗。
*   **排除系統程式**：不記錄桌面背景與 Windows 系統資料夾下的程式（命令提示字元、PowerShell、Windows Terminal 除外）。

## 🚀 使用方式

1.  啟動程式。啟動時會自動儲存一次目前的佈局。
2.  把視窗排好後，按「💾 手動儲存佈局」更新基準。
3.  之後螢幕關掉再打開會自動恢復；任何時候也可以按「🔄 恢復佈局」手動恢復。

## ⚠️ 注意事項

*   **佈局只存在記憶體**：關閉程式後紀錄就會消失。電腦重新開機後需要重新儲存。
*   **以「視窗」為單位記錄**：程式關掉再重開會被當成新視窗，記錄的是它新開時的位置。
*   只記錄可見、未最小化、有標題的視窗。

## 🛠️ 從原始碼執行

需要 Python 3 與 pywin32：

```bash
pip install -r requirements.txt
python session_window_manager.pyw
```

從原始碼執行時不會自動要求系統管理員權限；要移動系統管理員程式的視窗，請自行以系統管理員身分執行。

自行打包 exe：執行 `產生exe.bat`（需要 PyInstaller）。推送 `v*` 標籤時，GitHub Actions 會自動打包並發布 Release。
