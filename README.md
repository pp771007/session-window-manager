# 視窗佈局管理員 (Session Window Manager)

這是一個輕量級的 Windows 桌面工具，旨在幫助使用者快速儲存和恢復視窗的佈局。如果你經常需要為工作、遊戲或特定任務排列一組固定的應用程式視窗，這個工具可以讓你一鍵完成，省去手動調整的麻煩。

This is a lightweight desktop utility for Windows designed to help users quickly save and restore their window layouts. If you often arrange a specific set of application windows for work, gaming, or other tasks, this tool allows you to do it with a single click, saving you from tedious manual adjustments.

## ✨ 主要功能 (Features)

*   **一鍵儲存 (One-Click Save):** 立即擷取所有可見、非最小化視窗的位置、大小和層疊順序 (Z-order)。
*   **一鍵恢復 (One-Click Restore):** 將已記錄的視窗恢復到它們被儲存時的狀態，包括精確的位置、大小和前後順序。
*   **自動偵測新視窗 (Automatic New Window Detection):** 程式在背景每分鐘會自動掃描，並將新開啟的視窗加入佈局記憶體中，而不會影響原有的紀錄。
*   **智慧型恢復 (Intelligent Restore):** 即使視窗被最小化或最大化，也能將其恢復到正常的視窗狀態和位置。
*   **詳細報告與狀態列 (Detailed Reports & Status Bar):** 恢復後會提供摘要報告，且主介面下方有即時狀態列，顯示最新操作與時間。
*   **簡潔介面 (Simple GUI):** 使用 Tkinter 打造，直觀易用，方便操作。

## 🚀 如何使用 (How to Use)

### 1. 環境需求 (Prerequisites)

你需要在你的 Windows 電腦上安裝 Python 和 `pywin32` 函式庫。

*   **安裝 Python:**
    從 [Python 官網](https://www.python.org/downloads/) 下載並安裝最新版本的 Python。**請務必在安裝過程中勾選 "Add Python to PATH"**。

*   **安裝 pywin32:**
    打開命令提示字元 (CMD) 或 PowerShell，然後執行以下指令：
    ```bash
    pip install pywin32
    ```

### 2. 執行程式 (Running the Application)

1.  將程式碼儲存為 `session_window_manager.pyw`。
2.  (可選) 將一個名為 `favicon.ico` 的圖示檔案放在同一個資料夾，程式會自動將其設為視窗圖示。
3.  直接雙擊 `session_window_manager.pyw` 檔案即可執行。

## 📖 功能詳解 (Functional Breakdown)

#### 程式啟動 (On Startup)

*   程式啟動後，會立即在背景**靜默地儲存一次**所有當前可見視窗的佈局。這成為你的「初始佈局」。
*   同時，它會啟動一個定時器，準備在一分鐘後開始自動偵測新視窗。

#### 「手動儲存佈局 (覆蓋)」按鈕

*   點擊此按鈕會**完全清除**記憶體中現有的所有佈局紀錄。
*   然後，它會重新掃描螢幕上所有可見的視窗，並建立一個全新的佈局「快照」，包含精確的 Z-order (層疊順序)。
*   這適用於當你已經手動調整到一個完美的全新佈局，並希望將其設為新的基準點時。

#### 「恢復佈局」按鈕

*   點擊此按鈕會讀取記憶體中的所有視窗紀錄（包括初始的、手動儲存的和後來自動新增的）。
*   它會嘗試將每一個還在開啟的視窗恢復到它被記錄時的**位置和大小**。
*   它還會根據**最近一次手動儲存**的快照，嘗試恢復視窗之間的**前後層疊順序 (Z-Order)**。
*   完成後，狀態列會更新結果。若有權限問題，會彈出一個詳細的報告視窗。

#### 自動偵測 (Automatic Detection)

*   程式啟動一分鐘後，會開始第一次自動檢查。此後每分鐘檢查一次。
*   它只會尋找**新出現的、尚未被記錄的**視窗。
*   當找到新視窗時，它會將其位置和大小加入到記憶體中，而**不會**影響或覆蓋任何舊的紀錄，也不會改變已儲存的 Z-order。
*   偵測到新視窗的日誌會顯示在執行程式的**終端機視窗**中（如果透過終端機啟動），並更新主介面的狀態列。

## ⚠️ 注意事項 (Important Notes)

*   **佈局儲存於記憶體:** 所有佈局資訊都只存在於程式的執行期間。**關閉程式後，所有紀錄都會遺失。**
*   **系統管理員權限:** 如果你需要移動以「系統管理員身分」執行的程式視窗（例如：工作管理員），你需要同樣以「系統管理員身分」來執行此腳本。否則會出現「權限不足」的錯誤。
*   **Z-Order 恢復限制:** Z-Order 的恢復完全基於「手動儲存」時的快照。後來自動新增的視窗在恢復時不會參與精確的 Z-Order 排序，以避免打亂原始的層疊關係。
