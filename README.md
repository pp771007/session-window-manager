# Session Window Manager

![Python](https://img.shields.io/badge/python-3.9+-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-0078D6.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

一個輕量、僅限當前工作階段 (session-only) 的 Windows 視窗佈局管理工具。它能讓你一鍵儲存和恢復所有視窗的位置與大小，特別適合那些視窗標題會頻繁變動的應用程式。

![Screenshot of the tool](screenshot.png)
*(提示：請將此處的 `screenshot.png` 替換為你的工具實際運作的截圖)*

---

## ✨ 主要特色

*   **一鍵儲存與恢復**：簡潔的介面，只有「儲存佈局」和「恢復佈局」兩個核心功能。
*   **基於 PID 追蹤**：使用處理程序ID (PID) 來識別視窗，而非視窗標題。這意味著即使視窗標題改變（例如瀏覽器分頁切換），工具依然能準確找到並恢復它。
*   **純記憶體運作**：所有佈局資訊只儲存在記憶體中，關閉工具後不留下任何設定檔，保持系統乾淨。
*   **啟動時自動儲存**：程式一啟動，就會自動儲存當前的視窗佈局作為初始狀態。
*   **智慧型錯誤處理**：當遇到因權限不足而無法移動的視窗時，會優雅地跳過，並在最終報告中清晰列出是哪個應用程式無法被控制。
*   **置頂顯示**：工具視窗會保持在最上層，方便隨時操作。

## ⚙️ 安裝與需求

本工具僅適用於 **Windows** 作業系統，並需要 **Python 3** 環境。

1.  **複製專案**
    ```bash
    git clone https://github.com/YOUR_USERNAME/session-window-manager.git
    cd session-window-manager
    ```

2.  **安裝必要的套件**
    本工具依賴以下 Python 套件：
    ```bash
    pip install pygetwindow pywin32 psutil
    ```

## 🚀 如何使用

1.  **佈置你的工作區**：
    在你執行此工具**之前**，先將所有你想要管理的應用程式視窗移動和縮放到你偏好的位置和大小。

2.  **執行腳本**：
    打開命令提示字元 (CMD) 或 PowerShell，然後執行：
    ```bash
    python session-window-manager.py
    ```
    **⭐ 專業提示**：為了獲得最佳效果，建議以「**系統管理員身分**」執行此腳本。這能授權工具移動那些以更高權限執行的應用程式視窗（例如工作管理員、部分遊戲等），避免「存取被拒」的錯誤。

3.  **操作工具**：
    *   **自動儲存**：工具視窗出現時，它已經在背景自動幫你儲存了一次佈局。
    *   **手動儲存**：如果你在之後調整了新的佈局，可以點擊「**儲存佈局**」按鈕來覆蓋舊的記錄。
    *   **恢復佈局**：當視窗位置變亂時，點擊「**恢復佈局**」，所有被記錄的視窗將會回到你上次儲存的位置。

4.  **結束使用**：
    直接關閉工具視窗即可。所有儲存的佈局將從記憶體中清除。

## 🛠️ 技術原理

*   **視窗識別**：透過 `pygetwindow` 獲取所有視窗物件，再利用 `win32gui` 和 `win32process` API 從視窗句柄 (HWND) 取得其對應的處理程序ID (PID)。
*   **資料儲存**：佈局資料（包含PID、位置、大小及標題）被儲存在一個全域的 Python 列表中，實現純記憶體運作。
*   **權限處理**：在恢復佈局時，使用 `try...except` 捕捉 `pygetwindow.PyGetWindowException`，並特別檢查錯誤碼 `5` (Access is denied)，以識別並報告權限問題。

## ⚠️ 限制與注意事項

*   **僅限 Windows**：由於使用了 Windows 專屬的 API，此工具無法在 macOS 或 Linux 上運行。
*   **僅限當前工作階段**：這是一個**功能而非 bug**。一旦關閉程式或重新開機，所有儲存的 PID 都會失效，佈局記錄也會遺失。
*   **單一視窗對應**：目前的邏輯是為每個 PID 記錄一個視窗。如果一個應用程式（如 Chrome）用同一個 PID 開啟了多個視窗，工具只會記錄並恢復它找到的第一個。

## 📄 授權條款

本專案採用 [MIT License](LICENSE) 授權。
