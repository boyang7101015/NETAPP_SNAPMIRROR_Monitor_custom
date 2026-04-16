# NETAPP SNAPMIRROR Monitor Custom

這個專案用來查詢 NetApp ONTAP SnapMirror 狀態，產生每日報表，寫入歷史 CSV，並透過電子郵件寄送 HTML 報告與壓縮附件。

目前專案包含兩支主要 PowerShell 腳本：

- `Scripts/New-SettingsJson.ps1`
  用互動方式建立 `Config/settings.json`，並將 ONTAP 與 SMTP 密碼加密後存入設定檔。
- `Scripts/Send-SnapMirrorReport.ps1`
  讀取設定、呼叫 ONTAP REST API、建立報表、寫入 CSV / Log，最後寄出 HTML 郵件與 ZIP 附件。

## 功能說明

- 查詢 ONTAP SnapMirror 最新傳輸結果
- 依 `destination_volume` 產生每日報表
- 將結果寫入每月一份的歷史 CSV
- 比對本月或前月的上一筆記錄
- 產生 HTML 郵件內容
- 將當次 `csv` 與 `log` 壓縮成 `zip` 後作為郵件附件寄出
- 於主控台顯示本次執行摘要

## 專案結構

```text
NETAPP SNAPMIRROR/
├─ Config/
│  └─ settings.json
├─ Scripts/
│  ├─ New-SettingsJson.ps1
│  └─ Send-SnapMirrorReport.ps1
├─ History/               # 執行後自動建立
├─ Logs/                  # 執行後自動建立
└─ README.md
```

## 執行需求

- Windows PowerShell 5.1 或 PowerShell 7
- 可連線至 NetApp ONTAP REST API
- 可連線至 SMTP Server
- 執行帳號需能解密由 `New-SettingsJson.ps1` 產生的加密密碼

注意：
`New-SettingsJson.ps1` 使用 Windows DPAPI 加密密碼，因此通常必須在相同機器、相同 Windows 使用者帳號下執行 `Send-SnapMirrorReport.ps1`。

## 快速開始

### 1. 建立設定檔

在專案根目錄執行：

```powershell
powershell -ExecutionPolicy Bypass -File .\Scripts\New-SettingsJson.ps1
```

完成後會產生：

```text
Config\settings.json
```

如果你要指定輸出路徑：

```powershell
powershell -ExecutionPolicy Bypass -File .\Scripts\New-SettingsJson.ps1 -OutputPath .\Config\settings.json
```

### 2. 執行報表寄送

```powershell
powershell -ExecutionPolicy Bypass -File .\Scripts\Send-SnapMirrorReport.ps1
```

如果要指定不同設定檔：

```powershell
powershell -ExecutionPolicy Bypass -File .\Scripts\Send-SnapMirrorReport.ps1 -ConfigPath .\Config\settings.json
```

## 設定檔說明

預設設定檔位置為 `Config/settings.json`。

範例：

```json
{
  "Ontap": {
    "ClusterUrl": "https://cluster2.demo.netapp.com",
    "ApiPath": "/api/private/cli/snapmirror?fields=source-path,source-vserver,source-volume,destination-path,destination-vserver,destination-volume,last-transfer-type,last-transfer-error,last-transfer-size,last-transfer-duration,last-transfer-end-timestamp",
    "Username": "admin",
    "PasswordEncrypted": "<encrypted>",
    "IgnoreCertificate": true
  },
  "History": {
    "Folder": ".\\History",
    "DedupMode": "ByTransferResult"
  },
  "Mail": {
    "SmtpServer": "smtp.demo.com",
    "Port": 587,
    "UseSsl": true,
    "UseAuthentication": true,
    "Sender": "netapp-report@demo.com",
    "SenderPasswordEncrypted": "<encrypted>",
    "To": [
      "user1@demo.com"
    ],
    "Cc": [],
    "Bcc": []
  },
  "Log": {
    "Folder": ".\\Logs"
  }
}
```

### `Ontap`

- `ClusterUrl`: ONTAP 叢集 URL
- `ApiPath`: SnapMirror 查詢 API 路徑
- `Username`: ONTAP 帳號
- `Password` 或 `PasswordEncrypted`: ONTAP 密碼，建議使用加密欄位
- `IgnoreCertificate`: 是否忽略 TLS 憑證錯誤

### `History`

- `Folder`: 歷史 CSV 存放路徑
- `DedupMode`: 歷史資料去重模式

可用值：

- `None`: 不去重，全部寫入
- `ByCollectTime`: 以蒐集時間判斷重複
- `ByTransferResult`: 以傳輸結果判斷重複，這是預設值

### `Mail`

- `SmtpServer`: SMTP 主機
- `Port`: SMTP 連接埠
- `UseSsl`: 是否啟用 SSL/TLS
- `UseAuthentication`: 是否使用 SMTP 驗證
- `Sender`: 寄件者信箱
- `SenderPassword` 或 `SenderPasswordEncrypted`: SMTP 密碼
- `To`: 收件者清單，至少要有一筆
- `Cc`: 副本清單
- `Bcc`: 密件副本清單

### `Log`

- `Folder`: Log 與 ZIP 附件輸出路徑

## 輸出內容

執行 `Send-SnapMirrorReport.ps1` 後，通常會產生以下檔案：

- `History\SnapMirror_yyyy-MM.csv`
  每月歷史報表 CSV
- `Logs\SnapMirrorReport_yyyy-MM-dd.log`
  當日執行 Log
- `Logs\SnapMirrorSummary_yyyy-MM-dd_HHmmss.txt`
  當次主控台摘要文字檔
- `Logs\SnapMirrorReport_yyyy-MM-dd_HHmmss.zip`
  郵件附件壓縮檔，內含當次 CSV 與 Log

## 郵件內容

寄出的信件包含：

- HTML 內文
  顯示每個 `destination_volume` 的當前與前一次結果比較
- ZIP 附件
  內含本次使用的 CSV 與 Log 檔案

若當次找不到可附加的 `csv` 或 `log`，系統會記錄警告並改為寄送無附件郵件。

## 執行流程

`Send-SnapMirrorReport.ps1` 的主要流程如下：

1. 讀取 `settings.json`
2. 驗證設定內容
3. 呼叫 ONTAP REST API 取得 SnapMirror 資料
4. 將資料轉成目前報表記錄
5. 匯入本月與上月 CSV 進行比較
6. 依 `DedupMode` 判斷是否寫入歷史 CSV
7. 產生 HTML 報表
8. 將 CSV 與 Log 壓縮成 ZIP
9. 寄送電子郵件
10. 輸出摘要與 Log

## 排程建議

可搭配 Windows Task Scheduler 每日固定執行，例如：

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\NETAPP SNAPMIRROR\Scripts\Send-SnapMirrorReport.ps1"
```

建議使用與建立 `settings.json` 相同的 Windows 帳號來執行排程，避免加密密碼無法解密。

### 在 Windows 工作排程器加入每日 20:40 排程

1. 開啟 `工作排程器`
2. 在右側點選 `建立工作`
3. 在 `一般` 分頁輸入名稱，例如 `NetApp SnapMirror Daily Report`
4. 選擇執行帳號
5. 勾選 `不論使用者是否登入都要執行`
6. 勾選 `使用最高權限執行`
7. 切換到 `觸發程序`
8. 點選 `新增`
9. 將 `開始工作` 設為 `依排程`
10. 將排程設為 `每日`
11. 將開始時間設為 `20:40:00`
12. 確認 `已啟用` 有勾選
13. 切換到 `動作`
14. 點選 `新增`
15. `動作` 選擇 `啟動程式`
16. `程式或指令碼` 輸入：

```text
powershell.exe
```

17. `新增引數` 輸入：

```text
-ExecutionPolicy Bypass -File "C:\Path\To\NETAPP SNAPMIRROR\Scripts\Send-SnapMirrorReport.ps1"
```

18. `起始於` 輸入專案根目錄，例如：

```text
C:\Path\To\NETAPP SNAPMIRROR
```

19. 切換到 `條件`
20. 如果這台主機是伺服器，通常可取消 `只有在電腦使用交流電源時才啟動工作`
21. 切換到 `設定`
22. 建議勾選 `如果工作失敗，則每隔多久重新啟動一次`
23. 建議勾選 `如果工作執行時間超過下列時間，則將其停止`，避免異常卡住
24. 點選 `確定`
25. 輸入該 Windows 帳號密碼完成建立

建立完成後，建議先手動執行一次工作，確認以下項目都正常：

- 可以成功讀取 `Config\settings.json`
- 可以成功連線 ONTAP
- 可以成功寄送郵件
- `History`、`Logs`、`zip` 檔案都有正常產生

### 工作排程器參數範例

如果你的專案路徑是：

```text
C:\Users\administrator.BYLAB\Documents\powershell\NETAPP SNAPMIRROR
```

那麼可以使用：

- `程式或指令碼`

```text
powershell.exe
```

- `新增引數`

```text
-ExecutionPolicy Bypass -File "C:\Users\administrator.BYLAB\Documents\powershell\NETAPP SNAPMIRROR\Scripts\Send-SnapMirrorReport.ps1"
```

- `起始於`

```text
C:\Users\administrator.BYLAB\Documents\powershell\NETAPP SNAPMIRROR
```

## 常見問題

### 1. 無法解密設定檔中的密碼

請確認：

- `settings.json` 是用目前這個 Windows 使用者建立的
- 腳本是在同一台機器上執行

### 2. SMTP 驗證失敗

請確認：

- `Mail.SmtpServer`、`Port`、`UseSsl` 設定正確
- `Sender` 與密碼正確
- 郵件伺服器允許該帳號寄信

### 3. ONTAP API 無法連線

請確認：

- `ClusterUrl` 正確
- 網路連線正常
- 帳號具備查詢 SnapMirror 權限
- 若使用自簽憑證，可將 `IgnoreCertificate` 設為 `true`

## 備註

- `History.Folder` 與 `Log.Folder` 可使用相對路徑或絕對路徑
- 相對路徑會以設定檔所在資料夾為基準進行解析
- 腳本會自動建立不存在的輸出資料夾
