# 出差申報自動化 — 初次設定說明

## 第一步：安裝 Python 套件

```bash
cd /Users/cody/Documents/Claude/Projects/出差申報進化
pip install -r requirements.txt
```

---

## 第二步：取得 Google Calendar API 憑證

### 2-1  建立 Google Cloud 專案

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 點選左上角「選取專案」→「新增專案」
3. 專案名稱填入 `出差申報` → 建立

### 2-2  啟用 Google Calendar API

1. 左側選單 → **API 和服務** → **程式庫**
2. 搜尋 `Google Calendar API` → 點進去 → **啟用**

### 2-3  建立 OAuth 憑證

1. 左側選單 → **API 和服務** → **憑證**
2. 點選上方「**+ 建立憑證**」→「**OAuth 用戶端 ID**」
3. 應用程式類型選「**電腦版應用程式**」
4. 名稱隨意（例如：出差申報工具）→ 建立
5. 點選「**下載 JSON**」

### 2-4  放置憑證

將下載的 JSON 檔案**重新命名為 `credentials.json`**，
放到本專案資料夾：

```
出差申報進化/
  ├── credentials.json        ← 放這裡
  ├── 出差申報自動化.py
  └── ...
```

---

## 第三步：設定 OAuth 同意畫面（只需做一次）

1. 左側選單 → **API 和服務** → **OAuth 同意畫面**
2. 使用者類型選「**外部**」→ 建立
3. 填入應用程式名稱（例如：出差申報）、您的 Email → 儲存並繼續
4. 範圍頁面直接點「儲存並繼續」
5. 測試使用者：點「**+ Add Users**」，加入您自己的 Gmail 帳號 → 儲存並繼續

---

## 第四步：首次執行（授權）

```bash
python 出差申報自動化.py
```

- 首次執行會自動開啟瀏覽器
- 用您的 Google 帳號登入並允許存取
- 授權完成後，`token.pickle` 會自動儲存，**之後不需再授權**

---

## 每月使用方式

```bash
# 自動處理上個月（每月 5 號前執行）
python 出差申報自動化.py

# 指定年月（例如處理 2026 年 3 月）
python 出差申報自動化.py 2026 3
```

執行後會：
1. 從 Google 行事曆撈出包含「陳峙霖」的行程
2. 逐筆顯示，讓您確認或編輯（輸入 e 可編輯）
3. 自動產生 Excel 並開啟

---

## destinations.json 維護

每次出差新客戶，建議更新 `destinations.json`：

### 新增目的地
```json
"目的地名稱": { "city": "行政區", "mileage": 33, "inside_tainan": false }
```

### 新增公司對照
```json
"公司關鍵字": { "district": "目的地名稱", "full_name": "公司全名" }
```

`district` 的值必須對應 `destinations` 裡的某個 key。

---

## 注意事項

- **油費報告書** 的 I 欄（油資）已有公式 `=H*8`，不需手動填入
- **差旅申請單** 以下欄位仍需手動補填：
  - 高鐵費、火車費（依實際票價）
  - 計程車費、停車費（L/M 欄）
  - 旅館費（如有住宿）
  - 審核金額（主管欄）
- 多日出差（如 2 天），膳雜費會自動乘以天數（400×2 = 800）
