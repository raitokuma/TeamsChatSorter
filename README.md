# Teams「所有聊天」— 依最後更新時間排序

一個 Microsoft Teams 個人 App（固定在左側導覽列），把你**所有**的一對一、群組、會議聊天
放在同一個清單裡，依「最後訊息時間」由新到舊排序。解決 Teams Desktop 啟用聊天區段（sections）
後，聊天被分散在各區段、無法整體依時間排序的問題。

## 功能

- 全部聊天依最後訊息時間排序，分成「今天 / 昨天 / 過去 7 天 / 過去 30 天 / 更早」
- 未讀標示（粗體 + 圓點）、搜尋（名稱、成員、訊息內容）、篩選（未讀、一對一、群組、會議）
- 點擊（或 Enter）直接在 Teams 中開啟該聊天；方向鍵可上下移動
- 每 60 秒自動更新最新 50 筆，切回分頁時也會更新；每 30 分鐘完整同步一次
- 跟隨 Teams 主題（淺色 / 深色 / 高對比）
- 純前端：用 Teams 的 Nested App Authentication（NAA）取得 Graph 權杖，**不需要後端伺服器**

## 架構

```
Teams 左側 App ──► 靜態網頁 (Vite + TypeScript)
                      │  MSAL.js NAA（Teams 代為登入）
                      ▼
               Microsoft Graph
               GET /me/chats?$expand=lastMessagePreview
                    &$orderby=lastMessagePreview/createdDateTime desc
               POST /$batch  (未命名聊天的成員，用來顯示名稱)
```

| 檔案 | 用途 |
|---|---|
| `src/main.ts` | UI、排序、篩選、自動更新、開啟聊天 |
| `src/auth.ts` | 驗證（Teams 內用 NAA，一般瀏覽器用彈出視窗） |
| `src/graph.ts` | Graph API 呼叫（分頁、429 重試、$batch） |
| `appPackage/` | Teams manifest 範本與圖示 |
| `scripts/package.mjs` | 產生 Teams 安裝套件 zip（無需額外套件） |
| `.github/workflows/deploy.yml` | 部署到 GitHub Pages |

## 設定步驟

### 1. 決定網頁要放在哪裡（必須是 HTTPS）

建議用 GitHub Pages（免費）：把這個資料夾推到 GitHub repo，例如 `TeamsChatSorter`，
網址會是 `https://<你的帳號>.github.io/TeamsChatSorter/`。
其他選擇：Azure Static Web Apps、公司內部的 HTTPS 網站伺服器（把 `npm run build` 產生的 `dist/` 放上去即可）。

### 2. 在 Microsoft Entra ID 註冊應用程式

1. 開啟 [Azure Portal › App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) › **New registration**
2. 名稱：`Teams Chat Sorter`；帳戶類型：**Accounts in this organizational directory only**
3. Redirect URI：平台選 **Single-page application (SPA)**，填入
   `brk-multihub://<你的網域>`，例如 `brk-multihub://yourname.github.io`
   （只有網域，不含路徑；這是 Teams NAA 需要的）
4. 建立後到 **Authentication** › SPA 再加一個 `https://<你的網址>/redirect.html`
   （例如 `https://yourname.github.io/TeamsChatSorter/redirect.html`，用於在一般瀏覽器測試）
5. **API permissions** › Add › Microsoft Graph › **Delegated**：`User.Read`、`Chat.Read`
   （這兩個是委派權限、使用者可自行同意；若公司限制使用者同意，請 IT 按 **Grant admin consent**）
6. 記下 **Application (client) ID** 與 **Directory (tenant) ID**

### 3. 部署網頁

**GitHub Pages：**
1. repo › Settings › Pages › Source 選 **GitHub Actions**
2. repo › Settings › Secrets and variables › Actions › **Variables** 新增
   `VITE_CLIENT_ID`、`VITE_TENANT_ID`
3. 推送到 `main`，Actions 會自動建置與部署

**自行部署：**
```bash
cp .env.example .env      # 填入 VITE_CLIENT_ID / VITE_TENANT_ID
npm install
VITE_BASE=/TeamsChatSorter/ npm run build   # 若網站在根目錄則不用設 VITE_BASE
# 上傳 dist/ 內容
```
（Windows PowerShell：`$env:VITE_BASE="/TeamsChatSorter/"; npm run build`）

### 4. 產生 Teams 安裝套件

在 `.env` 加上：
```
APP_HOST_URL=https://yourname.github.io/TeamsChatSorter/
DEVELOPER_NAME=你的名字或公司
```
然後執行：
```bash
npm run package
```
會產生 `build/TeamsChatSorter.zip`。第一次執行會自動產生 `TEAMS_APP_ID` 寫入 `.env`，之後請保持不變
（更新 App 時把 `APP_VERSION` 調高，例如 `1.0.1`）。

### 5. 安裝到 Teams 並釘選到左側

1. Teams › **Apps** › **Manage your apps** › **Upload an app** › **Upload a custom app** › 選 zip
2. 按 **Add**
3. 在左側 App 列的「所有聊天」圖示上按右鍵 › **Pin**

> 若看不到「Upload a custom app」，代表組織停用了自訂 App 上傳。請 IT 在 Teams 系統管理中心
> 允許上傳，或由 IT 以 **Upload new app** 發佈到組織，並透過 **App setup policy** 直接釘選給使用者。

## 本機開發

Teams 要求 HTTPS。可用 [Dev Tunnels](https://learn.microsoft.com/azure/developer/dev-tunnels/get-started)：
```bash
npm run dev                                   # http://localhost:5173
devtunnel host -p 5173 --allow-anonymous      # 取得 https://xxxx.devtunnels.ms
```
然後把 tunnel 網域加進 Entra 的 `brk-multihub://` redirect URI，`APP_HOST_URL` 設為 tunnel 網址，再 `npm run package` 並上傳。
直接用瀏覽器開啟網址也可以測試（會用彈出視窗登入）。

## 限制與說明

- **更新頻率**：每 60 秒輪詢（只取最新 50 筆，很省流量）。即時推播需要 Graph change notifications + 後端，可日後擴充。
- **已隱藏的聊天**（在 Teams 中「隱藏」的）不會顯示，與 Teams 行為一致。
- **未讀判斷** 依 Graph `viewpoint.lastMessageReadDateTime`，可能與 Teams 用戶端有數秒落差。
- 無主題的群組/一對一聊天會額外呼叫 `/$batch` 取得成員名稱，並快取在瀏覽器中。
- 需要支援 NAA 的 Teams 用戶端（新版 Teams Desktop / Web 皆支援）。
- 個人 Microsoft 帳戶（Teams 免費版）不支援 `/me/chats` 委派權限，僅限公司或學校帳戶。
