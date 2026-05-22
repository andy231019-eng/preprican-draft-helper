# 投資人報告草稿產生器

內部作業工具，用於批次產生投資人報告 Gmail 草稿。

**功能流程：**
1. 上傳收件人 Excel（姓名、Email、語言偏好）
2. 上傳月報、季報、年報 PDF
3. 在「收件人 × 報表」矩陣中勾選每位收件人要收到哪些報告
4. 預覽草稿（To、CC、主旨、內文、附件）
5. 一鍵建立 Gmail 草稿（**不直接寄出**）

---

## 本機啟動

### 安裝依賴

```bash
pip install -r requirements.txt
```

### 建立 .env

複製範本並填入實際值：

```bash
cp .env.example .env
```

`.env` 範例：

```
APP_PASSWORD=your-strong-password
SECRET_KEY=（用 python3 -c "import secrets; print(secrets.token_hex(32))" 產生）
GMAIL_CLIENT_ID=你的 OAuth Client ID
GMAIL_CLIENT_SECRET=你的 OAuth Client Secret
GMAIL_REFRESH_TOKEN=你的 Refresh Token
GMAIL_USER=你的 Gmail 帳號
GMAIL_REDIRECT_URI=https://developers.google.com/oauthplayground
PORT=3000
```

### 啟動

```bash
python server.py
```

或用 gunicorn（與 Railway 行為一致）：

```bash
gunicorn server:app --bind 0.0.0.0:3000 --workers 1
```

開啟：[http://127.0.0.1:3000](http://127.0.0.1:3000)

---

## Railway 部署步驟

### 一、推上 GitHub

```bash
git init
git add server.py requirements.txt Procfile .gitignore .env.example README.md static/ report_centric_investor_demo_dataset.xlsx
git commit -m "Prepare Railway deployment"
git branch -M main
git remote add origin https://github.com/你的帳號/你的repo名稱.git
git push -u origin main
```

> **注意：** 絕對不要 `git add .env`，`.env` 含有 Gmail token 和密碼。

### 二、在 Railway 建立專案

1. 前往 [railway.app](https://railway.app) 並登入
2. 點 **New Project** → **Deploy from GitHub repo**
3. 授權並選擇剛推上去的 repo
4. Railway 會自動偵測 `Procfile` 並開始 build

### 三、設定 Railway Variables

在 Railway 專案頁面點 **Variables**，逐一新增：

| 變數名稱 | 說明 |
|---|---|
| `APP_PASSWORD` | 網站登入密碼（強密碼） |
| `SECRET_KEY` | Flask session 金鑰（隨機字串） |
| `GMAIL_CLIENT_ID` | Google OAuth Client ID |
| `GMAIL_CLIENT_SECRET` | Google OAuth Client Secret |
| `GMAIL_REFRESH_TOKEN` | Gmail Refresh Token |
| `GMAIL_USER` | 建立草稿的 Gmail 帳號 |
| `GMAIL_REDIRECT_URI` | `https://developers.google.com/oauthplayground` |

> `PORT` 不需要手動設定，Railway 會自動注入。

### 四、取得 Public URL

Variables 設定完成後，Railway 會自動重新部署。  
在專案頁面的 **Settings → Domains** 複製 Railway 提供的 `*.railway.app` 網址。

### 五、測試

1. 打開 Railway 提供的網址
2. 輸入 `APP_PASSWORD` 登入
3. 上傳收件人 Excel 與三份 PDF
4. 勾選收件人、產生預覽、建立草稿
5. 到 `GMAIL_USER` 的 Gmail **草稿匣**確認草稿是否存在

---

## Gmail API 設定（OAuth 2.0）

從本機改到 Railway 後，程式碼不需要修改，只需把 Gmail 相關變數放到 Railway Variables。

### 一、Google Cloud Console 設定

1. 前往 [console.cloud.google.com](https://console.cloud.google.com)
2. 建立新 Project 或使用既有 Project
3. 左側選單 → **APIs & Services** → **Library** → 搜尋 **Gmail API** → 啟用
4. 左側選單 → **APIs & Services** → **OAuth consent screen**
   - User Type 選 **Internal**（限公司帳號）或 **External**（需要審核）
   - App name 填入工具名稱
   - 如果是 External，在 **Test users** 加入要授權的 Gmail 帳號
5. 左側選單 → **APIs & Services** → **Credentials** → **Create Credentials** → **OAuth client ID**
   - Application type：**Web application**
   - Authorized redirect URIs 加入：
     ```
     https://developers.google.com/oauthplayground
     ```
   - （可選）也加入 Railway 網址的 callback：
     ```
     https://你的railway網址/oauth2callback
     ```
6. 建立後複製 **Client ID** 和 **Client Secret**

### 二、用 OAuth Playground 取得 Refresh Token

1. 前往 [OAuth 2.0 Playground](https://developers.google.com/oauthplayground/)
2. 右上角齒輪圖示 → 勾選 **Use your own OAuth credentials**
3. 填入剛才取得的 **Client ID** 和 **Client Secret**
4. 左側 Step 1：搜尋並選擇以下 scope：
   ```
   https://www.googleapis.com/auth/gmail.compose
   ```
5. 點 **Authorize APIs** → 用要建立草稿的 Gmail 帳號登入並同意授權
6. 自動進入 Step 2 → 點 **Exchange authorization code for tokens**
7. 複製回傳的 **Refresh token**

> Refresh token 只顯示一次，請立即複製並存入 Railway Variables。  
> 如果未來 token 失效（超過 6 個月未使用或帳號重新授權），重複以上步驟取得新 token。

### 三、填入 Railway Variables

| 變數 | 值 |
|---|---|
| `GMAIL_CLIENT_ID` | Google Cloud Console 的 Client ID |
| `GMAIL_CLIENT_SECRET` | Google Cloud Console 的 Client Secret |
| `GMAIL_REFRESH_TOKEN` | OAuth Playground 取得的 Refresh Token |
| `GMAIL_USER` | 建立草稿的 Gmail 帳號（`xxx@gmail.com`） |
| `GMAIL_REDIRECT_URI` | `https://developers.google.com/oauthplayground` |

### 四、測試草稿建立

1. 只上傳 1 位測試收件人（使用測試 email）
2. 上傳一個小型 PDF
3. 建立 Gmail 草稿
4. 到 `GMAIL_USER` 的 Gmail → **草稿匣** 確認：
   - To、CC 是否正確
   - 主旨是否正確
   - 附件是否存在且可開啟

---

## 收件人 Excel 格式

必要欄位（二選一）：

- `name` 或 `recipient_name`
- `email`

建議欄位：

| 欄位名 | 說明 |
|---|---|
| `recipient_id` | 唯一識別碼，空白時系統自動產生 |
| `nickname` | 暱稱，用於 Email 稱謂 |
| `cc` | 副本 Email |
| `language` | `zh` 或 `en`，決定用哪個模板 |
| `status` | `active`（預設）或 `inactive`（不寄送） |

欄位名稱大小寫不限，例如 `Email`、`EMAIL`、`email` 都可以。

---

## 草稿模板變數

| 變數 | 說明 |
|---|---|
| `{{fund_name}}` | 基金／專案名稱 |
| `{{report_period}}` | 報告期間 |
| `{{recipient_name}}` | 收件人姓名 |
| `{{nickname}}` | 暱稱（空白時 fallback 用姓名） |
| `{{email}}` | 收件人 Email |
| `{{cc}}` | 副本 |
| `{{attachment_list}}` | 附件清單（每行一個） |
| `{{report_count}}` | 附件數量 |
| `{{report_names}}` | 附件名稱（逗號分隔） |

---

## 注意事項與已知限制

### Railway 上傳檔案

Railway container 的本地儲存**不是永久性**的，每次重新部署後上傳的 PDF 和 Excel 都會消失。

**MVP 使用方式：** 每次使用前重新上傳收件人 Excel 和三份 PDF。

如果要長期正式使用，建議改接：
- [Google Cloud Storage](https://cloud.google.com/storage)
- [AWS S3](https://aws.amazon.com/s3/)
- [Supabase Storage](https://supabase.com/storage)

### Gmail 草稿匣

所有使用者按「建立草稿」，草稿都會建立在 `GMAIL_USER` 設定的同一個 Gmail 帳號的草稿匣。

### 密碼保護

目前只有 `APP_PASSWORD` 單一密碼保護。正式版建議改用：
- Google OAuth 登入（限定公司 Google 帳號）
- 或公司 SSO

### Google OAuth Testing 模式

如果 OAuth consent screen 還在 **Testing** 狀態：
- 只有加入 **Test users** 的帳號才能授權
- Token 有效期 7 天
- 若要給多人使用各自帳號，需要完成 Google 的 App 審核（Production 模式）

**MVP 建議：** 用單一公司 Gmail 帳號申請 OAuth、建立草稿，所有人透過這個帳號的草稿匣寄出。

### 敏感資料注意

- `.env` 絕對不要推上 GitHub
- 真實投資人名單 Excel 不要放進 git repo
- Gmail refresh token 只放在 Railway Variables，不放在任何程式碼裡
- 定期輪換 `APP_PASSWORD` 和 `SECRET_KEY`
