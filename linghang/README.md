# 領航課程徵件｜審查資料協作網站

- 網頁（GitHub Pages）：`https://jjpong.github.io/others/linghang.html`
- 資料庫：Google 試算表 + 雲端硬碟 + Apps Script（`Code.gs`）

## Apps Script 端（只需做一次；之後更新 Code.gs 時重做第 3 步）
1. 在試算表的「擴充功能 → Apps Script」貼上本資料夾的 `Code.gs`，並把第 13 行 `EDIT_PASSWORD` 改成你們的密碼。
2. 同一個專案裡新增 HTML 檔 `Index`，貼上 `Index.html`（讓舊的 Apps Script 網址也能用）。
3. 部署 → 管理部署作業 → ✏️ 編輯 → 版本選「新版本」→ 部署。存取權：「所有人」。
4. 複製結尾是 `/exec` 的網址。

## 連接網頁與資料庫
網頁已預設連接到目前的 Apps Script 資料庫（`linghang.html` 裡的 `DEFAULT_API`）。若重新建立了新的部署網址，可按黃色列的「連接資料庫」或用 `?api=` 帶入新網址。
也可以直接分享這個網址給夥伴：

```
https://jjpong.github.io/others/linghang.html?api=（你的 /exec 網址）
```

打開一次之後，瀏覽器會記住，之後只要開 `linghang.html` 即可。

## 檔案
- `../linghang.html`：GitHub Pages 上的網頁（與 `Index.html` 內容相同）
- `Index.html`：貼進 Apps Script 用的同一份網頁
- `Code.gs`：Apps Script 後端
