# receipt-agent

在本機執行的 Discord agent：指定頻道有人傳收據／發票照片，用 `claude -p` 辨識店家、日期、品項、類別、金額，記入 Excel 並回覆本月累計。輸入 `!本月` 可看自己本月各類別統計。

- 開始任何工作前先讀 docs/architecture.md
- 程式裡四段防護用「==== 防護 N … 開始/結束 ====」註解標出，預設全部開啟
- 不可 commit `.env`；DATA_DIR（照片與 Excel）也不要 commit
- 啟動：`uv run receipt-agent`
