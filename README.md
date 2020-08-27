# CNC Line Notify 上班班表通知Script



- **The purpose is to use the Google Apps Script scheduling service to receive work notifications the day before.**

- **The program will find the (getDate()+1 day) grid in the spreadsheet, and then find 5 names under this grid to notify. .**

-  **Remember to replace <發行權杖>,<試算表ID>,or something you want to change before execute in Google Apps Script.**
<發行權杖> can be generate in [https://notify-bot.line.me/zh_TW/](https://notify-bot.line.me/zh_TW/)
message sending uses Post Method specified in the API document [https://notify-bot.line.me/doc/en/](https://notify-bot.line.me/doc/en/)



## Spreadsheet Format
|8/30  | 8/31  | ... |
|--|--|--|
|工讀生A  | 工讀生E |...|
|工讀生B  | 工讀生F |...|
|工讀生C  | 工讀生G |...|
|工讀生D  | 工讀生H |...|
|工讀生E  | 工讀生I |...|
|.|.|.|
