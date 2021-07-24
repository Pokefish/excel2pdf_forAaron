# excel2pdf_forAaron
## Side project on 2021 Jun.

```
有一些內部機密已拿掉(ex: ragic api 內容、mapping對照表)
```
## 需求
整理所需資料，並以網頁作為篩選頁面，一鍵將內容下載成 Excel 並轉成 pdf。

## 製作方式

### 環境設置
1. python3.8.0
2. pip 套件
     ```
     pip install XlsxWriter ,requests ,pandas ,pywin32 ,Flask ,DateTime
     ```
3. 資料夾設置 
     
     將檔案設定於桌面的forAaron資料夾

### 流程
1. 用 flask 製作頁面，並傳入篩選值。

2. 依據 [Ragic HTTP API Integration Guide](https://www.ragic.com/intl/en/doc-api) 拿到資料

3. 整理成 dataframe 

4. 製作成 excel 

5. 用 win32com.client 套件 轉成 pdf 

6. 製作 .bat