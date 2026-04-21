# SharePoint 檔案下載通知   

隨著企業數位轉型，許多組織採用 Microsoft SharePoint 作為雲端檔案管理中心。然而，雲端環境的便利性也帶來了資安挑戰——管理單位難以即時監控大規模的檔案下載行為。為此，本專案開發了一套自動化系統，主動追蹤下載紀錄並透過組織架構進行發信告知，強化企業資安攔截能力。  

* 串接 Office 365 Management Activity API，抓取 FileDownloaded 操作事件。  
* 結合企業內部 SQL Server 員工資料庫，自動過濾離職員工並關聯下載者所屬之直屬主管。  

發信規則：  
* 自動發送信給下載者，建立資安意識。
* 自動根據部門別彙整下屬行為，產出 HTML 格式的數據報表發送至主管信箱，協助異常下載提醒。
* 內建職級邏輯判斷，排除特定高階主管或管理層的通報流程，減少不必要的干擾。(可依職別各別調整)
* 處理 UTC 的轉換，確保紀錄與實際發生時間吻合。


技術層面：
* Python 3.12  
* Microsoft MSAL
* SQL Server
* M365 SMTP

系統架構：
* 透過 Tenant ID / Client Secret 向 Azure AD 獲取 Access Token。
* 向 Office API 請求活動內容 URI，並分段下載詳細記錄 JSON。
* 解析 JSON 提取下載者資訊。
* 查詢資料庫獲取職稱與直屬主管信箱。
* 將紀錄按「下載者」與「主管」進行兩層 Hash Map 彙整。
