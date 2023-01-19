# MergeStudentPDF
從資料庫提取相關欄位資訊，產生html並使用iText將html內容轉為PDF檔案

此專案搭配iText套件產生最後成果：PDF檔案，以下為使用步驟
  
    1.使用HtmlConverter 針對產生出來的Html字串轉換成PDF

    2.使用PdfOutline 產生書籤內容

    3.使用PdfMerge 將透過 HtmlConverter 的PDF文件合併(包含書籤)

    4.相關程式碼在 Librarys/StudentPDFHandle.cs 及 Librarys/PDFHelper.cs

*此專案無針對iText做源碼修改
