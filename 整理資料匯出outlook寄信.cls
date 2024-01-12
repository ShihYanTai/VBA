Sub dataFilter()
    Dim pic As Picture
    
    Workbooks.Open Filename:="C:\Users\7600690\Desktop\(FWP020030)晶成製造通告(Production Notices)管制表-B.xlsx", ReadOnly:=True
    Windows("(FWP020030)晶成製造通告(Production Notices)管制表-B.xlsx").Activate

    ActiveSheet.Range("$A$1:$L$825").AutoFilter Field:=12, Criteria1:="="
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("新增 Microsoft Excel 工作表.xlsx").Activate
    Sheets("圖片").Select
    Range("A1").Select
    ActiveSheet.Pictures.Paste.Select
    ActiveSheet.Shapes(1).Name = "1"
    ActiveSheet.Shapes.SelectAll
    pic = Selection.Copy
    ActiveSheet.Shapes("1").Delete


    Dim oCht As Chart

    Set oCht = Charts.Add
        oCht.Name = "1"
        oCht.Activate
        oCht.Paste
        ActiveWindow.Zoom = 200
        
    '----------------------------------------------------------------------------------------------------------------------------------------------'
    
       'Setting up the Excel variables.
   Dim olApp As Object
   Dim olMailItm As Object
   Dim iCounter As Integer
   Dim Dest As Variant
   Dim SDest As String
   Dim mRange As Range
   Dim iMRange As Integer
   Dim EmailApp As Outlook.Application
   Dim EmailItem As Outlook.MailItem

   SDest = ""
   
   For Each cell In Worksheets("信箱").Range("A1:A2")
        SDest = cell.Value & ";" & SDest
   Next cell
   
   Set EmailApp = New Outlook.Application
   Set EmailItem = EmailApp.CreateItem(olMailItem)
   EmailItem.To = SDest
   EmailItem.Subject = "新年快樂"
   EmailItem.HTMLBody = "ㄅㄅ"
   EmailItem.Send
    
End Sub
