VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "工作表1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub readpsn()

'讀取時間

Dim Today, dayname, yesterdaynamemonth, yesterdaynamedate, yesterday As Integer

    Today = Now

    
    dayname = Format(Today, "yyyy" & "mm" & "dd")
  
'時間選擇判斷
    
If dayname = Format(Today, "dd") <> 1 Then

        yesterdaynamemonth = Format(Today, "yyyy" & "mm" & "dd") - 1

    ElseIf yesterdaynamemonth = "01" Then

        yesterdaynamemonth = Format(Today, "yyyy") - 1 & 12 & 31

    ElseIf yesterdaynamemonth = "03" Or "05" Or "07" Or "10" Or "12" Then
    
        yesterdaynamemonth = Format(Today, "yyyy") & "0" & Format(Today, "mm") - 1 & 30
    
    ElseIf yesterdaynamemonth = "04" Or "06" Or "08" Or "09" Or "11" Then
        
        yesterdaynamemonth = Format(Today, "yyyy") & "0" & Format(Today, "mm") - 1 & 31
    
End If


'執行讀取及整理

    ChDir "C:\Users\K05217\Downloads"
    
    Workbooks.Open Filename:="C:\Users\K05217\Downloads\Bumping PSN daily report_" & dayname & "_(Security C).xls"
        
'新增圖表分頁
    
    Sheets.Add After:=Sheets(Sheets.Count)
    
    Sheets("工作表1").Select
    
    Sheets("工作表1").Name = "圖片"

Sheets("PI table").Select
    
    Selection.EntireRow("20:20").Hidden = False
    
    Selection.EntireRow("1:18").Hidden = True
    
    Selection.EntireRow("21:81").Hidden = True
    
'讀取goal值

Dim goal As Single

    goal = Worksheets("PI table").Range("G24").Value * 100
    

'對數字做比對並更變顏色

Dim i, j As Integer

Dim mvalue As String
    
    For i = 23 To 24
    
        For Each cell1 In Worksheets("PI table").Range("G" & i, "G" & i)
        
            For Each cell In Worksheets("PI table").Range("H" & i, "N" & i)
        
               If cell.Value >= cell1.Value Then
                  
                With cell.Font
                .Color = -16776961
                .TintAndShade = 0
                End With
               
                mvalue = " un-meet goal"
               
               Else
               
                mvalue = ""
               
               End If
                
            Next cell
        
        Next cell1
    
    Next i
    
If mvalue = "" Then

    mvalue = " meet goal"

Else
    
    mvalue = " un-meet goal"
    
End If

    Range("D3:O44").Select
    
'轉成圖片

    Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    
    Sheets("圖片").Select
    
'選取圖複製的圖片
    
    ActiveSheet.Pictures.Paste.Select
    
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select

'第二分頁篩選

Sheets("PSN% Raw Data").Select

    Range("AB3:AD3").Select
    
    ActiveSheet.Range("$A$4:$CO$2288").AutoFilter Field:=4, Criteria1:="B6"
    
    ActiveSheet.Range("$A$4:$CO$2288").AutoFilter Field:=6, Criteria1:="PE"
    
    ActiveSheet.Range("$A$4:$CO$2288").AutoFilter Field:=7, Criteria1:="BE"
    
    ActiveWindow.SmallScroll Down:=-12
    
    ActiveSheet.Range("$A$4:$CO$2288").AutoFilter Field:=8, Criteria1:=Array( _
        "T215001_00", "T215002_00", "T216505_00", "T216508_00", "T219502_00", "T222011_00", _
        "T222012_00", "T229011_00", "T229012_00", "T239516_00", "T239518_00", "T239520_00", _
        "T240013_00", "T240015_00", "T240017_00", "T240020_00", "T240023_00", "T249001_00", _
        "T249101_00", "T505001_00", "T505007_00", "T505013_00", "T633401_00", "T507001_00", _
        "T633501_00"), Operator:= _
        xlFilterValues

'複製結果並轉成圖片

'隱藏不要的
    
    Columns("J:AA").Select
    
    Selection.EntireColumn.Hidden = True
    
    Range("D3:AD2288").Select
    
    Range("D3").Activate

' 圖表 巨集製作圖表

'    選擇PSN% Raw Data數據製作圖表

    Sheets("圖片").Select
    
    Range("A7").Select
    
    ActiveSheet.Shapes.AddChart.Select

    ActiveChart.ChartType = xlColumnClustered

    ActiveChart.SetSourceData Source:=Range("'PSN% Raw Data'!$H$435:$H$474,'PSN% Raw Data'!$AC$435:$AC$474")

    
    Dim xRg As Range
    
    Dim xChart As ChartObject
    
    Set xRg = Range("A7")
    
    Set xChart = ActiveSheet.ChartObjects(1)

    With xChart
        .Top = xRg(1).Top
        .Left = xRg(1).Left
    End With

    With Selection.Format.TextFrame2.TextRange.Font
        .Size = 10
    End With

    ActiveSheet.Shapes("圖表 2").ScaleHeight 1.8, msoFalse, ActiveWindow.ScrollColumn = 13
    
    ActiveSheet.Shapes("圖表 2").ScaleWidth 3, msoFalse, msoScaleFromTopLeft
    
    ActiveSheet.ChartObjects("圖表 2").Activate
    
    ActiveChart.ApplyLayout (8)
    
    ActiveChart.ChartTitle.Select
    
    Selection.Format.TextFrame2.TextRange.Font.Size = 40
    
    Selection.Format.TextFrame2.TextRange.Characters.Text = "PSN% Raw Data"
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 13).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 13).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 18
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    
    ActiveChart.Axes(xlValue).AxisTitle.Select
    
    Selection.Left = 8
    
    Selection.Top = 44.527
    
    ActiveWindow.Zoom = 70
    
    Selection.Format.TextFrame2.TextRange.Font.Size = 40
    
    Selection.Format.TextFrame2.TextRange.Characters.Text = "PSN Count"
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 9).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 9).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 18
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    
    Selection.Format.TextFrame2.TextRange.Font.Size = 40
    
    Selection.Format.TextFrame2.TextRange.Characters.Text = "DCOP & VALUE_NAME"
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 17).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 17).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 18
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    
'找出PSN Count 最大者




'第三分頁篩選


Sheets("PSN Raw Data").Select


    ActiveSheet.Range("$A$2:$AB$1139").AutoFilter Field:=4, Criteria1:=yesterdaynamemonth
                
    ActiveSheet.Range("$A$2:$AB$1139").AutoFilter Field:=5, Criteria1:="ASEB6"
    
    ActiveSheet.Range("$A$2:$AB$1139").AutoFilter Field:=7, Criteria1:=Array( _
        "2165", "2220", "2290", "2400", "6334", "2195", "2395", "2150", "2490", _
        "5050", "6335", "5070", "2491"), Operator:=xlFilterValues
        
        
'篩選完做轉成圖片

    Range("B2:K2").Select
    
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap

'選取圖複製的圖片

    Sheets("圖片").Select
    
    Range("A31").Select

    ActiveSheet.Pictures.Paste.Select


' 群組圖片 巨集(輸出圖片有問題)


    ActiveSheet.Shapes.Range(Array("Picture 3")).Select

    ActiveSheet.Shapes.Range(Array("Picture 1", "Chart 2", "Picture 3")).Select
    
    Selection.Copy
    
    ActiveSheet.PasteSpecial Format:="點陣圖", Link:=False, DisplayAsIcon:=False
    
    Selection.CopyPicture xlScreen, xlBitmap

    Dim oCht As Chart

    Set oCht = Charts.Add
        
        oCht.Activate

        ActiveWindow.Zoom = 200
        
    With oCht
        .Paste
        .Export Filename:="C:\Users\K05217\Downloads\圖片\" & dayname & ".jpg", Filtername:="JPG"
    End With
        
'選取數值

Dim PSNvalue As Single

Dim h As Single

Sheets("PI table").Select


'儲存整理好檔案

    ChDir "C:\Users\K05217\Downloads\整理好的"
    
    With ActiveWorkbook
    
        .SaveAs Filename:="C:\Users\K05217\Downloads\整理好的\" & dayname & " B6 PSN " & mvalue & " ( Goal " & goal & "% )" & "_(Security C).xls", FileFormat:=xlOpenXMLWorkbook, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        
        .Close SaveChanges:=False
    
    End With
    
    
End Sub



