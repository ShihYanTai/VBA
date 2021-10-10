VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub 出題()
Dim k, number As String
Dim i, h, j, a, c, countfalse, counttrue, score As Integer

Worksheets("工作表2").Range("A1").Value = "=RANDBETWEEN(1,4)"

    i = Worksheets("工作表2").Range("A1")

    a = 3 * (i - 1) + 2

    c = 3 * (i - 1) + 1

    d = i * 3

Worksheets("工作表2").Cells(14, 2).Value = "OP回答第" & i & "題組"

number = InputBox("請輸入工號")

For h = 1 To 10

    Worksheets("工作表1").Select

    Selection.Range("A1").Value = ""

    ActiveSheet.Pictures.Insert("Y:\ME\NEW Erik\yt\日考系統\" & i & "\" & h & ".jpg").Select

    Application.ScreenUpdating = True

k = InputBox("輸入答案", "代碼")

Selection.Delete

Worksheets("工作表2").Cells(h + 2, a).Value = k

If Worksheets("工作表2").Cells(h + 2, c).Value = Worksheets("工作表2").Cells(h + 2, a).Value Then

    Worksheets("工作表2").Cells(h + 2, d).Value = "對"

    counttrue = counttrue + 1

Else

    Worksheets("工作表2").Cells(h + 2, d).Value = "錯"

    countfalse = countfalse + 1

End If

Next

    MsgBox "錯" & countfalse & "題"

score = counttrue / 10 * 100

Worksheets("工作表2").Range("B15").Value = score

Worksheets("工作表2").Cells(16, 2).Value = countfalse

MsgBox "總分" & score

Sheets("工作表2").Visible = True

Worksheets("工作表2").Select

    Range("A1:N19").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With

Worksheets("工作表2").Copy

With ActiveWorkbook
     .SaveAs Filename:="Y:\ME\NEW Erik\yt\日考系統\" & number & "OP作答log.xlsx", FileFormat:=xlOpenXMLWorkbook
     .Close SaveChanges:=False

End With

Application.CutCopyMode = False

    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

Sheets("工作表2").Select

ActiveWindow.SelectedSheets.Visible = False

Worksheets("工作表1").Select

End Sub
