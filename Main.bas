Attribute VB_Name = "Main"

Sub Main()
' Define Variable
    Dim i As Integer
    Dim ne1 As Integer
    Dim ne2 As Integer
    Dim s1 As Integer
    Dim s2 As Integer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Sheets("Main_Data").Activate
    Sheets("Main_Data").UsedRange.Select
    Cells.Clear

' Location Path_File
    Workbooks.Open "C:\Users\NUTCHAPON.M\Documents\SAP\SAP GUI\export.XLSX"

' In process
    ActiveSheet.UsedRange.Select
    Selection.Copy
    Workbooks("newcar_macro.xlsm").Activate
    Sheets("Main_Data").Range("A1").PasteSpecial xlPasteAllUsingSourceTheme
    ' ************************************** '
    Application.DisplayAlerts = False
    Workbooks("export.xlsx").Close
    Application.DisplayAlerts = True
    '**************************************'
    Sheets("Main_Data").Copy after:=Sheets("Main_Data")
    ActiveSheet.Name = "rawData"

    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("F2:F205") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    '##########################################################'
    lastRow2 = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Range("A" & lastRow2).Select
    Selection.EntireRow.Select
    Selection.Borders.LineStyle = xlNone
    '##########################################################'
    With ActiveSheet.Sort
        .SetRange Range("A1:AK205")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Range( _
        "A:A,C:C,D:D,H:H,K:K,L:L,N:N,R:R,T:T,V:V,W:W,X:X,Y:Y,Z:Z,AB:AB,AC:AC,AD:AD,AE:AE,AF:AF,AK:AK" _
        ).Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("L:L").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
'------------------------------------------------------
' Run Senario
    Call run_Senario
    Application.DisplayAlerts = False
    Sheets("northeast").Delete
    Sheets("rawData").Delete
    Sheets("sourth").Delete
    Application.DisplayAlerts = True
    Sheets("Command").Activate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Complete Process"

End Sub


