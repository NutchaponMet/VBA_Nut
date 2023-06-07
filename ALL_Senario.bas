Attribute VB_Name = "ALL_Senario"
Sub Senario_1()
'
'senario1 = Array("HK1", "HM1", "HMB", "HML", "HMS")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    arrayNE1 = Array("HML", "HMS")
    ne1 = 0
    Range("a1").Select
    For i = 1 To 2
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
            , SearchFormat:=False).Activate
        ActiveCell.Rows("1:3").EntireRow.Select
        ActiveCell.Activate
        Selection.Insert Shift:=xlDown
        ActiveCell.Offset(1, 0).Range("A1").Select
        ne1 = ne1 + 1
    Next i
    Range("A1").Select
    arrayNE2 = Array("HMB", "HML", "HMS")
    ne2 = 0
    For i = 1 To 3
        Cells.Find(What:=arrayNE2(ne2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        ne2 = ne2 + 1
    Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("a2").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    arrayNE1 = Array("HML", "HMS")
    ne1 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Select
        Selection.Copy
        Sheets("northeast").Activate
        Range("a2").Select
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
        ActiveCell.Offset(-1, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        ne1 = ne1 + 1
    Next i
    Sheets("northeast").Copy
' path file change
    Call save_file
    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    arrayS = Array("HK1", "HM1")
    s1 = 0
    For i = 1 To 2
        Cells.Find(What:=arrayS(s1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        s1 = s1 + 1
    Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    arrayS2 = Array("HK1", "HM1")
    s2 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Range("A1").Select
        Cells.Find(What:=arrayS2(s2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Selection.Copy
        Sheets("sourth").Activate
        Cells.Find(What:=arrayS(s2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(i - 3, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        s2 = s2 + 1
    Next i
    Sheets("sourth").Copy
' path file change
    Call save_file
    
    
End Sub

Sub Senario_2()
'
'senario2 = Array("HM1", "HMB", "HML", "HMS")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete 'Delete Data
    Range("a1:a3").EntireRow.Delete

' Define Array
    arrayNE1 = Array("HML", "HMS")
    ne1 = 0
    Range("a1").Select
    For i = 1 To 2
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
            , SearchFormat:=False).Activate
        ActiveCell.Rows("1:3").EntireRow.Select
        ActiveCell.Activate
        Selection.Insert Shift:=xlDown
        ActiveCell.Offset(1, 0).Range("A1").Select
        ne1 = ne1 + 1
    Next i
    Range("A1").Select
    arrayNE2 = Array("HMB", "HML", "HMS")
    ne2 = 0
    For i = 1 To 3
        Cells.Find(What:=arrayNE2(ne2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        ne2 = ne2 + 1
    Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("a2").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    arrayNE1 = Array("HML", "HMS")
    ne1 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Select
        Selection.Copy
        Sheets("northeast").Activate
        Range("a2").Select
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
        ActiveCell.Offset(-1, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        ne1 = ne1 + 1
    Next i
    Sheets("northeast").Copy
' path file change
    Call savefile
    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    'Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        ', SearchFormat:=False).Activate
    'ActiveCell.Rows("1:3").EntireRow.Select
    'ActiveCell.Activate
    'Selection.Insert Shift:=xlDown
    'ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    'arrayS = Array("HK1", "HM1")
    's1 = 0
    'For i = 1 To 2
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        's1 = s1 + 1
    'Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    'arrayS2 = Array("HK", "HM")
    's2 = 0
    'For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
        's2 = s2 + 1
    'Next i
    Sheets("sourth").Copy
' path file change
    Call savefile
    
End Sub


Sub Senario_3()
'
'senario3 = Array("HK1", "HMB", "HML", "HMS")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    arrayNE1 = Array("HML", "HMS")
    ne1 = 0
    Range("a1").Select
    For i = 1 To 2
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
            , SearchFormat:=False).Activate
        ActiveCell.Rows("1:3").EntireRow.Select
        ActiveCell.Activate
        Selection.Insert Shift:=xlDown
        ActiveCell.Offset(1, 0).Range("A1").Select
        ne1 = ne1 + 1
    Next i
    Range("A1").Select
    arrayNE2 = Array("HMB", "HML", "HMS")
    ne2 = 0
    For i = 1 To 3
        Cells.Find(What:=arrayNE2(ne2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        ne2 = ne2 + 1
    Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    
    arrayNE1 = Array("HML", "HMS")
    ne1 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Select
        Selection.Copy
        Sheets("northeast").Activate
        Range("a2").Select
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
        ActiveCell.Offset(-1, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        ne1 = ne1 + 1
    Next i
    Sheets("northeast").Copy
' path file change
    Call savefile
    
    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    'Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        ', SearchFormat:=False).Activate
    'ActiveCell.Rows("1:3").EntireRow.Select
    'ActiveCell.Activate
    'Selection.Insert Shift:=xlDown
    'ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    'arrayS = Array("HK1", "HM1")
    's1 = 0
    'For i = 1 To 2
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        's1 = s1 + 1
    'Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    'arrayS2 = Array("HK", "HM")
    's2 = 0
    'For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
        's2 = s2 + 1
    'Next i
    Sheets("sourth").Copy
' path file change
    Call savefile
    
End Sub

Sub Senario_4()
'
'senario4 = array("HML", "HMS", "HM1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    Range("a1").Select
    'For i = 1 To 2
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    arrayNE2 = Array("HML", "HMS")
    ne2 = 0
    For i = 1 To 2
        Cells.Find(What:=arrayNE2(ne2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        ne2 = ne2 + 1
    Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    arrayNE1 = Array("HML", "HMS")
    ne1 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Select
        Selection.Copy
        Sheets("northeast").Activate
        Range("a2").Select
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
        ActiveCell.Offset(i - 3, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        ne1 = ne1 + 1
    Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    'Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        ', SearchFormat:=False).Activate
    'ActiveCell.Rows("1:3").EntireRow.Select
    'ActiveCell.Activate
    'Selection.Insert Shift:=xlDown
    'ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    'arrayS = Array("HK1", "HM1")
    's1 = 0
    'For i = 1 To 2
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        's1 = s1 + 1
    'Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    'arrayS2 = Array("HK", "HM")
    's2 = 0
    'For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
        's2 = s2 + 1
    'Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub
Sub Senario_5()
'
'senario5 = Array("HMS", "HMB", "HM1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    Range("a1").Select
    'For i = 1 To 2
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    arrayNE2 = Array("HMB", "HMS")
    ne2 = 0
    For i = 1 To 2
        Cells.Find(What:=arrayNE2(ne2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        ne2 = ne2 + 1
    Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    arrayNE1 = Array("HMB", "HMS")
    ne1 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Select
        Selection.Copy
        Sheets("northeast").Activate
        Range("a2").Select
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
        ActiveCell.Offset(i - 3, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        ne1 = ne1 + 1
    Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    'Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        ', SearchFormat:=False).Activate
    'ActiveCell.Rows("1:3").EntireRow.Select
    'ActiveCell.Activate
    'Selection.Insert Shift:=xlDown
    'ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    'arrayS = Array("HK1", "HM1")
    's1 = 0
    'For i = 1 To 2
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        's1 = s1 + 1
    'Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    'arrayS2 = Array("HK", "HM")
    's2 = 0
    'For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
        's2 = s2 + 1
    'Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_6()
'
'senario6 = Array("HML", "HMS", "HK1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"

' NorthEast

    Sheets("northeast").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete


' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    Range("a1").Select
    'For i = 1 To 2
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    arrayNE2 = Array("HML", "HMS")
    ne2 = 0
    For i = 1 To 2
        Cells.Find(What:=arrayNE2(ne2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        ne2 = ne2 + 1
    Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    arrayNE1 = Array("HML", "HMS")
    ne1 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Select
        Selection.Copy
        Sheets("northeast").Activate
        Range("a2").Select
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
        ActiveCell.Offset(i - 3, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        ne1 = ne1 + 1
    Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    'Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        ', SearchFormat:=False).Activate
    'ActiveCell.Rows("1:3").EntireRow.Select
    'ActiveCell.Activate
    'Selection.Insert Shift:=xlDown
    'ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    'arrayS = Array("HK1", "HM1")
    's1 = 0
    'For i = 1 To 2
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        's1 = s1 + 1
    'Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    'arrayS2 = Array("HK", "HM")
    's2 = 0
    'For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
        's2 = s2 + 1
    'Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_7()
'
'senario6 = Array("HML", "HMS", "HK1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    Range("a1").Select
    'For i = 1 To 2
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    arrayNE2 = Array("HML", "HMS")
    ne2 = 0
    For i = 1 To 2
        Cells.Find(What:=arrayNE2(ne2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        ne2 = ne2 + 1
    Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    arrayNE1 = Array("HML", "HMS")
    ne1 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Select
        Selection.Copy
        Sheets("northeast").Activate
        Range("a2").Select
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
        ActiveCell.Offset(i - 3, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        ne1 = ne1 + 1
    Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    'Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        ', SearchFormat:=False).Activate
    'ActiveCell.Rows("1:3").EntireRow.Select
    'ActiveCell.Activate
    'Selection.Insert Shift:=xlDown
    'ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    'arrayS = Array("HK1", "HM1")
    's1 = 0
    'For i = 1 To 2
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        's1 = s1 + 1
    'Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    'arrayS2 = Array("HK", "HM")
    's2 = 0
    'For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
        's2 = s2 + 1
    'Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub


Sub Senario_8()
'
'senario6 = Array("HML", "HMB", "HM1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    Range("a1").Select
    'For i = 1 To 2
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    arrayNE2 = Array("HMB", "HML")
    ne2 = 0
    For i = 1 To 2
        Cells.Find(What:=arrayNE2(ne2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        ne2 = ne2 + 1
    Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    arrayNE1 = Array("HMB", "HML")
    ne1 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Select
        Selection.Copy
        Sheets("northeast").Activate
        Range("a2").Select
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
        ActiveCell.Offset(i - 3, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        ne1 = ne1 + 1
    Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    'Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        ', SearchFormat:=False).Activate
    'ActiveCell.Rows("1:3").EntireRow.Select
    'ActiveCell.Activate
    'Selection.Insert Shift:=xlDown
    'ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    'arrayS = Array("HK1", "HM1")
    's1 = 0
    'For i = 1 To 2
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        's1 = s1 + 1
    'Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    'arrayS2 = Array("HK", "HM")
    's2 = 0
    'For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
        's2 = s2 + 1
    'Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_9()
'
'senario6 = Array("HML", "HMB", "HK1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    Range("a1").Select
    'For i = 1 To 2
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    arrayNE2 = Array("HMB", "HML")
    ne2 = 0
    For i = 1 To 2
        Cells.Find(What:=arrayNE2(ne2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        ne2 = ne2 + 1
    Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    arrayNE1 = Array("HMB", "HML")
    ne1 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Select
        Selection.Copy
        Sheets("northeast").Activate
        Range("a2").Select
        Cells.Find(What:=arrayNE1(ne1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
        ActiveCell.Offset(i - 3, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        ne1 = ne1 + 1
    Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    'Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        ', SearchFormat:=False).Activate
    'ActiveCell.Rows("1:3").EntireRow.Select
    'ActiveCell.Activate
    'Selection.Insert Shift:=xlDown
    'ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    'arrayS = Array("HK1", "HM1")
    's1 = 0
    'For i = 1 To 2
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        's1 = s1 + 1
    'Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    'arrayS2 = Array("HK", "HM")
    's2 = 0
    'For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
        's2 = s2 + 1
    'Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_10()
'
'senario6 = Array("HK1", "HM1", "HMB")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    ' Range("a1").Select
    'For i = 1 To 2
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    ' arrayNE2 = Array("HMB", "HML")
    ' ne2 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     ne2 = ne2 + 1
    ' Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    ' arrayNE1 = Array("HMB", "HML")
    ' ne1 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("a2").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    '     ne1 = ne1 + 1
    ' Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    arrayS = Array("HK1", "HM1")
    s1 = 0
    For i = 1 To 2
        Cells.Find(What:=arrayS(s1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        s1 = s1 + 1
    Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    arrayS2 = Array("HK1", "HM1")
    s2 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Range("A1").Select
        Cells.Find(What:=arrayS2(s2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Selection.Copy
        Sheets("sourth").Activate
        Cells.Find(What:=arrayS2(s2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(i - 3, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        s2 = s2 + 1
    Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_11()
'
'senario6 = Array("HK1", "HM1", "HML")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    ' Range("a1").Select
    'For i = 1 To 2
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    ' arrayNE2 = Array("HMB", "HML")
    ' ne2 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     ne2 = ne2 + 1
    ' Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    ' arrayNE1 = Array("HMB", "HML")
    ' ne1 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("a2").Select
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    '     ne1 = ne1 + 1
    ' Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    arrayS = Array("HK1", "HM1")
    s1 = 0
    For i = 1 To 2
        Cells.Find(What:=arrayS(s1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        s1 = s1 + 1
    Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    arrayS2 = Array("HK1", "HM1")
    s2 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Range("A1").Select
        Cells.Find(What:=arrayS2(s2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Selection.Copy
        Sheets("sourth").Activate
        Cells.Find(What:=arrayS2(s2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(i - 3, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        s2 = s2 + 1
    Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_12()
'
'senario6 = Array("HK1", "HM1", "HMS")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    ' Range("a1").Select
    'For i = 1 To 2
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    ' arrayNE2 = Array("HMB", "HML")
    ' ne2 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     ne2 = ne2 + 1
    ' Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    ' arrayNE1 = Array("HMB", "HML")
    ' ne1 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("a2").Select
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    '     ne1 = ne1 + 1
    ' Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A1").Select
    arrayS = Array("HK1", "HM1")
    s1 = 0
    For i = 1 To 2
        Cells.Find(What:=arrayS(s1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                        Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                        Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        s1 = s1 + 1
    Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    arrayS2 = Array("HK1", "HM1")
    s2 = 0
    For i = 1 To 2
        Sheets("TEXT_Source").Activate
        Range("A1").Select
        Cells.Find(What:=arrayS2(s2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Selection.Copy
        Sheets("sourth").Activate
        Cells.Find(What:=arrayS2(s2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(i - 3, -1).Select
        Selection.PasteSpecial
        Selection.Font.Bold = True
        s2 = s2 + 1
    Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_13()
'
'senario6 = Array("HML", "HM1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    ' Range("a1").Select
    'For i = 1 To 2
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    ' arrayNE2 = Array("HMB", "HML")
    ' ne2 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     ne2 = ne2 + 1
    ' Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    ' arrayNE1 = Array("HMB", "HML")
    ' ne1 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("a2").Select
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    'ne1 = ne1 + 1
    ' Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ' Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
    ' Range("A1").Select
    ' arrayS = Array("HK1", "HM1")
    ' s1 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     s1 = s1 + 1
    ' Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' arrayS2 = Array("HK1", "HM1")
    ' s2 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    '     s2 = s2 + 1
    ' Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_14()
'
'senario6 = Array("HMS", "HM1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    ' Range("a1").Select
    'For i = 1 To 2
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    ' arrayNE2 = Array("HMB", "HML")
    ' ne2 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     ne2 = ne2 + 1
    ' Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    ' arrayNE1 = Array("HMB", "HML")
    ' ne1 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("a2").Select
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    '"HM1"ne1 = ne1 + 1
    ' Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ' Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
    ' Range("A1").Select
    ' arrayS = Array("HK1", "HM1")
    ' s1 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     s1 = s1 + 1
    ' Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' arrayS2 = Array("HK1", "HM1")
    ' s2 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    '     s2 = s2 + 1
    ' Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_15()
'
'senario6 = Array("HMB", "HM1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    ' Range("a1").Select
    'For i = 1 To 2
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    ' arrayNE2 = Array("HMB", "HML")
    ' ne2 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     ne2 = ne2 + 1
    ' Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    ' arrayNE1 = Array("HMB", "HML")
    ' ne1 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("a2").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    ' ne1 = ne1 + 1
    ' Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ' Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
    ' Range("A1").Select
    ' arrayS = Array("HK1", "HM1")
    ' s1 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     s1 = s1 + 1
    ' Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' arrayS2 = Array("HK1", "HM1")
    ' s2 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    '     s2 = s2 + 1
    ' Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_16()
'
'senario6 = Array("HML", "HK1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    ' Range("a1").Select
    'For i = 1 To 2
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    ' arrayNE2 = Array("HMB", "HML")
    ' ne2 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     ne2 = ne2 + 1
    ' Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    ' arrayNE1 = Array("HMB", "HML")
    ' ne1 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("a2").Select
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    ' ne1 = ne1 + 1
    ' Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ' Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
    ' Range("A1").Select
    ' arrayS = Array("HK1", "HM1")
    ' s1 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     s1 = s1 + 1
    ' Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' arrayS2 = Array("HK1", "HM1")
    ' s2 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    '     s2 = s2 + 1
    ' Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_17()
'
'senario6 = Array("HMS", "HK1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    ' Range("a1").Select
    'For i = 1 To 2
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    ' arrayNE2 = Array("HMB", "HML")
    ' ne2 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     ne2 = ne2 + 1
    ' Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    ' arrayNE1 = Array("HMB", "HML")
    ' ne1 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("a2").Select
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    ' ne1 = ne1 + 1
    ' Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMS", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ' Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
    ' Range("A1").Select
    ' arrayS = Array("HK1", "HM1")
    ' s1 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     s1 = s1 + 1
    ' Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' arrayS2 = Array("HK1", "HM1")
    ' s2 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    '     s2 = s2 + 1
    ' Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub

Sub Senario_18()
'
'senario6 = Array("HMB", "HK1")
'
'
'-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Rows("1:4").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(2, 0).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Select
    Selection.PasteSpecial xlPasteAllUsingSourceTheme
    Sheets("rawData").Copy after:=Sheets("rawData")
    ActiveSheet.Name = "northeast"
    Sheets("northeast").Copy after:=Sheets("northeast")
    ActiveSheet.Name = "sourth"
' NorthEast
  
    Sheets("northeast").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("a1:a3").EntireRow.Delete

' Define Array
    'arrayNE1 = Array("HML", "HMS")
    'ne1 = 0
    ' Range("a1").Select
    'For i = 1 To 2
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
        'ne1 = ne1 + 1
    'Next i
    Range("A1").Select
    ' arrayNE2 = Array("HMB", "HML")
    ' ne2 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     ne2 = ne2 + 1
    ' Next i
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="northeast", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' Sheets("TEXT_Source").Activate
    ' Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '         xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    '         , SearchFormat:=False).Activate
    ' ActiveCell.Offset(0, -1).Select
    ' Selection.Copy
    ' Sheets("northeast").Activate
    
' Under Head Table
    Range("a2").Select
    'Cells.Find(What:="HML", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                'xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                ', SearchFormat:=False).Activate
    'ActiveCell.Offset(-2, -1).Select
    'Selection.PasteSpecial
    'Selection.Font.Bold = True
    
    ' arrayNE1 = Array("HMB", "HML")
    ' ne1 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Selection.Copy
    Sheets("northeast").Activate
    Range("a2").Select
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
            , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    ' ne1 = ne1 + 1
    ' Next i
    Sheets("northeast").Copy
' path file change
    Call savefile

    
' Sourth
 
    Sheets("sourth").Activate
    Cells.Find(What:="HMB", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ' Cells.Find(What:="HM1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    '     xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    '     , SearchFormat:=False).Activate
    ' ActiveCell.Rows("1:3").EntireRow.Select
    ' ActiveCell.Activate
    ' Selection.Insert Shift:=xlDown
    ' ActiveCell.Offset(1, 0).Range("A1").Select
    ' Range("A1").Select
    ' arrayS = Array("HK1", "HM1")
    ' s1 = 0
    ' For i = 1 To 2
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                    Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
    '     s1 = s1 + 1
    ' Next i
    
' TEXT Manipulation
    ActiveSheet.Range("A1:Q1").Merge
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="sourth", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Range("A1").PasteSpecial
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    ' arrayS2 = Array("HK1", "HM1")
    ' s2 = 0
    ' For i = 1 To 2
    Sheets("TEXT_Source").Activate
    Range("A1").Select
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.Copy
    Sheets("sourth").Activate
    Cells.Find(What:="HK1", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                , SearchFormat:=False).Activate
    ActiveCell.Offset(-2, -1).Select
    Selection.PasteSpecial
    Selection.Font.Bold = True
    '     s2 = s2 + 1
    ' Next i
    Sheets("sourth").Copy
    
' path file change
    Call savefile
End Sub
