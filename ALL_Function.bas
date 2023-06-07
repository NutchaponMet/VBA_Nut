Attribute VB_Name = "ALL_Function"
Function loop_Senario1()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario1 As Variant
    Dim se1 As Integer
    senario1 = Array("HK1", "HM1", "HMB", "HML", "HMS")
    se1 = 0
    For Each i In senario1
        Cells.Find(What:=senario1(se1), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se1 = se1 + 1
    Next i
    loop_Senario1 = True
Exit Function
endProc:
    loop_Senario1 = False
End Function

Function loop_Senario2()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario2 As Variant
    Dim se2 As Integer
    senario2 = Array("HM1", "HMB", "HML", "HMS")
    se2 = 0
    For i = 1 To 4
        Cells.Find(What:=senario2(se2), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se2 = se2 + 1
        Debug.Print se2
    Next i
    loop_Senario2 = True
Exit Function
endProc:
    loop_Senario2 = False
End Function

Function loop_Senario3()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario3 As Variant
    Dim se3 As Integer
    senario3 = Array("HK1", "HMB", "HML", "HMS")
    se3 = 0
    For i = 1 To 4
        Cells.Find(What:=senario3(se3), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se3 = se3 + 1
    Next i
    loop_Senario3 = True
Exit Function
endProc:
    loop_Senario3 = False
End Function

Function loop_Senario4()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario4 As Variant
    Dim se4 As Integer
    senario4 = Array("HML", "HMS", "HM1")
    se4 = 0
    For i = 1 To 3
        Cells.Find(What:=senario4(se4), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se4 = se4 + 1
    Next i
    loop_Senario4 = True
Exit Function
endProc:
    loop_Senario4 = False
End Function

Function loop_Senario5()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario5 As Variant
    Dim se5 As Integer
    senario5 = Array("HMS", "HMB", "HM1")
    se5 = 0
    For i = 1 To 3
        Cells.Find(What:=senario5(se5), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se5 = se5 + 1
    Next i
    loop_Senario5 = True
Exit Function
endProc:
    loop_Senario5 = False
End Function

Function loop_Senario6()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario6 As Variant
    Dim se6 As Integer
    senario6 = Array("HML", "HMS", "HK1")
    se6 = 0
    For i = 1 To 3
        Cells.Find(What:=senario6(se6), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se6 = se6 + 1
    Next i
    loop_Senario6 = True
Exit Function
endProc:
    loop_Senario6 = False
End Function

Function loop_Senario7()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario7 As Variant
    Dim se7 As Integer
    senario7 = Array("HMS", "HMB", "HK1")
    se7 = 0

    For i = 1 To 3
        Cells.Find(What:=senario7(se7), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se7 = se7 + 1
    Next i
    loop_Senario7 = True
Exit Function
endProc:
    loop_Senario7 = False
End Function

Function loop_Senario8()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario8 As Variant
    Dim se8 As Integer
    senario8 = Array("HML", "HMB", "HM1")
    se8 = 0
    For i = 1 To 3
        Cells.Find(What:=senario8(se8), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se8 = se8 + 1
    Next i
    loop_Senario8 = True
Exit Function
endProc:
    loop_Senario8 = False
End Function

Function loop_Senario9()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario9 As Variant
    Dim se9 As Integer
    senario9 = Array("HML", "HMB", "HK1")
    se9 = 0
    For i = 1 To 3
        Cells.Find(What:=senario9(se9), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se9 = se9 + 1
    Next i
    loop_Senario9 = True
Exit Function
endProc:
    loop_Senario9 = False
End Function

Function loop_Senario10()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario10 As Variant
    Dim se10 As Integer
    senario10 = Array("HK1", "HM1", "HMB")
    se10 = 0
    For i = 1 To 3
        Cells.Find(What:=senario10(se10), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se10 = se10 + 1
    Next i
    loop_Senario10 = True
Exit Function
endProc:
    loop_Senario10 = False
End Function

Function loop_Senario11()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario11 As Variant
    Dim se11 As Integer
    senario11 = Array("HK1", "HM1", "HML")
    se11 = 0
    For i = 1 To 3
        Cells.Find(What:=senario11(se11), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se11 = se11 + 1
    Next i
    loop_Senario11 = True
Exit Function
endProc:
    loop_Senario11 = False
End Function

Function loop_Senario12()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario12 As Variant
    Dim se12 As Integer
    senario12 = Array("HK1", "HM1", "HMS")
    se12 = 0
    For i = 1 To 3
        Cells.Find(What:=senario12(se12), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se12 = se12 + 1
    Next i
    loop_Senario12 = True
Exit Function
endProc:
    loop_Senario12 = False
End Function

Function loop_Senario13()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario13 As Variant
    Dim se13 As Integer
    senario13 = Array("HML", "HM1")
    se13 = 0
    For i = 1 To 2
        Cells.Find(What:=senario13(se13), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se13 = se13 + 1
    Next i
    loop_Senario13 = True
Exit Function
endProc:
    loop_Senario13 = False
End Function

Function loop_Senario14()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario14 As Variant
    Dim se14 As Integer
    senario14 = Array("HMS", "HM1")
    se14 = 0

    For i = 1 To 2
        Cells.Find(What:=senario14(se14), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se14 = se14 + 1
    Next i
    loop_Senario14 = True
Exit Function
endProc:
    loop_Senario14 = False
End Function

Function loop_Senario15()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario15 As Variant
    Dim se15 As Integer
    senario15 = Array("HMB", "HM1")
    se15 = 0

    For i = 1 To 2
        Cells.Find(What:=senario15(se15), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se15 = se15 + 1
    Next i
    loop_Senario15 = True
Exit Function
endProc:
    loop_Senario15 = False
End Function

Function loop_Senario16()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario16 As Variant
    Dim se16 As Integer
    senario16 = Array("HML", "HK1")
    se16 = 0

    For i = 1 To 2
        Cells.Find(What:=senario16(se16), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se16 = se16 + 1
    Next i
    loop_Senario16 = True
Exit Function
endProc:
    loop_Senario16 = False
End Function

Function loop_Senario17()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario17 As Variant
    Dim se17 As Integer
    senario17 = Array("HMS", "HK1")
    se17 = 0

    For i = 1 To 2
        Cells.Find(What:=senario17(se17), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se17 = se17 + 1
    Next i
    loop_Senario17 = True
Exit Function
endProc:
    loop_Senario17 = False
End Function

Function loop_Senario18()
On Error GoTo endProc
    Dim i As Variant
    Dim cell As Range
    Dim senario18 As Variant
    Dim se18 As Integer
    senario18 = Array("HMB", "HK1")
    se18 = 0

    For i = 1 To 2
        Cells.Find(What:=senario18(se18), after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        se18 = se18 + 1
    Next i
    loop_Senario18 = True
Exit Function
endProc:
    loop_Senario18 = False
End Function
