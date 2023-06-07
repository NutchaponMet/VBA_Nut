Attribute VB_Name = "UserFunc"
Function save_file()
    If ActiveSheet.Name = "sourth" Then
        ActiveWorkbook.SaveAs Filename:="C:\Users\NUTCHAPON.M\Desktop\SAP GUI\Sourth.xlsx"
        ActiveWorkbook.Close
    Else
        ActiveWorkbook.SaveAs Filename:="C:\Users\NUTCHAPON.M\Desktop\SAP GUI\North_East.xlsx"
        ActiveWorkbook.Close
    End If
End Function
Function run_Senario()
    If loop_Senario1 = True Then
        Call Senario_1
    ElseIf loop_Senario2 = True Then
        Call Senario_2
    ElseIf loop_Senario3 = True Then
        Call Senario_3
    ElseIf loop_Senario4 = True Then
        Call Senario_4
    ElseIf loop_Senario5 = True Then
        Call Senario_5
    ElseIf loop_Senario6 = True Then
        Call Senario_6
    ElseIf loop_Senario7 = True Then
        Call Senario_7
    ElseIf loop_Senario8 = True Then
        Call Senario_8
    ElseIf loop_Senario9 = True Then
        Call Senario_9
    ElseIf loop_Senario10 = True Then
        Call Senario_10
    ElseIf loop_Senario11 = True Then
        Call Senario_11
    ElseIf loop_Senario12 = True Then
        Call Senario_12
    ElseIf loop_Senario13 = True Then
        Call Senario_13
    ElseIf loop_Senario14 = True Then
        Call Senario_14
    ElseIf loop_Senario15 = True Then
        Call Senario_15
    ElseIf loop_Senario16 = True Then
        Call Senario_16
    ElseIf loop_Senario17 = True Then
        Call Senario_17
    ElseIf loop_Senario18 = True Then
        Call Senario_18
    End If
End Function
