Attribute VB_Name = "ModuleCommon"
Option Explicit

'매크로 실행 속도 향상
Public Sub UpdateStart()
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End Sub

Public Sub UpdateEnd()

    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

Sub 이름영역_수정()

    ThisWorkbook.Names("이름").RefersTo = Range("지정")
    
End Sub

Sub 드롭다운()

    With Range("범위").Validation
        .Delete
        .Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="=sheet_list"
    End With
    
End Sub



