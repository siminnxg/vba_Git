Attribute VB_Name = "ModuleCommon"
Option Explicit

'��ũ�� ���� �ӵ� ���
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

Sub �̸�����_����()

    ThisWorkbook.Names("�̸�").RefersTo = Range("����")
    
End Sub

Sub ��Ӵٿ�()

    With Range("����").Validation
        .Delete
        .Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="=sheet_list"
    End With
    
End Sub



