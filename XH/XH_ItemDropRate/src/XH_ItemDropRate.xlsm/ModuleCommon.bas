Attribute VB_Name = "ModuleCommon"
Option Explicit

Public ���ϸ�� As Range
Public ����� As Range
Public ��� As Range

Public i, j, k As Integer
Public cell As Range
Public control As control

'ȭ�� ������Ʈ �����Ͽ� ���� �ӵ� ����
Public Sub UpdateStart()
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End Sub

'ȭ�� ������Ʈ ����
Public Sub UpdateEnd()

    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

Public Sub RunForm()
    
    UserForm1.Show 0
    
End Sub

Public Sub ClearMain()

    With Sheets("Main").UsedRange
        
        .ClearContents
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
                
    End With
    
     With Sheets("��޿���").UsedRange
        
        .ClearContents
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
                
    End With
    
    Sheets("Main").Range("A3").Select
    
End Sub

