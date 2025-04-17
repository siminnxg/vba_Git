Attribute VB_Name = "ModuleCommon"
Option Explicit

Public ���ϰ�� As Range

Public �˻��� As Range

Public �˻���� As Range
Public �˻����_���� As Range
Public �˻����_�� As Range

Public Ű��� As Range
Public Ű���_���� As Range
Public Ű���_�� As Range

Public ġƮŰ As Range

Public Ÿ�� As ListObject

Public i, j, k As Variant '---�ݺ��� ��� ����
Public cell As Range

Public Sub SetRange()

    With Sheets("Main")
        
        Set �˻��� = .Range("B6")
        
        'Ű��� ���� ����
        Set Ű���_���� = .Range("B9")
        If IsError(Ű���_����.Value) Then
            Set Ű���_�� = Ű���_����
        Else
            Set Ű���_�� = Ű���_����.Offset(-1, 0).End(xlDown)
        End If
        Set Ű��� = Range(Ű���_����, Ű���_��)
        
        '���ϰ�� ���� ����
        Set ���ϰ�� = .Range("B3")
        
        '�˻���� ���� ����
        Set �˻����_���� = .Range("E3")
        If �˻����_����.Value = "" Then
            Set �˻����_�� = �˻����_����
        Else
            Set �˻����_�� = �˻����_����.Offset(-1, 0).End(xlDown)
        End If
        Set �˻���� = Range(�˻����_����, �˻����_��)
        
        'ġƮŰ ���� ����
        Set ġƮŰ = .Range("K3")
        
    End With
    
    With Sheets("etc")
        
        Set Ÿ�� = .ListObjects(1)
        
    End With
    
End Sub

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
