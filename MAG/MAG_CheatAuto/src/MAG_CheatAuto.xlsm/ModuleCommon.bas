Attribute VB_Name = "ModuleCommon"
Option Explicit

'###################################################
'���� ��� ����, ���
'###################################################


'Search ����
Public �˻��� As Range
Public Ű��� As Range
Public Ű���_���� As Range
Public Ű���_�� As Range

'ġƮŰ ���� ����
Public �˻���� As Range
Public �˻����_���� As Range
Public �˻����_�� As Range
Public �˻��ɼ�_���� As Range
Public �˻��ɼ�_���� As Range

'Cheat List ����
Public ġƮŰ As Range
Public ġƮŰ_���� As Range
Public ġƮŰ_�� As Range
Public ������ As Range
Public ������_�� As Range

'etc ��Ʈ ����
Public ���ϰ�� As Range
Public Ÿ�� As ListObject

Public rngCheat1 As Range '---������ ���� ġƮŰ ����
Public rngCheat2 As Range '---���� �ɼ� ������ ���� ġƮŰ ����

Public i, j, k As Variant
Public cnt As Variant
Public cell, cell2 As Range


Public Sub SetRange()

    With ThisWorkbook.Sheets("Main")
        
        'ġƮŰ1, 2 ���� ����
        Set rngCheat1 = .Range("E:E,H:J").Columns
        Set rngCheat2 = .Range("K:L,P:P,S:T,W:W").Columns
        
        Set �˻��� = .Range("B7")
        
        'Ű��� ���� ����
        Set Ű���_���� = .Range("B10")
        If IsError(Ű���_����.Value) Then
            Set Ű���_�� = Ű���_����
        Else
            Set Ű���_�� = Ű���_����.Offset(-1, 0).End(xlDown)
        End If
        Set Ű��� = Range(Ű���_����, Ű���_��)
        
        '�˻���� ���� ����
        'Cheat1 / Cheat2 ����
        If rngCheat2.Hidden = True Then
            Set �˻����_���� = .Range("E7")
        Else
            Set �˻����_���� = .Range("L7")
            Set �˻��ɼ�_���� = �˻����_����.Offset(0, 3)
            Set �˻��ɼ�_���� = .Range("T5")
        End If
        
        If �˻����_����.Value = "" Then
            Set �˻����_�� = �˻����_����
        Else
            Set �˻����_�� = �˻����_����.Offset(-1, 0).End(xlDown)
        End If
        Set �˻���� = Range(�˻����_����, �˻����_��)
        
        'ġƮŰ ���� ����
        Set ġƮŰ_���� = .Range("X7")
        If IsEmpty(ġƮŰ_����) Then
            Set ġƮŰ_�� = ġƮŰ_����
        Else
            Set ġƮŰ_�� = ġƮŰ_����.Offset(-1, 0).End(xlDown)
        End If
        Set ġƮŰ = Range(ġƮŰ_����, ġƮŰ_��)
        
        Set ������ = ġƮŰ_����.Offset(0, 2)
        Set ������_�� = ������.Offset(1, 0).End(xlDown)
    End With
    
    With ThisWorkbook.Sheets("etc")
        
        Set Ÿ�� = .ListObjects(1)
        
        Set ���ϰ�� = .Range("H2")
        
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
