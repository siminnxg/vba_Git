Attribute VB_Name = "ModuleCheat2"
Option Explicit

'###################################################
'���� �ɼ� ������ ���� ġƮŰ ���
'###################################################

'M1.Inven.RequestCreateEquipmentRandomOption (������TID) (����) (�� ����) (�ɼ� 1) (�ɼ� 1 ��ġ) (�ɼ� 2) (�ɼ� 2 ��ġ) (�ɼ� 3) (�ɼ� 3 ��ġ) (�ɼ� 4) (�ɼ� 4 ��ġ)
Public Sub Cheat2()
    
    Dim strCheatKey As String
    Dim strCheatTid As String
    Dim strCheatStat As String
    
    Dim cnt As Variant
    
    Call SetRange
    
    If Range("option")(1) = "" Then
        Exit Sub
    End If
    
    i = 0
    
    '���õ� kEY ����Ʈ Ȯ��
    For Each cell In �˻����
        If cell.Borders.LineStyle = xlContinuous Then
            strCheatKey = "M1.Inven.RequestCreateEquipmentRandomOption " & _
                            cell.Offset(0, 1).Value & " 100 5 "
            Exit For
        End If
    Next
    
    '���õ� �ɼ��� ���� ��� ����
    If IsNull(strCheatKey) Then
        Exit Sub
    End If
    
    For Each cell In Range("Option").Offset(0, 1)
    
        If cell.Borders.LineStyle = xlContinuous Then
            
            'TID ����
            strCheatTid = cell.Offset(0, 1).Value
            
            'MAX �� ����
            If �˻��ɼ�_���� = False Then
                strCheatStat = cell.Offset(0, 4).Value
                
            'MIN �� ����
            Else
                strCheatStat = cell.Offset(0, 3).Value
            End If
            
            strCheatKey = strCheatKey & strCheatTid & " " & strCheatStat & " "
            
            cnt = cnt + 1
            
        End If
        
    Next
    
    '�ɼ� �ִ� 4��, ���� �� 0 0 �Է�
    For i = cnt To 3
        
        strCheatKey = strCheatKey & "0 0 "
        
    Next
    
    If ġƮŰ_��.Value = "" Then
        ġƮŰ_��.Value = strCheatKey
    Else
        ġƮŰ_��.Offset(1, 0).Value = strCheatKey
    End If
    
End Sub
