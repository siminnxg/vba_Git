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
    
    cnt = 0
    
    For Each cell In �˻����.Offset(0, 9)
        If IsEmpty(cell) = False Then
            
            If ġƮŰ_��.Value = "" Then
                ġƮŰ_��.Value = cell.Value
                
            Else
                ġƮŰ_��.Offset(1, 0).Value = cell.Value
                
            End If
            
            
        End If
    Next
End Sub
