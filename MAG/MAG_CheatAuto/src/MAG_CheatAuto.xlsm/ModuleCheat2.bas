Attribute VB_Name = "ModuleCheat2"
Option Explicit

'###################################################
'���� �ɼ� ������ ���� ġƮŰ ���
'###################################################


'======================================================================================================
'ġƮŰ2 [Cheat ����] ��ư Ŭ�� �� ����


Public Sub Cheat2()
    
    Call SetRange
    
    '���õ� KEY ������ŭ ����
    For Each cell In �˻����.Offset(0, 9)
        
        '�ӽ÷� ������ ġƮŰ�� ���� �� ����
        If IsEmpty(cell) = False Then
            
            If ġƮŰ_��.Value = "" Then
            
                ġƮŰ_��.Value = cell.Value
                
            Else
            
                ġƮŰ_��.Offset(1, 0).Value = cell.Value
                
            End If
        End If
    Next
    
    '���� ������ ����Ʈ ǥ��
    Call LoadTxt
    
End Sub
