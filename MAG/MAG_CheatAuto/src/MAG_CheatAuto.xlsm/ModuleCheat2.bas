Attribute VB_Name = "ModuleCheat2"
Option Explicit

'###################################################
'���� �ɼ� ������ ���� ġƮŰ ���
'###################################################


'======================================================================================================
'ġƮŰ2 [Cheat ����] ��ư Ŭ�� �� ����


Public Sub Cheat2()
    
    Dim strCheatKey As String
    
    
    Call SetRange
    Call UpdateStart
           
    '���õ� KEY �� ���� �� ����
    If IsEmpty(�˻����_����) = True Then
            
        MsgBox "���õ� KEY�� �������� �ʽ��ϴ�."
        
        Call UpdateEnd
        
        Exit Sub
            
    End If
    
    ġƮŰ.ClearContents '---ġƮŰ ���� �ʱ�ȭ
    
    '���õ� KEY ������ŭ ����
    For Each cell In �˻����.Offset(0, 10)
        
        Call SetRange
        
        '�ӽ÷� ������ ġƮŰ�� ���� �� ����
        If IsEmpty(cell) = True Then
            
            If IsEmpty(cell.Offset(0, -8)) = False Then
            
                strCheatKey = "M1.Inven.RequestCreateEquipmentRandomOption " & cell.Offset(0, -8).Value & " 100 5 0 0 0 0 0 0 0 0"
            
            Else
                
                strCheatKey = "��ȸ�� TID�� �������� �ʽ��ϴ�."
                
            End If
            
        Else
        
            strCheatKey = cell.Value
            
        End If
        
        If ġƮŰ_����.Value = "" Then

            ġƮŰ_����.Value = strCheatKey

        Else

            ġƮŰ_��.Offset(1, 0).Value = strCheatKey

        End If
            
    Next
    
    '���� ������ ����Ʈ ǥ��
    Call LoadTxt
    
    Call UpdateEnd
    
End Sub


Public Sub Cheat2TID()
    
    Dim strShtname As String
    Dim rngFind As Range
    
    Call SetRange
    
    For Each cell In �˻����
    
        '������ Ÿ�� �� �������� KEY �˻� �� GroupId ����
        For i = 1 To 3
        
            strShtname = Ÿ��.ListColumns("����").DataBodyRange(i).Value '---��Ʈ ���������� ����
            
            Set rngFind = Sheets(strShtname).UsedRange.Find(cell.Value, Lookat:=xlWhole) '---���õ� ���� �˻��� ��Ʈ�� �˻�
            
            '�˻��� ������ ���� �� ����
            If Not rngFind Is Nothing Then
            
                cell.Offset(0, 2) = rngFind.Offset(0, -1).Value '---TID ����
                
                cell.Offset(0, 3) = rngFind.Offset(99, 1).Value '---100���� �׷� ID ����
                
                Exit For '---�˻� �� ��ٷ� �ݺ� ����
                
            End If
        Next
    Next

End Sub
