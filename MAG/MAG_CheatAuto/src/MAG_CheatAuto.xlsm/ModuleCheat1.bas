Attribute VB_Name = "ModuleCheat1"
Option Explicit
'###################################################
'Cheat1 RequestCreateItem ġƮŰ ���� ���
'###################################################

'======================================================================================================
'ġƮŰ1 [Cheat ����] ��ư Ŭ�� �� ����


Public Sub Cheat1()
    
    '# ���� ���� #
    Dim strFileName As Variant '---��ȸ�ؾ��� ������ ���� �迭
    Dim rngFindCell As Range '---key�� ��ȸ�� �� ��ġ ����
    Dim rngRune As Range '---�� �������� ��ȸ�� �� ��ġ ����
    
    '# ���� ���� #
    Call UpdateStart
    Call SetRange
    
    '���õ� Ű�� ���� �� ����
    If �˻����_����.Value = "" Then
    
        MsgBox "���õ� Key�� �����ϴ�."
        
        GoTo ����
        
    End If
    
    '��ȸ�� ������ �迭�� ����
    With Ÿ��.ListColumns("����").DataBodyRange
    
        ReDim strFileName(1 To .Cells.Count)
        
        For i = 1 To UBound(strFileName)
        
            strFileName(i) = Ÿ��.ListColumns("����").DataBodyRange(i).Value
            
        Next
        
    End With
    
    '���õ� KEY ������ŭ �ݺ�
    For Each cell In �˻����
        
        If cell.Offset(0, 2) = "" Then
            '���� ������ŭ �ݺ�
            For i = 1 To UBound(strFileName)
                
                '�� ��ȸ
                If strFileName(i) = "RuneUIData" Then
                    
                    Set rngFindCell = Sheets(strFileName(i)).UsedRange.Find(cell.Value, Lookat:=xlWhole)
                    
                    If Not rngFindCell Is Nothing Then
                    
                        Set rngRune = Sheets("RuneData").UsedRange.Find(rngFindCell.Offset(0, -1).Value, Lookat:=xlWhole)
                        
                        If Not rngRune Is Nothing Then
                        
                            cell.Offset(0, 2).Value = rngRune.Offset(0, 1).Value '---������ TID �Է�
                            
                            cell.Offset(0, 3).Value = strFileName(i) '---������ Ÿ�Կ� ������ �Է�
                            
                            GoTo ������
                        
                        End If
                    End If
                Else
                    
                    Set rngFindCell = Sheets(strFileName(i)).UsedRange.Find(cell.Value, Lookat:=xlWhole) '---key ������ �� ���� ��ȸ
                    
                    '��ȸ�Ǿ��� �� ����
                    If Not rngFindCell Is Nothing Then
                    
                        cell.Offset(0, 2).Value = rngFindCell.Offset(0, -1).Value '---������ TID �Է�
                        
                        cell.Offset(0, 3).Value = strFileName(i) '---������ Ÿ�Կ� ������ �Է�
                        
                        GoTo ������
                        
                    End If
                End If
            Next
        End If
        
������:
    Next
    
    'ġƮŰ ����
    Call CheatCreatItem
    
    ġƮŰ_����.Offset(-1, 0).Value = "�ϰ� �Է� ��� �� [�޸��� �Է�] ��ư�� Ŭ�����ּ���." '---��ܿ� �ȳ� ���� ǥ��
    
����:

    Call UpdateEnd
    
End Sub

'======================================================================================================
'ġƮŰ ����


Public Sub CheatCreatItem()
    
    '# ���� ���� #
    
    Dim InItemType As Variant '---������ Ÿ�� ���� ����
    Dim InTemplateId As Variant '---������ TID ���� ����
    Dim InCount As Variant '---������ ���� ���� ����
    Dim InLevel As Variant '---������ ���� ���� ����
        
        
    '# ���� ���� #
    
    '���� ���� ȣ��
    Call SetRange
    
    ġƮŰ.ClearContents '---ġƮŰ ���� �ʱ�ȭ
    
    '���õ� KEY ������ŭ �ݺ�
    For i = 0 To �˻����.Cells.Count - 1
        
        With �˻����(i + 1)
            
            InTemplateId = .Offset(0, 2).Value '---TID �� ����
            
            '������ ���� ������ Ÿ�� ����
            '����, �����ǰ, ������
            If .Offset(0, 3).Value = "RangedWeaponData" Or .Offset(0, 2).Value = "AccessoryData" Or .Offset(0, 2).Value = "ReactorData" Then
            
                InItemType = 2
            
            '���
            ElseIf .Offset(0, 3).Value = "ConsumableItemData" Then
            
                InItemType = 3
            
            '��
            ElseIf .Offset(0, 3).Value = "RuneUIData" Then
            
                InItemType = 4
             
            'Ŀ���͸���¡
            ElseIf .Offset(0, 3).Value = "CustomizingItemData" Then
            
                InItemType = 7
            
            '�Ƹ��� ���� ������
            ElseIf .Offset(0, 3).Value = "TuningBoardJewelData" Then
                
                InItemType = 14
                
            End If
            
            InCount = .Offset(0, 4).Value '---������ ���� ����
            
            '���� �� 1�� �Է�
            If InCount = 0 Then
            
                InCount = 1
                
            End If
                        
            InLevel = .Offset(0, 5).Value '---������ ���� ����
            
            '���� �� ���� 100 �Է�
            If InLevel = 0 Then
            
                InLevel = 100
                
            End If
        
        End With
        
        '������ ID ���� �� �ȳ� ���� ǥ��
        If InTemplateId = 0 Then
        
            ġƮŰ_����.Offset(i, 0).Value = "��ȸ�� TID�� �������� �ʽ��ϴ�."
        
        'ġƮŰ �Է�
        Else
        
            ġƮŰ_����.Offset(i, 0).Value = "RequestCreateItem " & InItemType & " " & InTemplateId & " " & _
                                        InCount & " " & InLevel
        End If
    
    Next
    
    '���� ������ ����Ʈ ǥ��
    Call LoadTxt
    
End Sub
