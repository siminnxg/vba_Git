Attribute VB_Name = "ModuleCheat2"
Option Explicit
'���� �ɼ� ������ ���� ġƮŰ ���


'######################################################################################################

'ġƮŰ2 [Cheat ����] ��ư Ŭ�� �� ����

'######################################################################################################

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
            
                strCheatKey = "RequestCreateEquipmentAllOptions " & cell.Offset(0, -8).Value & " 100 4 True ()"
            
            Else
                
                strCheatKey = "��ȸ�� TID�� �������� �ʽ��ϴ�."
                
            End If
            
        Else
        
            strCheatKey = cell.Value
            
        End If
        
        '���õ� �ɼ�, �ھ� �ɼ� ��� ���� ��� ó��
        If cell = "" And cell.Offset(0, 2) = "" Then
            
            strCheatKey = "M1.Inven.RequestCreateEquipmentRandomOption " & cell.Offset(0, -8).Value & " 100 5 0 0 0 0 0 0 0 0"
            
            GoTo ġƮŰ�Է�
            
        End If
        
        '�ھ� �ɼ� üũ�Ǿ� ���� �� ����
        If �ھ�üũ = True Then
            
            '���õ� �ھ� �ɼ��� ���� ��� true -> false �� ����
            If cell.Offset(0, 2) = "" Then
                
                strCheatKey = Replace(strCheatKey, " True", " False")
                    
            '���õ� �ھ� �ɼ� �ִ� ��� ġƮŰ�� �߰�
            Else
            
                strCheatKey = strCheatKey & cell.Offset(0, 2)
            
            End If
        
        '�ھ� �ɼ� üũ ���� �� ġƮŰ false �� ����
        Else
            
            strCheatKey = Replace(strCheatKey, " True", " False")
            
        End If
        
ġƮŰ�Է�:
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

Public Sub Cheat2Core()
    
    Dim strCheatKey As String
    Dim strCoreAdr As String '---�ھ� ����Ʈ �Է� ��ġ, ���� ���� ����
    
    Application.EnableEvents = False
    
    strCheatKey = " ("
    
    '�ھ� �ɼ� ����Ʈ���� ���� ���� ��ȸ
    For Each cell In Range("Core#").Resize(, 1).Offset(0, 4)
        
        '�Էµ� ���� �ִ� ��� ����
        If cell <> "" Then
                        
            Call CellBorder(cell.Offset(0, -4))
            
            '�Է��� ������ŭ �ݺ�
            For i = 1 To cell.Value
                
                'max, min ���¿� ���� ġƮŰ �Է�
                If Range("Core_Max") = True Then
                    
                    '�ھ� �ɼ� ��ġ�� ������ �� min <-> max �ݴ�� ����
                    If cell.Offset(0, -1) < 0 Then
                    
                        strCheatKey = strCheatKey & """" & cell.Offset(0, -3).Value & ":min"","
                        
                    Else
                    
                        strCheatKey = strCheatKey & """" & cell.Offset(0, -3).Value & ":max"","
                        
                    End If
                    
                Else
                    
                    '�ھ� �ɼ� ��ġ�� ������ �� min <-> max �ݴ�� ����
                    If cell.Offset(0, -1) < 0 Then
                    
                        strCheatKey = strCheatKey & """" & cell.Offset(0, -3).Value & ":max"","
                        
                    Else
                    
                        strCheatKey = strCheatKey & """" & cell.Offset(0, -3).Value & ":min"","
                    
                    End If
                End If
                
            Next
            
            strCoreAdr = strCoreAdr & cell.Address & "," & cell.Value & "&"
            
        Else
        
            cell.Offset(0, -4).Borders.LineStyle = xlNone
            
        End If
        
    
        
    Next
    
    strCheatKey = strCheatKey & ")"
    
    '�Է��� ������ ���� ��� ����
    If strCheatKey = " ()" Then
               
        strCheatKey = ""
        
        Cells(Range(Range("selKey").Value).Row, Range("Core").Column + 5).ClearContents
        
    End If
    
    '�ӽ� ġƮŰ �Է�
    Range(Range("selKey")).Offset(0, 11) = strCheatKey
    
    '�Էµ� ��ġ�� ���� �Է�
    Cells(Range(Range("selKey").Value).Row, Range("Core").Column + 5) = strCoreAdr
    
End Sub

'======================================================================================================
'�ھ� ����Ʈ �и�

Public Sub CoreList()

    Dim rngCoreList As Range
    
    Set rngCoreList = Range("�ھ��Ʈ").CurrentRegion.Columns(1)
        
    For Each cell In rngCoreList.Cells
        
        If InStr(1, cell.Value, "�ھ�Ÿ��A") = 0 Then
            
            If InStr(1, cell.Offset(1, 0).Value, "�ھ�Ÿ��A") > 0 Then
                
                ThisWorkbook.Names("�ھ�_����").RefersTo = Range(Range("�ھ��Ʈ"), cell.Offset(0, 3))
                
                Exit For
                
            End If
            
        End If
    
    Next
        
    ThisWorkbook.Names("�ھ�_�����ǰ").RefersTo = Range(cell.Offset(1, 3), cell.End(xlDown))
    
End Sub
