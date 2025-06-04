Attribute VB_Name = "ModuleMain"
Option Explicit

'======================================================================================================
'Search �������� [KEY ����] ��ư Ŭ�� �� ����


Public Sub SelectKey()
    
    Call SetRange
    
    i = 2
    j = ""
    
    If �˻���.Offset(0, 1) = "Ŀ���͸���¡" Then
        
        i = 3
        j = "CustomizingItemData"
    
    End If
    
    'Ű ����� ��ȸ�ϸ� ���õ� �� Ȯ��
    For Each cell In Ű���
    
        '���� ���� �׵θ��� ����� ������ �� ����
        If cell.Borders.LineStyle = xlContinuous Then
            
            '�˻� ��� ù ���� ������� �� ����
            If �˻����_����.Value = "" Then
            
                �˻����_����.Resize(1, i) = cell.Resize(1, i).Value '---ù �࿡ KEY �Է�
                
                �˻����_����.Offset(0, i) = j
                
            '�˻� ��� ù �࿡ ���� ������ �� ����
            Else
            
                Set �˻����_�� = �˻����_��.Offset(1, 0)
                
                �˻����_��.Resize(1, i) = cell.Resize(1, i).Value '---�� �࿡ KEY �Է�
                
                �˻����_��.Offset(0, i) = j
                
            End If
        End If
    Next
    
    If rngCheat2.Hidden = False And IsEmpty(�˻����_����) = False Then
        
        Call Cheat2TID
        
    End If
    
    Ű���.Borders.LineStyle = xlNone '---�̵� �Ϸ� �� �׵θ� �ʱ�ȭ
    
    �˻���.Select '---�˻��� ������ Ŀ�� �̵�
       
End Sub

'======================================================================================================
'ġƮŰ1 [�ʱ�ȭ] ��ư Ŭ�� �� ����

 
Public Sub ClearKEYList1()

    Call SetRange
    
    �˻����.Resize(, 6).ClearContents
    
End Sub


'======================================================================================================
'ġƮŰ2 [�ʱ�ȭ] ��ư Ŭ�� �� ����


Public Sub ClearKEYList2()

    Call SetRange
    
    �˻����.Offset(0, 1).Borders.LineStyle = xlNone
    �˻����.Offset(0, 10).Resize(, 2).Clear
    �˻����.Resize(, 4).ClearContents
    
    Range(�˻��ɼ�_����, �˻��ɼ�_����.End(xlDown)).Borders.LineStyle = xlNone
    Range(�˻��ɼ�_����, �˻��ɼ�_����.End(xlDown)).ClearContents
        
    Range("Option").Offset(0, 1).Borders.LineStyle = xlNone
End Sub

'======================================================================================================
'ġƮŰ ����Ʈ [�ʱ�ȭ] ��ư Ŭ�� �� ����


Public Sub ClearCheatList()
    
    Call SetRange
    
    ġƮŰ.ClearContents
    
    ġƮŰ_����.Offset(-1, 0).Value = "�ϰ� �Է� ��� �� [�޸��� �Է�] ��ư�� Ŭ�����ּ���."
    
End Sub

'======================================================================================================
'Search ���� [�ʱ�ȭ] ��ư Ŭ�� �� ����


Public Sub ClearSearchList()
    
    Call SetRange
    
    Ű���.Borders.LineStyle = xlNone '---�׵θ� �ʱ�ȭ
    
    �˻��� = "" '---�˻� ���� �ʱ�ȭ
    
End Sub

'======================================================================================================
'[�޸��� �Է�] ��ư Ŭ�� �� ����


Public Sub WriteCheat()
    
    '# ���� ���� #
    
    Dim path As String '---���� ��� ���� ����
    Dim strPreset As String '---�����¸� ���� ����
    Dim strContents As String '---�޸��� ���� ���� ����
    Dim lines() As String '---�޸��� ������ �ٹٲ� ������ �����Ͽ� ���� ����
    Dim check As Boolean '---�Էµ� �����¸� üũ ����
    Dim modifiedContent As String '---���� �ԷµǾ� �ִ� ���� ��� ����
    
    
    '# ���� ���� #
    
    Call SetRange
    
    'ġƮŰ ����� ��������� ����
    If IsEmpty(ġƮŰ_����) Then
    
        MsgBox "������ ġƮŰ�� �����ϴ�."
        
        Exit Sub
        
    End If
    
    Call UpdateStart
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt" '---���� ��� ����
    
    '�����¸� ������ �� ������ �Է�
    If ������ = "" Then
    
        strPreset = "<Mag_CreateItem>"
        
    Else
    
        strPreset = "<" & ������.Value & ">"
        
    End If
        
    '������ ������ ���� ��� �ű� ����
    If Dir(path, vbDirectory) = "" Then
        
        Open path For Output As #1
                        
            Print #1, strPreset
            
            'ġƮŰ ������ ���鼭 �ݺ�
            For Each cell In ġƮŰ
                
                '��ȸ�� TID~~ �� ����
                If InStr(cell.Value, "��ȸ��") = 0 Then
                
                    Print #1, cell.Value '---�ۼ��� ġƮ �޸��忡 �Է�
                    
                End If
                
            Next
            
            Print #1, vbCrLf
        Close
        
        GoTo ����
    
    End If
    
    '������ ���� ������ �� ����
    '<Mag_CreateItem> ġƮŰ�� �Էµ� �������� ����� ������ �ٸ� �����¸��� ��� �� �ٿ��ֱ�
    If strPreset = "<Mag_CreateItem>" Then
    
        '�޸��� �б�
        Open path For Input As #1
        
            strContents = Input$(LOF(1), 1)
            
        Close #1
                
        lines = Split(strContents, vbCrLf) '---�ٹٲ��� ������ �����Ͽ� �迭�� ����
        
        '�Էµ� ���� ���� ��� �� �Է����� �̵�
        If UBound(lines) = -1 Then
        
            GoTo �̾��
            
        End If
        
        'i = MsgBox("<Mag_CreateItem> �������� ����ðڽ��ϱ�?", vbYesNo) '---����� ���� ����
        
        '����ڰ� NO ���� �� ����
        If MsgBox("<Mag_CreateItem> �������� ����ðڽ��ϱ�?", vbYesNo) = 7 Then
        
            GoTo ����
            
        End If
        
        '<Mag_CreateItem> �� �����ϰ� �����¸��� �Էµ� �� ���
        For i = 0 To UBound(lines) - 1
            
            '������ �� Ȯ��
            If InStr(lines(i), "<") > 0 And lines(i) <> "<Mag_CreateItem>" Then
            
                check = True
                
            End If
            
            If check = True Then
                    
                    modifiedContent = modifiedContent & lines(i) '---�ش� ������ ���
                    
                    '������ ���� �� ����
                    If i < UBound(lines) - 1 Then
                        
                        modifiedContent = modifiedContent & vbCrLf '---�ٹٲ� �߰�
                        
                    End If
            End If
        Next
        
        '�޸��� ���� ����
        Open path For Output As #1
        
            Print #1, strPreset
            
            'ġƮŰ ������ ���鼭 �ݺ�
            For Each cell In ġƮŰ
                
                '��ȸ�� TID~~ �� ����
                If InStr(cell.Value, "��ȸ��") = 0 Then
                
                    Print #1, cell.Value '---�ۼ��� ġƮ �޸��忡 �Է�
                    
                End If
                
            Next
            
            Print #1, vbCrLf & vbCrLf & modifiedContent '---������ ����ص� ���� �Է�
            
        Close
        
        GoTo ����
        
    End If
    
    '������ ������ �� üũ
    For Each cell In Range(LoadTxt)
        
        If cell.Value = strPreset Then
        
            MsgBox strPreset & " : ������ ������ ���� �����մϴ�."
            
            Call UpdateEnd
            
            Exit Sub
            
        End If
        
    Next
    
�̾��:

    'txt���Ͽ� �̾��
    Open path For Append As #1
    
        Print #1, strPreset '---�����¸� �Է�
            
        'ġƮŰ ������ ���鼭 �ݺ�
        For Each cell In ġƮŰ
            
            '��ȸ�� TID~~ �� ����
            If InStr(cell.Value, "��ȸ��") = 0 Then
            
                Print #1, cell.Value '---�ۼ��� ġƮ �޸��忡 �Է�
                
            End If
        Next
        
        Print #1, vbCrLf
        
    Close
        
����:
    ġƮŰ_����.Offset(-1, 0).Value = "M1.CheatUsingPreset " & path & " """ & strPreset & """"
    
    Call LoadTxt
    
    Call UpdateEnd

End Sub

'======================================================================================================
'[�޸��� ����] ��ư Ŭ�� �� ����


Public Sub OpenTxt()
    
    Dim path As String
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
    
    '��ο� �޸��� ������ ������ ����
    If Dir(path, vbDirectory) = "" Then
    
        MsgBox "�޸����� �������ּ���."
        
        Exit Sub
        
    End If
    
    Shell "notepad.exe " & Chr(34) & path & Chr(34), vbNormalFocus '---�޸��� ����

End Sub

'======================================================================================================
'Cheat ���Ͽ��� �����¸��� ã�� ����Ʈ�� ���


Public Function LoadTxt()
    
    Dim path As String
    Dim strContents As String
    Dim lines() As String
    Dim strPresetList() As Variant
    
    Call SetRange
    Call UpdateStart
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
        
    ������.Offset(2, 0).Resize(1000, 1).ClearContents
    
    '�����Ǿ��ִ� ġƮŰ ������ ���� ��� ����
    If Dir(path, vbDirectory) = "" Then
    
        Exit Function
        Call UpdateEnd
        
    End If
    
    '���Ͽ��� ������ �� �б�
    Open path For Binary As #1
    
        strContents = Space$(LOF(1))
        
        Get #1, , strContents
        
    Close #1
    
    lines = Split(strContents, vbCrLf) '---�ٹٲ��� �������� �и�
    
    ReDim strPresetList(0 To 0)
    
    j = 0
    
    '�����¸� ����Ʈ ����
    For i = 0 To UBound(lines)
        
        '������ �� Ȯ�� �� ����
        If InStr(lines(i), "<") > 0 Then
        
            strPresetList(j) = lines(i)
            
            j = j + 1
            
            ReDim Preserve strPresetList(0 To j) '---�� �� ��ŭ �迭 Ȯ��
            
        End If
    Next
        
    For i = 0 To UBound(strPresetList)
    
        ������.Offset(2 + i, 0) = strPresetList(i)
        
    Next
    
    LoadTxt = ������.Offset(2, 0).Resize(i, 1).Address
    
    Set ������_�� = ������.Offset(1 + i, 0)
    
    Call UpdateEnd
    
End Function

'======================================================================================================
'[�޸��� �ʱ�ȭ] ��ư Ŭ�� �� ����


Public Sub ClearTxt()
    
    Dim path As String
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt" '---�޸��� ��� ����
    
    '������ ������ ���� �� ����
    If Dir(path, vbDirectory) = "" Then
    
        MsgBox "������ ������ �������� �ʽ��ϴ�."
        
        Exit Sub
        
    End If
    
    If MsgBox("�޸����� �ʱ�ȭ �Ͻðڽ��ϱ�?", vbYesNo) = vbYes Then
    
        'txt���� �ʱ�ȭ
        Open path For Output As #1
        Close
        
        '������ ����Ʈ ����
        Call LoadTxt
        
        ġƮŰ_����.Offset(-1, 0).Value = "�ϰ� �Է� ��� �� [�޸��� �Է�] ��ư�� Ŭ�����ּ���."
    
    End If
    
End Sub
