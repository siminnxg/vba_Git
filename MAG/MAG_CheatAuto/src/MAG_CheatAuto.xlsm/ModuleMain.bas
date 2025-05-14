Attribute VB_Name = "ModuleMain"
Option Explicit

Public Sub SelectKey()
    
    Call SetRange
    
    'Ű ����� ��ȸ�ϸ� ���õ� �� Ȯ��
    For Each cell In Ű���
        If cell.Borders.LineStyle = xlContinuous Then
            
            '�˻� ��� ����ִ� ���
            If �˻����_����.Value = "" Then
                �˻����_����.Value = cell.Value
                
            '�˻� ��� ���� �����ϴ� ���
            Else
                Set �˻����_�� = �˻����_��.Offset(1, 0)
                �˻����_��.Value = cell.Value
            End If
        End If
    Next
    
    Ű���.Borders.LineStyle = xlNone '---�̵� �Ϸ� �� �׵θ� �ʱ�ȭ
    �˻���.Select

End Sub

Public Sub ClearKEYList1()

    Call SetRange
    
    �˻����.Resize(, 5).ClearContents
    
End Sub

Public Sub ClearKEYList2()

    Call SetRange
    
    �˻����.Borders.LineStyle = xlNone
    �˻����.Offset(0, 9).Resize(, 2).Clear
    �˻����.Resize(, 3).ClearContents
    
    Range(�˻��ɼ�_����, �˻��ɼ�_����.End(xlDown)).Borders.LineStyle = xlNone
    Range(�˻��ɼ�_����, �˻��ɼ�_����.End(xlDown)).ClearContents
        
    Range("Option").Offset(0, 1).Borders.LineStyle = xlNone
End Sub

Public Sub ClearCheatList()
    
    Call SetRange
    
    ġƮŰ.ClearContents
    ������.Offset(2, 0).Resize(100, 1).ClearContents
    
    ġƮŰ_����.Offset(-1, 0).Value = "�ϰ� �Է� ��� �� [�޸��� ����] ��ư�� Ŭ�����ּ���."
    
End Sub

'Search ���� �ʱ�ȭ
Public Sub ClearSearchList()
    
    Call SetRange
    
    Ű���.Borders.LineStyle = xlNone '---�׵θ� �ʱ�ȭ
    
    �˻��� = "" '---�˻� ���� �ʱ�ȭ
    
End Sub

Public Sub WriteCheat()
    
    Dim path As String
    Dim strPreset As String
    Dim strContents As String
    Dim lines() As String
    Dim check As Boolean
    Dim modifiedContent As String
    Dim index As Variant
    
    Call SetRange

    If IsEmpty(ġƮŰ_����) Then
        MsgBox "������ ġƮŰ�� �����ϴ�."
        Exit Sub
    End If
    
    Call UpdateStart
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
    
    '�����¸� Ȯ��
    If ������ = "" Then
        strPreset = "<Mag_CreatItem>"
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
    
    If strPreset = "<Mag_CreatItem>" Then
        Open path For Input As #1
            strContents = Input$(LOF(1), 1)
        Close #1
                
        lines = Split(strContents, vbCrLf)
        
        If UBound(lines) = -1 Then
            GoTo �̾��
        End If
        
        If lines(0) = "<Mag_CreatItem>" Then
            index = 1
        Else
            index = 0
        End If
            
        For i = index To UBound(lines) - 1
            If InStr(lines(i), "<") > 0 Then
                check = True
            End If

            If check = True Then
                    
                    modifiedContent = modifiedContent & lines(i)
                    
                    If i < UBound(lines) - 1 Then
                        
                        modifiedContent = modifiedContent & vbCrLf
                        
                    End If

            End If
        Next
        
        Open path For Output As #1
            Print #1, strPreset
            
            'ġƮŰ ������ ���鼭 �ݺ�
            For Each cell In ġƮŰ
                
                '��ȸ�� TID~~ �� ����
                If InStr(cell.Value, "��ȸ��") = 0 Then
                    Print #1, cell.Value '---�ۼ��� ġƮ �޸��忡 �Է�
                End If
                
            Next
            
            Print #1, vbCrLf & vbCrLf & modifiedContent
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
        
����:
    ġƮŰ_����.Offset(-1, 0).Value = "M1.CheatUsingPreset " & path & " """ & strPreset & """"
    
    Call LoadTxt
    
    Call UpdateEnd

End Sub

'Cheat ���� ����
Public Sub OpenTxt()
    
    Dim path As String
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"

    If Dir(path, vbDirectory) = "" Then
        MsgBox "�޸����� �������ּ���."
        Exit Sub
    End If
    
    Shell "notepad.exe " & Chr(34) & path & Chr(34), vbNormalFocus

End Sub

'Cheat ���Ͽ��� �����¸��� ã�� ����Ʈ�� ���
Public Function LoadTxt()
    
    Dim path As String
    Dim strContents As String
    Dim lines() As String
    Dim strPresetList() As Variant
    
    Call SetRange
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
        
    ������.Offset(2, 0).Resize(1000, 1).ClearContents
    
    '�����Ǿ��ִ� ġƮŰ ������ ���� ��� ����
    If Dir(path, vbDirectory) = "" Then
        Exit Function
    End If
    
    '���Ͽ��� ������ �� �б�
    Open path For Binary As #1
        strContents = Space$(LOF(1))
        Get #1, , strContents
    Close #1
    
    lines = Split(strContents, vbCrLf)
    
    ReDim strPresetList(0 To 0)
    
    j = 0
    
    '�����¸� ����Ʈ ����
    For i = 0 To UBound(lines)
        If InStr(lines(i), "<") > 0 Then
            strPresetList(j) = lines(i)
            j = j + 1
            ReDim Preserve strPresetList(0 To j)
        End If
    Next
        
    For i = 0 To UBound(strPresetList)
        ������.Offset(2 + i, 0) = strPresetList(i)
    Next
    
    LoadTxt = ������.Offset(2, 0).Resize(i, 1).Address
    
End Function
