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
    Dim fileContent As String
    Dim lines() As String
    Dim modifiedContent As String
    Dim check As Boolean
    Dim strPreset As String
    Dim strPresetList As Variant
    
    Call SetRange

    If IsEmpty(ġƮŰ_����) Then
        MsgBox "������ ġƮŰ�� �����ϴ�."
        Exit Sub
    End If
    
    Call UpdateStart
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"
    
    '�����¸� Ȯ��
    If ġƮŰ_����.Offset(0, 1) = "" Then
        strPreset = "<Mag_CreatItem>"
    Else
        strPreset = "<" & ġƮŰ_����.Offset(0, 1).Value & ">"
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
        Close
        
        Exit Sub
        
    End If
    
    '���Ͽ��� ������ �� �б�
    Open path For Binary As #1
        fileContent = Space$(LOF(1))
        Get #1, , fileContent
    Close #1
    
'    Open path For Input As #1
''        Do While Not EOF(1)
'            'Input #1, fileContent
'            If LOF(1) > 0 Then
'                fileContent = Input$(LOF(1), 1)
'            End If
''        Loop
'    Close #1
    
    lines = Split(fileContent, vbCrLf)
    
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
        
    ġƮŰ_����.Offset(2, 1).Resize(j, 1) = strPresetList
    
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
    
    Close
        
'    Else
'        Open path For Input As #1
'        fileContent = Input$(LOF(1), 1)
'        Close #1
'
'        lines = Split(fileContent, vbCrLf)
'
'        For i = 0 To UBound(lines)
'            If lines(i) = "<Mag_CreatItem2>" Then
'                check = True
'            End If
'
'            If check = True Then
'
'                    modifiedContent = modifiedContent & lines(i) & vbCrLf
'
'            End If
'        Next
'
'        Open path For Output As #1
'
'        Print #1, "<Mag_CreatItem>"
'
'        For Each cell In ġƮŰ
'
'            '��ȸ�� TID~~ �� ����
'            If InStr(cell.Value, "��ȸ��") = 0 Then
'                Print #1, cell.Value '---�ۼ��� ġƮ �޸��忡 �Է�
'            End If
'
'        Next
'
'        Print #1, modifiedContent
'        Close #1
    
    ġƮŰ_����.Offset(-1, 0).Value = "M1.CheatUsingPreset " & path & " """ & strPreset & """"
    
    Call UpdateEnd

End Sub

Public Sub OpenTxt()
    
    Dim path As String
    
    path = ThisWorkbook.path & "\Mag_Cheat.txt"

    If Dir(path, vbDirectory) = "" Then
        MsgBox "�޸����� �������ּ���."
        Exit Sub
    End If
    
    Shell "notepad.exe " & Chr(34) & path & Chr(34), vbNormalFocus

End Sub
