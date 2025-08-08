Attribute VB_Name = "ModuleMain"
Option Explicit


'##########################################################
'ParcelID ������ ���ڰ� ���Ե� ��� ID ���� ���� �Լ�
'##########################################################


Function ID����(strID As String) As Long
    
    Dim strSplit As Variant
    
    If strID = "" Then
        
        Exit Function
        
    End If
    
    '����ٰ� ���ԵǾ� ������ ����� ���� ID ���� ����
    If InStr(strID, "_") Then
        
        strSplit = Split(strID, "_")
        
        ID���� = CLng(strSplit(1))
    
    '����ٰ� ������ ��ü �� ����
    Else
        
        ID���� = CLng(strID)
    
    End If
    
End Function

'##########################################################
'Result ������ �Է��ϴ� �Լ�
'���ǿ� �ϳ��� �ɸ��� Fail ó��
'##########################################################


Function ���Ȯ��(rngUser As Range, rngResult As Range) As String

        
    'GachaGroupID, ParcelID ������ ��� ���� ǥ��
    If rngUser(1) = "" Or rngUser(2) = "" Then
        ���Ȯ�� = ""
        Exit Function
    End If
    
    '����1 : GachaGroupID, ParcelID ��ġ�ϴ� ���� ���� ��� ��ü �� ���� ǥ��
    If rngResult(1) = "" Then
        ���Ȯ�� = "FAIL"
        Exit Function
    End If


    '����2 : Rarity ��ġ ���� Ȯ��
    '��ҹ��� ���� x
    If StrComp(rngUser(3), rngResult(8), vbTextCompare) <> 0 Then
        ���Ȯ�� = "FAIL"
        Exit Function
    End If


    '����3 : Name ���� DevName �� ���ԵǴ��� Ȯ��
    '���� ����, ��ҹ��� ���� x
    If rngUser(4) = "" Or _
        InStr(1, rngResult(7), rngUser(4), 1) = 0 And _
        InStr(1, Replace(rngResult(7), " ", ""), Replace(rngUser(4), " ", ""), 1) = 0 Then
        ���Ȯ�� = "FAIL"
        Exit Function
    End If


    '����4 : Prob ���� 1 �̻����� Ȯ��
    If rngResult(11) < 1 Then
        ���Ȯ�� = "FAIL"
        Exit Function
    End If


    '����5 : IsExport ���� TRUE ���� Ȯ��
    If rngResult(13).Value <> "True" Then
        ���Ȯ�� = "FAIL"
        Exit Function
    End If
    
    
    '��� ���ǿ��� ��� �� PASS ó��
    ���Ȯ�� = "PASS"
        
End Function


'##########################################################
'����ڰ� ������ ��� ������ ������ �ֽ�ȭ
'##########################################################


Sub RefreshData()
    
    Dim strFolder As String
    Dim path As String
    Dim fSO As Object
    
��������:
    
    '�ҷ�����
    strFolder = ThisWorkbook.CustomDocumentProperties("�������")
    
    path = strFolder & "\GachaElement.xlsx"
    
    Set fSO = CreateObject("Scripting.FileSystemObject")
    
    '�Է��� ��� �� ���� ���� ���� ����
    If CheckFile(strFolder) = True Then
        
        '���õ� ������ ������ ����
        If SearchFolder = True Then
        
            Exit Sub
        
        '���õ� ������ ������ ���� ���� ���� �� ����
        Else
            GoTo ��������
            
        End If
        
    End If
    
    '����ڰ� ������ ���� ��η� ���� ��� ����
    ActiveWorkbook.Queries.Item("Address").Formula = Chr(34) & strFolder & Chr(34) & " meta [IsParameterQuery=true, Type=""Any"", IsParameterQueryRequired=true]"
    
    '���� ���ΰ�ħ
    ActiveWorkbook.RefreshAll
    
    Sheets("Main").Range("L2") = "GachaElement ���� ������ ���� �ð� : " & fSO.GetFile(path).Datelastmodified

End Sub


'##########################################################
'����ڰ� �Է��� ��ο� GachaElement ���� ���� ���� Ȯ��
'##########################################################


Function CheckFile(strAddress As String) As Boolean

    Dim path As String
    
    path = strAddress & "\GachaElement.xlsx"
    
    If Dir(path, vbDirectory) = "" Then
    
        MsgBox "'" & strAddress & "' ��ο� GachaElement.xlsx ������ �������� �ʽ��ϴ�. " & vbCrLf & vbCrLf & _
                "��θ� �������ּ���."
        
        CheckFile = True
        
    End If

End Function


'##########################################################
'���� Ž���� �����Ͽ� ���� ��� ��������
'##########################################################


Function SearchFolder() As Boolean
    
    Dim Selected As Long '������ ���� ���� ���� ����
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "������ �����ϼ���"
        
        Selected = .Show '���� Ž���� ����
        
        '���õ� ������ �ִ� ��� ����
        If Selected = -1 Then
        
            ThisWorkbook.CustomDocumentProperties("�������").Value = .SelectedItems(1)
        
        '���õ� ������ ���� ��� �˸�
        Else
        
            MsgBox "���õ� ������ �����ϴ�."
            SearchFolder = True
            
        End If
    End With
    
End Function


'##########################################################
'��Ʈ ��ȣ ����
'##########################################################


Public Sub Unprotect()
    
    Dim sht As Worksheet
    
    Set sht = Sheets("Main")
    
    If sht.ProtectContents = True Then
        
        sht.Unprotect
        
    End If
    
End Sub

Public Sub ClearSht()

    Range("������Է�").Clear
    
End Sub

Sub ShowForm()
    
    UserForm1.Show
    
End Sub
