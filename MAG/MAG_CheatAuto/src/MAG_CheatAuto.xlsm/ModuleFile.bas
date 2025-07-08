Attribute VB_Name = "ModuleFile"
Option Explicit

'======================================================================================================
'[��� ����] ��ư Ŭ�� �� �� ǥ��


Public Sub FloderForm()

    UserForm1.Show
    
End Sub

'======================================================================================================
'[������ ����] ��ư Ŭ�� �� ����

Public Sub RefreshData()
    
    Dim strFolder As String
    
    strFolder = LatestFolder
    
    If strFolder = "" Then
    
        Exit Sub
        
    End If
    
    '����ڰ� ������ ���� ��η� ���� ��� ����
    ActiveWorkbook.Queries.Item("Address").Formula = Chr(34) & strFolder & Chr(34) & " meta [IsParameterQuery=true, Type=""Any"", IsParameterQueryRequired=true]"
    
    ActiveWorkbook.RefreshAll
    
End Sub

Public Sub WriteTextCheat()
    
    Dim filePath As String
    
    filePath = ThisWorkbook.path
    
End Sub

'======================================================================================================
'������ ������ ����� PC�� �����ϴ� �� Ȯ��


Public Function CheckFolder() As Boolean
    
    Dim strFolder As String
    
    strFolder = Sheets("etc").Range("H2").Value
        
    '�Էµ� ��ΰ� �����ϴ��� Ȯ��
    If Dir(strFolder, vbDirectory) = "" Then
        
        Sheets("etc").Range("H2").Value = "C:\Users\" & Environ("USERNAME") & "\Downloads"
        
        CheckFolder = True
        
        Exit Function
    End If
    
End Function

'======================================================================================================
'������ ������ ����� PC�� �����ϴ� �� Ȯ��


Public Function CheckFile(strFilePath As String) As Boolean
        
    '�Էµ� ��ΰ� �����ϴ��� Ȯ��
    If Dir(strFilePath, vbDirectory) = "" Then
    
        MsgBox strFilePath & " ������ �������� �ʴ� �����Դϴ�." & vbCrLf & vbCrLf & _
                "��θ� Ȯ�����ּ���."
                 
        CheckFile = True
        
        Exit Function
        
    End If
    
End Function

'======================================================================================================
'������ ���� ��ο��� ���� �ֽ� ������ ���� ����


Public Function LatestFolder()
    
    Dim path As String '---���� ��� ���� ����
    Dim strFolderList As String
    Dim strSpl As Variant '---������ �и� �� ���� ����
    Dim temp As Long
    Dim strLatest As String
    
    path = Sheets("etc").Range("H2").Value '---���� ������ ���� ��� ����
    
    LatestFolder = path
    
    Sheets("etc").Columns("E:E").Clear
    
    cnt = 1
    
    '���õ� ��ΰ� MAG ������ ������ �ƴ� ��� ����
    If InStr(path, "TFD_") = 0 Then
        
        path = path & "\"
        
        strFolderList = Dir(path, vbDirectory) '---���õ� ��� ������ �ִ� ���� ����Ʈ ����
        
        Do While strFolderList <> ""
            
            '�������� . .. �� ��� ����
            If strFolderList <> "." And strFolderList <> ".." And InStr(strFolderList, "TFD_") > 0 Then
            
                '��ΰ� ���� ������ �´��� Ȯ��
                If (GetAttr(path & strFolderList) And vbDirectory) = vbDirectory Then
                
                    Sheets("etc").Cells(cnt, 5).Value = strFolderList
                    
                    cnt = cnt + 1
                        
                End If
            End If
            
            strFolderList = Dir
            
        Loop
        
        '�˻��� ������ ���� �� ����
        If cnt = 1 Then
            
            LatestFolder = path
            
            Exit Function
            
        Else
        
            temp = 0
            
            '�˻��� ���� ������ŭ �ݺ�
            For i = 1 To cnt
            
                With Sheets("etc").Cells(i, 5)
                    
                    strSpl = Split(.Value, "_CL") '---CL �������� �������� ����
                            
                    '������ _CL �� �����ϴ��� üũ
                    If UBound(strSpl) = 1 Then
                    
                        
                        '������ �ֽ� üũ
                        If Val(strSpl(1)) > temp Then
                        
                            temp = Val(strSpl(1))
                            
                            strLatest = .Value
                            
                        End If
                    End If
                End With
                
            Next
            
            '���� ��� ����
            LatestFolder = path & "\" & strLatest
            
        End If
    End If

End Function
