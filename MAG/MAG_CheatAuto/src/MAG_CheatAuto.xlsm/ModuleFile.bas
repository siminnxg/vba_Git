Attribute VB_Name = "ModuleFile"
Option Explicit

Public Sub FloderForm()
    UserForm1.Show
End Sub

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

Public Function CheckFolder() As Boolean
    
    Dim strFolder As String
    
    strFolder = Sheets("etc").Range("H2").Value
        
    '�Էµ� ��ΰ� �����ϴ��� Ȯ��
    If Dir(strFolder, vbDirectory) = "" Then
        MsgBox "���� ������ ��δ� �������� �ʴ� ����Դϴ�." & vbCrLf & _
                "��θ� �缳�����ּ���."
                
        CheckFolder = True
        
        Exit Function
    End If
    
End Function

Public Function CheckFile(strFilePath As String) As Boolean
        
    '�Էµ� ��ΰ� �����ϴ��� Ȯ��
    If Dir(strFilePath, vbDirectory) = "" Then
        MsgBox strFilePath & " ������ �������� �ʴ� �����Դϴ�." & vbCrLf & vbCrLf & _
                "��θ� Ȯ�����ּ���."
                
        CheckFile = True
        Exit Function
    End If
    
End Function

Public Function LatestFolder()
    
    Dim row As Integer
    Dim path As String
    Dim strFolderList As String
    Dim strSpl As Variant
    Dim temp As Long
    Dim strLatest As String
    Dim strFolder As String
    
    strFolder = Sheets("etc").Range("H2").Value
    
    LatestFolder = strFolder
    
    Sheets("etc").Columns("E:E").Clear
    
    row = 1
    
    '���õ� ��ΰ� MAG ������ ������ �ƴ� ��� ����
    If InStr(strFolder, "TFD_") = 0 Then
        
        strFolder = strFolder & "\"
        
        strFolderList = Dir(strFolder, vbDirectory)
        
        Do While strFolderList <> ""
            
            '�������� . .. �� ��� ����
            If strFolderList <> "." And strFolderList <> ".." And InStr(strFolderList, "TFD_") > 0 Then
            
                '�������� Ȯ��
                If (GetAttr(strFolder & strFolderList) And vbDirectory) = vbDirectory Then
                
                    Sheets("etc").Cells(row, 5).Value = strFolderList
                    row = row + 1
                        
                End If
                
            End If
            
            strFolderList = Dir
        Loop
        
        '�˻��� ������ ���� ��� �˸� ó��
        If row = 1 Then
            MsgBox "������ ��ο� MAG ������ ������ �������� �ʽ��ϴ�."
            LatestFolder = ""
            Exit Function
            
        Else
            temp = 0
            
            '�˻��� ���� ������ŭ �ݺ�
            For i = 1 To row
            
                With Sheets("etc").Cells(i, 5)
                    
                    strSpl = Split(.Value, "_CL")
                            
                    '������ _CL �� �����ϴ��� üũ
                    If UBound(strSpl) = 1 Then
                        
                        '�ֽ� ���� üũ
                        If Val(strSpl(1)) > temp Then
                            temp = Val(strSpl(1))
                            strLatest = .Value
                        End If
                        
                    End If
                    
                End With
                
            Next
            
            '���� ��� ����
            LatestFolder = strFolder & "\" & strLatest
            
        End If
    End If

End Function
