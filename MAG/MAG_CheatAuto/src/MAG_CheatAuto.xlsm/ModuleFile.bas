Attribute VB_Name = "ModuleFile"
Option Explicit

Public Sub FloderForm()
    UserForm1.Show
End Sub

Public Sub RefreshData()
    
    '����ڰ� ������ ���� ��η� ���� ��� ����
    ActiveWorkbook.Queries.Item("Address").Formula = Chr(34) & Sheets("etc").Range("H2").Value & Chr(34) & " meta [IsParameterQuery=true, Type=""Any"", IsParameterQueryRequired=true]"
    
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


