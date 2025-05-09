Attribute VB_Name = "ModuleFile"
Option Explicit

Public Sub FloderForm()
    UserForm1.Show
End Sub

Public Sub RefreshData()
    
    '사용자가 지정한 폴더 경로로 쿼리 경로 변경
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
        
    '입력된 경로가 존재하는지 확인
    If Dir(strFolder, vbDirectory) = "" Then
        MsgBox "현재 설정된 경로는 존재하지 않는 경로입니다." & vbCrLf & _
                "경로를 재설정해주세요."
                
        CheckFolder = True
        
        Exit Function
    End If
    
End Function

Public Function CheckFile(strFilePath As String) As Boolean
        
    '입력된 경로가 존재하는지 확인
    If Dir(strFilePath, vbDirectory) = "" Then
        MsgBox strFilePath & " 파일은 존재하지 않는 파일입니다." & vbCrLf & vbCrLf & _
                "경로를 확인해주세요."
                
        CheckFile = True
        Exit Function
    End If
    
End Function


