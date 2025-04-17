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
