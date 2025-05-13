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
    
    '사용자가 지정한 폴더 경로로 쿼리 경로 변경
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
    
    '선택된 경로가 MAG 데이터 폴더가 아닌 경우 동작
    If InStr(strFolder, "TFD_") = 0 Then
        
        strFolder = strFolder & "\"
        
        strFolderList = Dir(strFolder, vbDirectory)
        
        Do While strFolderList <> ""
            
            '폴더명이 . .. 인 경우 제외
            If strFolderList <> "." And strFolderList <> ".." And InStr(strFolderList, "TFD_") > 0 Then
            
                '폴더인지 확인
                If (GetAttr(strFolder & strFolderList) And vbDirectory) = vbDirectory Then
                
                    Sheets("etc").Cells(row, 5).Value = strFolderList
                    row = row + 1
                        
                End If
                
            End If
            
            strFolderList = Dir
        Loop
        
        '검색된 폴더가 없는 경우 알림 처리
        If row = 1 Then
            MsgBox "설정된 경로에 MAG 데이터 폴더가 존재하지 않습니다."
            LatestFolder = ""
            Exit Function
            
        Else
            temp = 0
            
            '검색된 폴더 개수만큼 반복
            For i = 1 To row
            
                With Sheets("etc").Cells(i, 5)
                    
                    strSpl = Split(.Value, "_CL")
                            
                    '폴더명에 _CL 이 존재하는지 체크
                    If UBound(strSpl) = 1 Then
                        
                        '최신 버전 체크
                        If Val(strSpl(1)) > temp Then
                            temp = Val(strSpl(1))
                            strLatest = .Value
                        End If
                        
                    End If
                    
                End With
                
            Next
            
            '폴더 경로 지정
            LatestFolder = strFolder & "\" & strLatest
            
        End If
    End If

End Function
