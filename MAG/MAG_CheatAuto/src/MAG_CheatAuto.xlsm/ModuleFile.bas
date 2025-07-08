Attribute VB_Name = "ModuleFile"
Option Explicit

'======================================================================================================
'[경로 변경] 버튼 클릭 시 폼 표시


Public Sub FloderForm()

    UserForm1.Show
    
End Sub

'======================================================================================================
'[데이터 갱신] 버튼 클릭 시 동작

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

'======================================================================================================
'설정된 폴더가 사용자 PC에 존재하는 지 확인


Public Function CheckFolder() As Boolean
    
    Dim strFolder As String
    
    strFolder = Sheets("etc").Range("H2").Value
        
    '입력된 경로가 존재하는지 확인
    If Dir(strFolder, vbDirectory) = "" Then
        
        Sheets("etc").Range("H2").Value = "C:\Users\" & Environ("USERNAME") & "\Downloads"
        
        CheckFolder = True
        
        Exit Function
    End If
    
End Function

'======================================================================================================
'설정된 파일이 사용자 PC에 존재하는 지 확인


Public Function CheckFile(strFilePath As String) As Boolean
        
    '입력된 경로가 존재하는지 확인
    If Dir(strFilePath, vbDirectory) = "" Then
    
        MsgBox strFilePath & " 파일은 존재하지 않는 파일입니다." & vbCrLf & vbCrLf & _
                "경로를 확인해주세요."
                 
        CheckFile = True
        
        Exit Function
        
    End If
    
End Function

'======================================================================================================
'설정된 폴더 경로에서 가장 최신 리비전 폴더 추출


Public Function LatestFolder()
    
    Dim path As String '---폴더 경로 저장 변수
    Dim strFolderList As String
    Dim strSpl As Variant '---폴더명 분리 후 저장 변수
    Dim temp As Long
    Dim strLatest As String
    
    path = Sheets("etc").Range("H2").Value '---현재 지정된 폴더 경로 저장
    
    LatestFolder = path
    
    Sheets("etc").Columns("E:E").Clear
    
    cnt = 1
    
    '선택된 경로가 MAG 데이터 폴더가 아닌 경우 동작
    If InStr(path, "TFD_") = 0 Then
        
        path = path & "\"
        
        strFolderList = Dir(path, vbDirectory) '---선택된 경로 하위에 있는 폴더 리스트 추출
        
        Do While strFolderList <> ""
            
            '폴더명이 . .. 인 경우 제외
            If strFolderList <> "." And strFolderList <> ".." And InStr(strFolderList, "TFD_") > 0 Then
            
                '경로가 폴더 형식이 맞는지 확인
                If (GetAttr(path & strFolderList) And vbDirectory) = vbDirectory Then
                
                    Sheets("etc").Cells(cnt, 5).Value = strFolderList
                    
                    cnt = cnt + 1
                        
                End If
            End If
            
            strFolderList = Dir
            
        Loop
        
        '검색된 폴더가 없을 때 종료
        If cnt = 1 Then
            
            LatestFolder = path
            
            Exit Function
            
        Else
        
            temp = 0
            
            '검색된 폴더 개수만큼 반복
            For i = 1 To cnt
            
                With Sheets("etc").Cells(i, 5)
                    
                    strSpl = Split(.Value, "_CL") '---CL 기준으로 리비전만 추출
                            
                    '폴더명에 _CL 이 존재하는지 체크
                    If UBound(strSpl) = 1 Then
                    
                        
                        '리비전 최신 체크
                        If Val(strSpl(1)) > temp Then
                        
                            temp = Val(strSpl(1))
                            
                            strLatest = .Value
                            
                        End If
                    End If
                End With
                
            Next
            
            '폴더 경로 지정
            LatestFolder = path & "\" & strLatest
            
        End If
    End If

End Function
