Attribute VB_Name = "ModuleError"
Option Explicit

'=====================================================================
'매크로 : CheckUserData
'대상 시트 : Main 시트
'동작 : 사용자가 데이터를 입력했는지 체크합니다.
'=====================================================================
Public Function CheckUserData() As Boolean
    
    '파일 경로 입력 체크
    If 파일경로(1) = "" Then
        MsgBox "파일 경로를 입력해주세요."
        CheckUserData = True
        Exit Function
    
    '검색어 입력 체크
    ElseIf 검색어 = "" Then
        MsgBox "검색어를 입력해주세요."
        CheckUserData = True
        Exit Function
        
    End If
    
    '파일명 입력 체크
    For i = 1 To 파일명.count
    
        If 파일명(i) = "" Then
            MsgBox "파일명을 입력해주세요."
            CheckUserData = True
            Exit Function
            
        End If
    Next
    
End Function

'=====================================================================
'매크로 : CheckFile
'동작 : 사용자가 입력한 파일이 실제로 존재하는지 체크합니다.
'=====================================================================
Public Function CheckFile() As Boolean
    
    Dim strFile As String
    
    '입력된 파일 개수만큼 반복
    For j = 1 To 파일명.count
        
        '엑셀 파일인지 확인
        If InStr(파일명(j), ".xl") = 0 Then
            MsgBox 파일명(i) & "은(는) 엑셀 형식의 파일이 아닙니다."
            CheckFile = True
            Exit Function
            
        End If
        
        strFile = 파일경로(j) & "\" & 파일명(j)
        
        '입력된 경로에 입력된 파일명이 존재하는지 확인
        If Dir(strFile, vbDirectory) = "" Then
            MsgBox strFile & "은(는) 존재하지 않는 파일입니다."
            CheckFile = True
            Exit Function
            
        End If
        
    Next
    
End Function

'=====================================================================
'매크로 : CheckSheet
'동작 : 사용자가 입력한 파일 내 입력한 시트명이 존재하는지 체크합니다.
'=====================================================================
Public Function CheckSheet(Wb, strSheet) As Boolean

    For j = 1 To Wb.Sheets.count
        
        '시트명이 일치하는지 체크
        If Wb.Sheets(j).Name = strSheet Then
            Exit Function
            
        End If
    Next
        
    CheckSheet = True
    
End Function
