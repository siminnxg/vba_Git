Attribute VB_Name = "ModuleMain"
Option Explicit


'##########################################################
'ParcelID 값에서 문자가 포함된 경우 ID 값만 추출 함수
'##########################################################


Function ID추출(strID As String) As Long
    
    Dim strSplit As Variant
    
    If strID = "" Then
        
        Exit Function
        
    End If
    
    '언더바가 포함되어 있으면 언더바 이후 ID 값만 추출
    If InStr(strID, "_") Then
        
        strSplit = Split(strID, "_")
        
        ID추출 = CLng(strSplit(1))
    
    '언더바가 없으면 전체 값 추출
    Else
        
        ID추출 = CLng(strID)
    
    End If
    
End Function

'##########################################################
'Result 영역에 입력하는 함수
'조건에 하나라도 걸리면 Fail 처리
'##########################################################


Function 결과확인(rngUser As Range, rngResult As Range) As String

        
    'GachaGroupID, ParcelID 공백인 경우 공백 표시
    If rngUser(1) = "" Or rngUser(2) = "" Then
        결과확인 = ""
        Exit Function
    End If
    
    '조건1 : GachaGroupID, ParcelID 일치하는 값이 없는 경우 전체 행 강조 표시
    If rngResult(1) = "" Then
        결과확인 = "FAIL"
        Exit Function
    End If


    '조건2 : Rarity 일치 여부 확인
    '대소문자 구분 x
    If StrComp(rngUser(3), rngResult(8), vbTextCompare) <> 0 Then
        결과확인 = "FAIL"
        Exit Function
    End If


    '조건3 : Name 값이 DevName 에 포함되는지 확인
    '띄어쓰기 무시, 대소문자 구분 x
    If rngUser(4) = "" Or _
        InStr(1, rngResult(7), rngUser(4), 1) = 0 And _
        InStr(1, Replace(rngResult(7), " ", ""), Replace(rngUser(4), " ", ""), 1) = 0 Then
        결과확인 = "FAIL"
        Exit Function
    End If


    '조건4 : Prob 값이 1 이상인지 확인
    If rngResult(11) < 1 Then
        결과확인 = "FAIL"
        Exit Function
    End If


    '조건5 : IsExport 값이 TRUE 인지 확인
    If rngResult(13).Value <> "True" Then
        결과확인 = "FAIL"
        Exit Function
    End If
    
    
    '모든 조건에서 통과 시 PASS 처리
    결과확인 = "PASS"
        
End Function


'##########################################################
'사용자가 설정한 경로 데이터 문서로 최신화
'##########################################################


Sub RefreshData()
    
    Dim strFolder As String
    Dim path As String
    Dim fSO As Object
    
문서검증:
    
    '불러오기
    strFolder = ThisWorkbook.CustomDocumentProperties("폴더경로")
    
    path = strFolder & "\GachaElement.xlsx"
    
    Set fSO = CreateObject("Scripting.FileSystemObject")
    
    '입력한 경로 내 문서 존재 여부 검증
    If CheckFile(strFolder) = True Then
        
        '선택된 폴더가 없으면 종료
        If SearchFolder = True Then
        
            Exit Sub
        
        '선택된 폴더가 있으면 문서 존재 여부 재 검증
        Else
            GoTo 문서검증
            
        End If
        
    End If
    
    '사용자가 지정한 폴더 경로로 쿼리 경로 변경
    ActiveWorkbook.Queries.Item("Address").Formula = Chr(34) & strFolder & Chr(34) & " meta [IsParameterQuery=true, Type=""Any"", IsParameterQueryRequired=true]"
    
    '쿼리 새로고침
    ActiveWorkbook.RefreshAll
    
    Sheets("Main").Range("L2") = "GachaElement 문서 마지막 수정 시간 : " & fSO.GetFile(path).Datelastmodified

End Sub


'##########################################################
'사용자가 입력한 경로에 GachaElement 문서 존재 여부 확인
'##########################################################


Function CheckFile(strAddress As String) As Boolean

    Dim path As String
    
    path = strAddress & "\GachaElement.xlsx"
    
    If Dir(path, vbDirectory) = "" Then
    
        MsgBox "'" & strAddress & "' 경로에 GachaElement.xlsx 문서가 존재하지 않습니다. " & vbCrLf & vbCrLf & _
                "경로를 설정해주세요."
        
        CheckFile = True
        
    End If

End Function


'##########################################################
'폴더 탐색기 오픈하여 폴더 경로 가져오기
'##########################################################


Function SearchFolder() As Boolean
    
    Dim Selected As Long '선택한 파일 정보 저장 변수
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "폴더를 선택하세요"
        
        Selected = .Show '파일 탐색기 열기
        
        '선택된 폴더가 있는 경우 동작
        If Selected = -1 Then
        
            ThisWorkbook.CustomDocumentProperties("폴더경로").Value = .SelectedItems(1)
        
        '선택된 폴더가 없는 경우 알림
        Else
        
            MsgBox "선택된 폴더가 없습니다."
            SearchFolder = True
            
        End If
    End With
    
End Function


'##########################################################
'시트 보호 해제
'##########################################################


Public Sub Unprotect()
    
    Dim sht As Worksheet
    
    Set sht = Sheets("Main")
    
    If sht.ProtectContents = True Then
        
        sht.Unprotect
        
    End If
    
End Sub

Public Sub ClearSht()

    Range("사용자입력").Clear
    
End Sub

Sub ShowForm()
    
    UserForm1.Show
    
End Sub
