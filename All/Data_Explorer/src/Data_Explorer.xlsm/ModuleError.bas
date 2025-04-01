Attribute VB_Name = "ModuleError"
Option Explicit
'=====================================================================
'예외처리 관련 모듈
'=====================================================================

'동일한 시트명 체크
Public Function CheckQuery()
    
    Dim sheet_count As Variant '---시트 개수 저장 변수
        
    sheet_count = ActiveWorkbook.Sheets.Count '---현재 파일에서 시트 개수 체크
    
    '---시트 개수만큼 반복
    For i = 1 To sheet_count
        
        If preset = ActiveWorkbook.Sheets(i).Name Then '---입력한 프리셋 명이 현재 생성되어있는 시트와 동일한 이름인지 확인
        
            CheckQuery = 1
        
        End If
    Next
    
End Function

'=====================================================================
'열 리스트 호출 여부 확인
Public Function CheckCategory()
    
    '---열 리스트 영역 첫번째 셀 공백 체크
    If 열목록_시작.Value = "" Then
    
        CheckCategory = 1
        
    End If
    
End Function

'=====================================================================
'파일 존재 여부 체크
Public Function CheckFile(ByVal path_ As String) As Boolean
        
    CheckFile = (Dir(path_, vbDirectory) <> "") '---입력된 경로에 입력된 파일명이 존재하는지 확인
 
End Function

'=====================================================================
'파일 오픈 상태 체크
Public Function CheckFileOpen(CheckFile As String) As Boolean

    Dim wb As Variant
    
    On Error Resume Next
    
    Set wb = Workbooks(CheckFile)
        
    If Not wb Is Nothing Then
    
        CheckFileOpen = True
        
    Else
    
        CheckFileOpen = False
        
    End If
    
    On Error GoTo 0
    
End Function

'=====================================================================
'프리셋 이름 체크
Public Function CheckPresetName()
    
    Dim preset_name_index As Variant '---프리셋명 순서
    Dim check As Boolean '---동일한 프리셋명 체크
    
    preset_name_index = 1
    
    With Range("preset_list")
        Do
            For i = 2 To .Cells.Count
                
                '---동일한 프리셋 이름이 존재하는 경우 처리
                If StrComp(.Cells(i).Value, "프리셋" & preset_name_index) = 0 Then
                
                    preset_name_index = preset_name_index + 1
                    check = True
                    Exit For
                    
                End If
                
                check = False
            Next
            
            '---동일한 프리셋 이름이 없으면 종료
            If check = False Then
            
                Exit Do
                
            End If
        Loop
    End With
    
    '---프리셋 이름 반환
    CheckPresetName = "프리셋" & preset_name_index
    
End Function

