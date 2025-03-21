Attribute VB_Name = "ModuleButton"
'=====================================================================
'UI 버튼 동작 모듈
'=====================================================================

'카테고리 추가 버튼
Public Sub Button_AddCategory()
        
    Call UpdateStart
    
    If AddCategory = 0 Then
        
        GoTo event_exe
        
    End If
        
    Call UpdateEnd
    Exit Sub

'---선택된 열 존재 시 처리
event_exe:
    
    Call UpdateEnd
    검색어_시작.Value = "" '---Home 시트 이벤트 동작
        
    '---가져온 데이터 열 너비 자동 맞춤
    Range("DATA").EntireColumn.AutoFit
    
End Sub

'=====================================================================
'카테고리 선택 초기화 버튼
Public Sub Button_ResetCategory()

    Call UpdateStart
    
    Call ResetCategory
        
    Call UpdateEnd

End Sub

'=====================================================================
'카테고리 전체 선택 버튼
Public Sub Button_SelectAllCategory()
    
    Call UpdateStart
    
    SelectAllCategory
               
    Call UpdateEnd

End Sub

'=====================================================================
'검색 초기화 버튼
Public Sub Button_ResetSearch()

    Call UpdateStart
    
    If ResetSearch = 0 Then
        
        GoTo event_exe
        
    End If
            
    Call UpdateEnd

'---선택된 열 존재 시 처리
event_exe:

    Call UpdateEnd
    검색어_시작.Value = "" '---Home 시트 이벤트 동작
End Sub

