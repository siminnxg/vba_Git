Attribute VB_Name = "ModuleButton"
'=====================================================================
'UI 버튼 동작 모듈
'=====================================================================

'카테고리 추가 버튼
Public Sub button_category_add()
        
    Call update_start
    
    If category_add = 0 Then
        
        GoTo event_exe
        
    End If
        
    Call update_end
    Exit Sub

'---선택된 열 존재 시 처리
event_exe:
    
    Call update_end
    search_user_start.Value = "" '---Home 시트 이벤트 동작
        
    '---가져온 데이터 열 너비 자동 맞춤
    Range("DATA").EntireColumn.AutoFit
    
End Sub

'=====================================================================
'카테고리 선택 초기화 버튼
Public Sub button_category_clickreset()

    Call update_start
    
    Call category_reset
        
    Call update_end

End Sub

'=====================================================================
'카테고리 전체 선택 버튼
Public Sub button_category_allselect()
    
    Call update_start
    
    category_all_select
               
    Call update_end

End Sub

'=====================================================================
'검색 초기화 버튼
Public Sub button_search_reset()

    Call update_start
    
    If search_reset = 0 Then
        
        GoTo event_exe
        
    End If
            
    Call update_end

'---선택된 열 존재 시 처리
event_exe:

    Call update_end
    search_user_start.Value = "" '---Home 시트 이벤트 동작
End Sub

