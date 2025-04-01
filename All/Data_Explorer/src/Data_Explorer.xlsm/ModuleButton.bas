Attribute VB_Name = "ModuleButton"
'=====================================================================
'UI 버튼 동작 모듈
'=====================================================================

'카테고리 추가 버튼
Public Sub Button_AddCategory()
        
    Call UpdateStart
    Call SaveSearch '---검색중이던 내용 저장
    
    Call AddCategory
        
    Call UpdateEnd
    Exit Sub

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
        
        Call PasteData
        
    End If
    
    Call UpdateEnd
    
End Sub

'=====================================================================
'열 선택 버튼
Public Sub Button_HideCategory()

    With Sheets("Search").Columns("B:C")
        If .Hidden = True Then
            Call HideHomeCategory(False)
        Else
            Call HideHomeCategory(True)
        End If
    End With
    
End Sub

'=====================================================================
'검색 페이지 이동 버튼
Public Sub Button_GotoSearch()
    
    If Sheets("Search").Visible = True Then
        Sheets("Search").Select
        
    End If
    
End Sub

'=====================================================================
'SearchForm 열기
Public Sub Button_SearchForm()
    
    SearchForm.Show
    
End Sub

Public Sub Button_AutoFill()
    
    Call SetRange
    
    Call AutoFill(검색키워드_시작.Value)
    
End Sub

