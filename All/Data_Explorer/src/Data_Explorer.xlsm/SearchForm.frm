VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchForm 
   Caption         =   "고급 검색"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10185
   OleObjectBlob   =   "SearchForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    
    Dim i As Variant
    
    Call SetRange
    
    For i = 0 To Range(검색키워드_시작, 검색키워드_끝).Cells.Count
        Me.List_Category.AddItem 검색키워드_시작.Offset(0, i).Value
    Next
    
End Sub

'=====================================================================
'자동 채우기
Private Sub Button_AutoFill_Click()
    
    On Error Resume Next
    
    '선택 여부 체크
    If IsNull(Me.List_Category.Value) Then
        MsgBox "선택된 열이 없습니다."
        Exit Sub
    End If
    
    Call SetRange
    Call AutoFill(Me.List_Category.Value)
    
End Sub

Private Sub Button_Refresh_Click()
    
    On Error Resume Next
    
    Call SetRange
    
    '프리셋 이름으로 생성된 시트의 listobject 새로고침
    Sheets(현재프리셋.Value).ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
    
    검색어_시작.Value = ""
    
End Sub


Private Sub Button_UniqueLoad_Click()
    
    Dim varSelCol As Variant
    Dim objData As Object
    Dim varUnique As Variant
    
    On Error Resume Next
    
    '선택 여부 체크
    If IsNull(Me.List_Category.Value) Then
        MsgBox "선택된 열이 없습니다."
        Exit Sub
    End If
        
    Call SetRange
    
    Set objData = Sheets(현재프리셋.Value).ListObjects(1)
    varSelCol = objData.ListColumns(Me.List_Category.Value).Index
    
    varUnique = Application.WorksheetFunction.Unique(Sheets(현재프리셋.Value).Columns(varSelCol)) '---선택된 열의 고유값 추출
    
    With Me.List_Search
    .List = varUnique '---리스트 박스에 표시
    .List(0) = Me.List_Category.ListIndex & " : " & .List(0)
    End With
End Sub

'=====================================================================
'리스트 박스 선택 초기화
Private Sub Button_SelClear_Click()
        
    Dim i As Variant
    
    On Error Resume Next
    
    With Me.List_Search
        For i = 0 To .ListCount
        
            Me.List_Search.Selected(i) = False
            
        Next
    End With
End Sub

Private Sub Button_Copy_Click()

    Dim i As Variant
    Dim varSelCol As Variant
    Dim strSearch As String
    
    On Error Resume Next
    
    Call SetRange
    
    '선택 값 확인
    With Me.List_Search
    
        varSelCol = Split(.List(0), " ") '---첫번째 값 공백으로 구분하여 열 순서 찾기
              
        For i = 1 To .ListCount - 1
            If .Selected(i) = True Then
                
                '선택 값 한 문장으로 결합
                If strSearch = vbNullString Then
                
                    strSearch = .List(i)
                    
                Else
                
                    strSearch = strSearch & "," & .List(i)
                    
                End If
            End If
        Next
    End With
    
    '선택 값 없는 경우 처리
    If strSearch = vbNullString Then
        MsgBox "선택된 값이 없습니다."
        Exit Sub
    End If
    
    검색어_시작.Offset(0, varSelCol(0)) = strSearch
    
End Sub

