Attribute VB_Name = "ModuleButton"
'=====================================================================
'UI ��ư ���� ���
'=====================================================================

'ī�װ� �߰� ��ư
Public Sub Button_AddCategory()
        
    Call UpdateStart
    Call SaveSearch '---�˻����̴� ���� ����
    
    Call AddCategory
        
    Call UpdateEnd
    Exit Sub

End Sub

'=====================================================================
'ī�װ� ���� �ʱ�ȭ ��ư
Public Sub Button_ResetCategory()

    Call UpdateStart
    
    Call ResetCategory
        
    Call UpdateEnd

End Sub

'=====================================================================
'ī�װ� ��ü ���� ��ư
Public Sub Button_SelectAllCategory()
    
    Call UpdateStart
    
    SelectAllCategory
               
    Call UpdateEnd

End Sub

'=====================================================================
'�˻� �ʱ�ȭ ��ư
Public Sub Button_ResetSearch()

    Call UpdateStart
    
    If ResetSearch = 0 Then
        
        Call PasteData
        
    End If
    
    Call UpdateEnd
    
End Sub

'=====================================================================
'�� ���� ��ư
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
'�˻� ������ �̵� ��ư
Public Sub Button_GotoSearch()
    
    If Sheets("Search").Visible = True Then
        Sheets("Search").Select
        
    End If
    
End Sub

'=====================================================================
'SearchForm ����
Public Sub Button_SearchForm()
    
    SearchForm.Show
    
End Sub

Public Sub Button_AutoFill()
    
    Call SetRange
    
    Call AutoFill(�˻�Ű����_����.Value)
    
End Sub

