Attribute VB_Name = "ModuleButton"
'=====================================================================
'UI ��ư ���� ���
'=====================================================================

'ī�װ� �߰� ��ư
Public Sub Button_AddCategory()
        
    Call UpdateStart
    
    If AddCategory = 0 Then
        
        GoTo event_exe
        
    End If
        
    Call UpdateEnd
    Exit Sub

'---���õ� �� ���� �� ó��
event_exe:
    
    Call UpdateEnd
    �˻���_����.Value = "" '---Home ��Ʈ �̺�Ʈ ����
        
    '---������ ������ �� �ʺ� �ڵ� ����
    Range("DATA").EntireColumn.AutoFit
    
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
        
        GoTo event_exe
        
    End If
            
    Call UpdateEnd

'---���õ� �� ���� �� ó��
event_exe:

    Call UpdateEnd
    �˻���_����.Value = "" '---Home ��Ʈ �̺�Ʈ ����
End Sub

