Attribute VB_Name = "ModuleButton"
'=====================================================================
'UI ��ư ���� ���
'=====================================================================

'ī�װ� �߰� ��ư
Public Sub button_category_add()
        
    Call update_start
    
    If category_add = 0 Then
        
        GoTo event_exe
        
    End If
        
    Call update_end
    Exit Sub

'---���õ� �� ���� �� ó��
event_exe:
    
    Call update_end
    search_user_start.Value = "" '---Home ��Ʈ �̺�Ʈ ����
        
    '---������ ������ �� �ʺ� �ڵ� ����
    Range("DATA").EntireColumn.AutoFit
    
End Sub

'=====================================================================
'ī�װ� ���� �ʱ�ȭ ��ư
Public Sub button_category_clickreset()

    Call update_start
    
    Call category_reset
        
    Call update_end

End Sub

'=====================================================================
'ī�װ� ��ü ���� ��ư
Public Sub button_category_allselect()
    
    Call update_start
    
    category_all_select
               
    Call update_end

End Sub

'=====================================================================
'�˻� �ʱ�ȭ ��ư
Public Sub button_search_reset()

    Call update_start
    
    If search_reset = 0 Then
        
        GoTo event_exe
        
    End If
            
    Call update_end

'---���õ� �� ���� �� ó��
event_exe:

    Call update_end
    search_user_start.Value = "" '---Home ��Ʈ �̺�Ʈ ����
End Sub

