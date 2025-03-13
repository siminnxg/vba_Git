Attribute VB_Name = "ModuleDesign"
'=====================================================================
'UI ������ ���� ���
'=====================================================================

'###������ ���� ����###
Public category_sel_color As Variant
Public user_input_color As Variant

'=====================================================================
'���� �޴� �����
Sub home_menu_hide()

    Dim menu_range As Range '---�޴� ���� ����
        
    Set menu_range = Sheets("Home").Columns("A:D")
    
    '---�޴� ���� ���� ���� Ȯ�� �� ����, ǥ�� ó��
    If menu_range.Hidden = True Then
    
        menu_range.Hidden = False
        
    Else
    
        menu_range.Hidden = True
                
    End If

End Sub

'=====================================================================
'���� ���� ����
Public Sub color_set()

    category_sel_color = RGB(166, 201, 236) '---����ڰ� ���� �� ǥ�õǴ� ����
        
    user_input_color = RGB(255, 178, 111) '---����ڰ� �Է��� �� �ִ� ����
    
End Sub

'=====================================================================
'�˻��� �����Ͱ� ���� �� ������ ���� �����
Public Function home_data_hide()
       
    Dim data_range As Range '---�����Ͱ� ǥ�õǴ� ���� ����
    
    Call range_set '---�������� ����ϴ� ���� ��ġ ȣ��
    
    Set data_range = Sheets("Home").Columns("I:K") '---������ ���� ����
    
    '---���� ȣ��� ������ ���� üũ �� ����, ǥ�� ó��
    If act_sheet_name = Empty Then
        
        data_range.Hidden = True
        
    Else
    
        data_range.Hidden = False
        
    End If
    
    Sheets("Home").Range("A1").Select
    
End Function

'=====================================================================
'�� ���� ���� �����
Public Sub HideCategoryRng()
    
    '---�� ���� ���� ����
    Dim rngCategory As Range
    
    Set rngCategory = Sheets("Home").Columns("G:H")
    
    Call range_set '---�������� ����ϴ� ���� ��ġ ȣ��
    
    If act_sheet_name = Empty Then
        
        If rngCategory.Hidden = False Then
            
            Sheets("Home").Shapes("Pic_Open").Visible = True
            Sheets("Home").Shapes("PIC_Close").Visible = False
        
            rngCategory.Hidden = True
        End If
        
        Exit Sub
        
    End If
    
    '---�޴� ���� ���� ���� Ȯ�� �� ����, ǥ�� ó��
    If rngCategory.Hidden = True Then
    
        Sheets("Home").Shapes("Pic_Open").Visible = False
        Sheets("Home").Shapes("PIC_Close").Visible = True
        
        rngCategory.Hidden = False
        
    Else
    
        Sheets("Home").Shapes("Pic_Open").Visible = True
        Sheets("Home").Shapes("PIC_Close").Visible = False
        
        rngCategory.Hidden = True
                
    End If
    
End Sub
