Attribute VB_Name = "ModuleDesign"
'=====================================================================
'UI ������ ���� ���
'=====================================================================

'###������ ���� ����###
Public colorCategorySel As Variant
Public colorUserInput As Variant

'=====================================================================
'���� �޴� �����
Sub HideHomeMenu()

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
Public Sub SetColor()

    colorCategorySel = RGB(166, 201, 236) '---����ڰ� ���� �� ǥ�õǴ� ����
        
    colorUserInput = RGB(255, 178, 111) '---����ڰ� �Է��� �� �ִ� ����
    
End Sub

'=====================================================================
'�˻��� �����Ͱ� ���� �� ������ ���� �����
Public Function HideHomeData()
       
    Dim data_range As Range '---�����Ͱ� ǥ�õǴ� ���� ����
    
    Call SetRange '---�������� ����ϴ� ���� ��ġ ȣ��
    
    Set data_range = Sheets("Home").Columns("I:K") '---������ ���� ����
    
    '---���� ȣ��� ������ ���� üũ �� ����, ǥ�� ó��
    If ���������� = Empty Then
        
        data_range.Hidden = True
        
    Else
    
        data_range.Hidden = False
        
    End If
    
    Sheets("Home").Range("A1").Select
    
End Function

'=====================================================================
'�� ���� ���� �����
Public Sub HideHomeCategory()
    
    '---�� ���� ���� ����
    Dim rngCategory As Range
    
    Set rngCategory = Sheets("Home").Columns("G:H")
    
    Call SetRange '---�������� ����ϴ� ���� ��ġ ȣ��
    
    If ���������� = Empty Then
        
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
