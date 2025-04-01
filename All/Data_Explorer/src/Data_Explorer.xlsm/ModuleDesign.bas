Attribute VB_Name = "ModuleDesign"
'=====================================================================
'UI ������ ���� ���
'=====================================================================

'###������ ���� ����###
Public colorCategorySel As Variant
Public colorUserInput As Variant

'=====================================================================
'���� ���� ����
Public Sub SetColor()

    colorCategorySel = RGB(166, 201, 236) '---����ڰ� ���� �� ǥ�õǴ� ����
        
    colorUserInput = RGB(255, 178, 111) '---����ڰ� �Է��� �� �ִ� ����
    
End Sub

'=====================================================================
'�˻��� �����Ͱ� ���� �� ������ ���� �����
Public Function HideSearchSht(booCheck As Boolean)
    
    '��Ʈ ����, ǥ�� ó��
    With Sheets("Search")
        If booCheck = True Then
            .Visible = 2
        Else
            .Visible = -1
            .Select
        End If
    End With
    
End Function

'=====================================================================
'�� ���� ���� �����
Public Sub HideHomeCategory(booCheck As Boolean)
        
    Dim rngCategory As Range '---�� ���� ���� ����
    
    Set rngCategory = Sheets("Search").Columns("B:C")
    
    Call SetRange
    
    '����, ǥ�� ó��
    If booCheck = True Then
        rngCategory.Hidden = True
        
        '�˻� ���� Ʋ ����
        �˻�Ű����_����.Offset(1, Ʋ����.Value).Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        Range("A1").Select
        
        
    Else
        rngCategory.Hidden = False
        
        '�� ���� ���� Ʋ ����
        Columns("D:D").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        Range("A1").Select
    End If
    
End Sub
