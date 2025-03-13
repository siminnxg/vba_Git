Attribute VB_Name = "ModuleMain"
'=====================================================================
'��ũ�� : FIle_Load
'��� ��Ʈ : Home ��Ʈ
'���� : �ҷ����� ��ư ����, �ԷµǾ� �ִ� ���� ������ �������� ������ ȣ��
'=====================================================================
Public Sub FIle_Load()

    '###���� ����###
    
    Dim var As Variant '---�ӽ� ����
    Dim File_name_val As Variant '---Ȯ���� ������ ���ϸ�
    
    '###���� ����###
    
    '---������ �߻��ϸ� ����
    On Error GoTo exit_error
    
    '---sub : ȭ�� ������Ʈ ����
    Call update_start
    
    '---sub : �������� ����ϴ� ���� ��ġ ȣ��
    Call range_set
    
    '---����� �Է� ���� ���� �� �˸� ǥ��
    If File_adr = "" Or File_name = "" Or sheet_name = "" Then
    
        MsgBox "���� ������ ��� �Է����ּ���." & vbCrLf & "(���� ���, �̸�, ��Ʈ)"
        GoTo exit_sub

    End If
    
    '---function : ������ ���� �� �ӽ� �̸� ����
    If preset = "" Or preset = "������" Then
        
        preset = preset_name_check
        user_file_preset = preset
    
    End If
    
    '---���� �̸����� Ȯ����(.xl~) �и�
    var = Split(File_name, ".")
    File_name_val = var(0)
    
    '---function : ������ �̸����� ��Ʈ, ���� �̹� �����Ǿ� �ִٸ� ����
    If query_check = 1 Then
        
        MsgBox ("������ �����¸��� �����մϴ�.")
        GoTo exit_sub
    
    '---function : �Է��� ��ο� ���� ���� ���� üũ
    ElseIf FileExists(File_adr & "\" & File_name) = False Then
        
        MsgBox (File_adr & " ��ο� " & File_name & " ������ �������� �ʽ��ϴ�.")
        GoTo exit_sub
    
    '---������ ���� ������ �ƴ� ��� ó��
    ElseIf InStr(File_name, ".xl") = 0 Then
     
         MsgBox "������ ���� ������ �ƴմϴ�."
         GoTo exit_sub
         
    Else
    
        '---sub : �� ����, �˻� ���� �ʱ�ȭ
        Call home_data_clear
                
        '---������ �̸����� ��Ʈ ����
        ActiveWorkbook.Worksheets.Add after:=Sheets("Home")
        ActiveSheet.Name = preset
        
        '---�Էµ� ��θ� �������� ���� �ҷ�����
        ActiveWorkbook.Queries.Add Name:=preset, _
        Formula:="let Source = Excel.Workbook(File.Contents(""" & File_adr & "\" & File_name & """), null, true), #""" & _
                sheet_name & "_Sheet"" = Source{[Item=""" & sheet_name & """, Kind=""Sheet""]}[Data], " & _
                "FilteredData = Table.PromoteHeaders(#""" & sheet_name & "_Sheet"") " & _
        "in FilteredData"

        '---���� �߰� �� ��� �����͸� �ؽ�Ʈ Ÿ������ ���� �� ���
'                "��翭���� = Table.TransformColumnTypes(FilteredData, " & _
'        "List.Transform(Table.ColumnNames(FilteredData), each {_, type text})) " & _
'        "in ��翭����"
        
        '---����� ���� ������ ��������
        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & preset & ";Extended Properties=""""" _
            , Destination:=Range("$A$1")).QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [" & preset & "]")
            .Refresh BackgroundQuery:=False
        End With
        
        '---sub : ������ ����
        Call preset_save
        
    End If
        
    Sheets("Home").Select
    
    '---���� ȣ�� �� �� ���� ���� ���� ǥ��
    Sheets("Home").Columns("G:H").Hidden = False
    Sheets("Home").Shapes("Pic_Open").Visible = False
    Sheets("Home").Shapes("PIC_Close").Visible = True
    
    '---sub : ī�װ� ����Ʈ ȣ��
    Call search_list(preset)
    
    '---function : ī�װ� ��ü ���� �� �߰�
    If category_all_select = 0 Then
        
        varCheckUpdate = Empty
        Call category_all_select
        Call button_category_add
        
    End If

'---���� ó��
exit_sub:
    
    Call update_end
    Range("A1").Select
    Exit Sub
    
'---���� �߻� ó��
exit_error:
    
    '---��Ʈ�� ���� �� ó�� (��Ʈ�� ���� üũ, ������ ���� üũ)
    If Left(Err.Description, 6) = "�Է��� ��Ʈ" Or Err.Number = -2147024809 Then
        
        '---��Ʈ ���� �� �ý��� ���� �̳���
        Application.DisplayAlerts = False
        
        '---������ ��Ʈ ����
        ActiveSheet.Delete
        Sheets("Home").Select
        
        Application.DisplayAlerts = True
        
        '---�˸� ǥ��
        MsgBox "�����¸��� �������ּ���" & vbCrLf & vbCrLf & Err.Description
        GoTo exit_sub
        
    End If
    
    '---�� �� ���� ó��
    MsgBox ("������ �߻��߽��ϴ�. ���� �ڵ� : " & Err.Number & vbCrLf & Err.Description)
    
    Call update_end
    Range("A1").Select
    
End Sub

'=====================================================================
'��ũ�� : File_search
'��� ��Ʈ : Home ��Ʈ, etc ��Ʈ
'���� : ���� �˻� ��ư Ŭ�� �� ����, ���� Ž���� ���� �� ����ڰ� ������ ������ ���, �̸� ȣ��
'=====================================================================
Public Sub File_search()
    
    '###���� ����###
    
    Dim sel_File As Variant '---������ ���ϸ� ���� ����
    Dim wb As Object '---������ ���� ������Ʈ ���� ����
    Dim sheet_count As Variant '---�ҷ��� ������ ��Ʈ ���� ���� ����
    Dim file_open_check As String '---�ҷ��� ���� ���� ���� Ȯ�� ����
    Dim var() As String '---�ӽ� ��� ����
    
    '###���� ����###
    
    '--- sub �������� ����ϴ� ���� ��ġ, ���� ȣ��
    Call update_start
    Call range_set
    
    '---���� �ּ� ��������
    sel_File = File_adr & "\" & File_name
    
    '---etc��Ʈ �� ��Ʈ�� ���� ���� �ʱ�ȭ
    Sheets("etc").Range("A:A").Clear
    
    '---���� ��ΰ� �ԷµǾ� ������ �ش� ��η� ����
    '---(�߸��� ��� �Է� �� �ڵ����� ���õ�)
    If File_adr <> "" Then
     
         Application.FileDialog(msoFileDialogFilePicker).InitialFileName = File_adr
         
    End If
    
    '---���� Ž���� ����
     With Application.FileDialog(msoFileDialogFilePicker)
         .Filters.Add "��������", "*.xls; *.xlsx; *.xlsm" '---���� �������� ����
         .Show
         
         '---���� �� ���� �� ���� ó��
         If .SelectedItems.Count = 0 Then
         
             MsgBox "������ �������� �ʾҽ��ϴ�."
             GoTo exit_sub
             
         End If
         
         '---������ ������ sel_File ������ ����
         sel_File = .SelectedItems(1)
         
     End With
     
     '---������ ������ ���� ������ �ƴ� ��� ó��
     If InStr(sel_File, ".xl") = 0 Then
     
         MsgBox "���� ������ �������ּ���."
         GoTo exit_sub
         
     End If
                    
    '---���� ��� �� �̸� �и� �� ����
    max = InStrRev(sel_File, "\")
    user_file_adr.Value = Left(sel_File, max - 1)
    user_file_name.Value = Mid(sel_File, max + 1)
    
    '---sub : ���� ���� ��ȣ��
    Call file_info_load
           
    '---function : ����ڰ� ������ ������ �̹� �����ִ��� Ȯ��
    file_open_check = IsCheckOpen(user_file_name.Value)
               
    '---������ ������ ��Ʈ�� �ҷ�����
    Set wb = GetObject(sel_File)
    
    '---�ش� ������ ��Ʈ ���� Ȯ��
    sheet_count = wb.Sheets.Count
    
    '---��Ʈ ������ŭ ����
    For n = 1 To sheet_count
    
        Sheets("etc").Range("a" & n) = wb.Sheets(n).Name
        
    Next
    
    '---sheet_list �̸� ���� ������
    ThisWorkbook.Names("sheet_list").RefersTo = Sheets("etc").Range("A1", Sheets("etc").Cells(n - 1, 1))
    
    '---������ �������� �ʴ� �����̶�� ���� ����
    If file_open_check = True Then
        
        wb.Close
        
    End If
    
    '---��Ʈ ��� ��� �ٿ����� ǥ��
    With user_file_sheet.Validation
        .Delete
        .Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="=sheet_list"
    End With
    
    '---ù��° ��Ʈ�� �⺻������ ǥ��
    user_file_sheet.Value = Sheets("etc").Range("A1").Value
    
'---���� ó��
exit_sub:

    Call update_end
    Range("A1").Select
    
End Sub

'=====================================================================
'��ũ�� : search_list
'��� ��Ʈ : Home ��Ʈ
'���� : ���� �˻� ��ư Ŭ�� �� ����, ���� Ž���� ���� �� ����ڰ� ������ ������ ���, �̸� ȣ��
'=====================================================================
Public Sub search_list(sheet_name)
    
    '###���� ����###
        
    Dim category() As Variant   '---ī���� ������ ������ �迭
    Dim category_Row As Range   '---�ҷ��� ������ ��Ʈ���� ù �� ���� ����
    Dim Target As Range         '---���� ���õ� ������ �� ���� ����
    Dim category_range As Range '---ī�װ� ���� ����
    Dim varCategoryRow As Variant
    Dim data_range As ListObject
    
    
    '###���� ����###
        
    act_sheet_name.Value = sheet_name
    
    Set data_range = Sheets("������2").ListObjects(1)
    
    '---���� �����¸� etc ��Ʈ���� ��ġ �˻�
    search_row = Range("preset_list").Find(what:=act_sheet_name, lookat:=xlWhole).Row
    
    '---etc ������ ����Ʈ �������� �Ӹ��� �� �� ��������
     varCategoryRow = Range("preset_list")(search_row).Offset(0, 5).Value
     
     If varCategoryRow = Empty Then
        
        varCategoryRow = 1
        
     End If
    
    '---�� ����, �˻� ���� �����
    Call home_data_hide
    
    '---ī�װ� ����Ʈ ���� �ʱ�ȭ
    'act_category_list.Clear
    
    '---������ ��Ʈ�� ù��° �� ���� ����
    If varCategoryRow = 1 Or varCategoryRow > data_range.ListRows.Count Then
       
        Set category_range = data_range.HeaderRowRange
    
    Else
        
        Set category_range = data_range.ListRows(varCategoryRow - 1).Range
        
    End If
    
    '---�迭 ũ�� ������
    ReDim category(category_range.Columns.Count - 1)
    
    '---�迭�� ù���� �� �� ���� ����
    For i = 0 To category_range.Columns.Count - 1
        
        If category_range(i + 1) <> Empty Then
            category(i) = category_range(i + 1)
        End If
        
    Next
    
    '---Home ��Ʈ�� �迭 �� �ѷ��ֱ�
    For i = 0 To UBound(category)
        
        If category(i) <> Empty Then
        act_sheet_name.Offset(i + 1, 0).Value = category(i)
        End If
    Next
    
    '---ī�װ� ���� �� �Ҵ�
    Call range_set
    
    '---ī�װ� ���� �׵θ� ����
    With act_category_list
            
            .Borders.LineStyle = 1
            .Borders.Weight = xlThin
            .Borders.ColorIndex = 1
        
    End With
    
    '---���� �� ���� �� �ʱ�ȭ
    search_FixRow.ClearContents
        
End Sub

'=====================================================================
'��ũ�� : category_add
'��� ��Ʈ : Home ��Ʈ
'���� : ���� �˻� ��ư Ŭ�� �� ����, ���� Ž���� ���� �� ����ڰ� ������ ������ ���, �̸� ȣ��
'=====================================================================
Public Function category_add()

    'category_add = 1 : ���� �߻�
    'category_add = 0 : ���� ����
    
    '###���� ����###
    
    Dim now_cell As Range
    Dim category_count As Variant
    Dim select_category As Range
    Dim search_row As Variant
    
    '###���� ����###
    
    '---�������� ����ϴ� ���� ��ġ, ���� ȣ��
    Call color_set
    Call range_set
    
    '---function : ī�װ��� �������� ������ ���� ����
    If category_check = 1 Then
    
        Range("notice") = "ī�װ� ����Ʈ�� �������� �ʽ��ϴ�."
        Range("notice").Font.Color = vbRed
        
        category_add = 1
        Exit Function
        
    End If
       
    Range("notice") = ""
    
    category_count = 0
    
    If search_category_start <> Empty Then
        
        '---sub : ���� ���� Ȯ�� �� �ʱ�ȭ
        Call search_reset
            
        '---�˻�, ī�װ� ���� �ʱ�ȭ
        Range("DATA").Clear
        
    End If
    
    '---���õ� ī�װ� üũ
    For i = 1 To act_category_list.Rows.Count
        
        '---��ȸ���� �� ����
        Set now_cell = act_category_start.Offset(i - 1, 0)
        
        '---�������� ���� ���� üũ
        If Not now_cell.Interior.Color = vbWhite Then
            
            '---���õ� �� ������ �� �� �� �߰�
            search_category_start.Offset(0, category_count).Value = now_cell.Value
            
            '---���õ� �� ���� �� ����
            search_user_start.Offset(0, category_count).Interior.Color = user_input_color
            
            '---�� ���� üũ
            category_count = category_count + 1
            
            '---select_category ������ ���õ� �� �ּ� ����
            If select_category Is Nothing Then
            
                Set select_category = now_cell
                
            Else
            
                Set select_category = Union(select_category, now_cell)
                
            End If
        End If
    Next
    
    '---���õ� ī�װ��� ���� ���
    If category_count = 0 Then
    
        Range("notice") = "���õ� ī�װ��� �����ϴ�."
        Range("notice").Font.Color = vbRed
        
        category_add = 1
        
    Else
    
        '---ī�װ� �߰��� �� �ʺ� ����
        Range(search_category_start, search_category_start.Offset(0, category_count - 1)).EntireColumn.AutoFit
        
        '---�˻� ���� �ؽ�Ʈ Ÿ������ ����
        Range(search_user_start, search_user_start.Offset(0, category_count - 1)).NumberFormatLocal = "@"
        
        category_add = 0
        
    End If
    
    With Sheets("etc")

        '---etc ��Ʈ������ ������ ��ġ ã��
        search_row = .Range("preset_list").Find(what:=act_sheet_name.Value, lookat:=xlWhole).Row

        '---���õ� �� �ּ� �����¿� �����ϱ�
        If Not select_category Is Nothing Then

            .Cells(search_row, 6) = select_category.Address

        Else

            .Cells(search_row, 6).Clear

        End If

    End With

End Function

'=====================================================================
'��ũ�� : category_reset
'��� ��Ʈ : Home ��Ʈ
'���� : �� ���� ���� ���� ���� ��ư ����, �� ����Ʈ ��ü �� ���� ���ֱ�
'=====================================================================
Public Sub category_reset()

    '---�������� ����ϴ� ���� ��ġ ȣ��
    Call range_set
    
    '---function : ǥ�õ� �� �������� ������ ���� ����
    If category_check = 1 Then
    
        Range("notice") = "ī�װ� ����Ʈ�� �������� �ʽ��ϴ�."
        Range("notice").Font.Color = vbRed
        
        Exit Sub
    End If
    
    '---ī�װ� ���� ���� �ʱ�ȭ
    act_category_list.Interior.Color = vbWhite
    
End Sub

'=====================================================================
'��ũ�� : category_all_select
'��� ��Ʈ : Home ��Ʈ
'���� : �� ���� ���� ��ü ���� ��ư ����, �� ����Ʈ ��ü �� ���� ����
'=====================================================================
Public Function category_all_select()
    
    'category_all_select = 1 : ǥ�õ� ���� ����
    
    '---�������� ����ϴ� ���� ��ġ, ���� ȣ��
    Call color_set
    Call range_set
    
    '---function : ǥ�õ� ���� �������� ������ ���� ����
    If category_check = 1 Then
    
        Range("notice") = "ī�װ� ����Ʈ�� �������� �ʽ��ϴ�."
        Range("notice").Font.Color = vbRed
        category_all_select = 1
        Exit Function
        
    End If
    
    '---�� ��ü ���� �� ���� ����
    act_category_list.Interior.Color = category_sel_color
    category_all_select = 0

End Function

'=====================================================================
'��ũ�� : search_reset
'��� ��Ʈ : Home ��Ʈ
'���� : �˻� ���� �˻� �ʱ�ȭ ��ư, �˻����̴� ���� �ʱ�ȭ �� ������ ��Ʈ ���� ����
'=====================================================================
Public Function search_reset()
    
    'search_reset = 1 : ���õ� ���� ����
    'search_reset = 0 : ���� ����
        
    '---�������� ����ϴ� ���� ��ġ ȣ��
    Call range_set
    
    '---�˻��� �ʱ�ȭ
    Range(search_user_start, search_category_start.End(xlToRight).Offset(-1, 0)).ClearContents
    
    '---���õ� �� �� ���� �ʱ�ȭ
    Range(search_category_start, search_category_start.End(xlToRight)).ClearFormats
    
    '---���õ� ���� �������� ������ ���� ����
    If search_category_start = "" Then
    
        Range("notice") = "���õ� ī�װ��� �������� �ʽ��ϴ�."
        Range("notice").Font.Color = vbRed
        
        search_reset = 1
        
        Exit Function
        
    End If
    
    '---������ ��Ʈ ������ ���Ͱ� �ɷ��ִٸ� ����
    If Sheets(CStr(act_sheet_name.Value)).AutoFilter.FilterMode = True Then

        Sheets(act_sheet_name.Value).ShowAllData

    End If
    
    search_reset = 0
    
End Function


Public Sub EditCategoryRow()
    
    Dim varCategoryRow As Variant
    Dim search_row As Variant
    
    Call update_start
    Call range_set
    
    '---��� �Է� �ڽ� ǥ��
    varCategoryRow = InputBox("�Ӹ��۷� ����� ���� �Է����ּ���. (������ �Է�)", "�� ����")
    
    '---�Է��� ���� ���� ��� ����
    If varCategoryRow = Empty Then
        
        Exit Sub
        
    End If
    
    '--- ������ �ƴ� ��� �˸� ǥ��
    If varCategoryRow < 1 Or IsNumeric(varCategoryRow) = False Then
        
        MsgBox "1 �̻��� ������ �Է����ּ���."
        Exit Sub
        
    End If
        
    '---���� �����¸� etc ��Ʈ���� ��ġ �˻�
    search_row = Range("preset_list").Find(what:=act_sheet_name, lookat:=xlWhole).Row
    
    '---etc ������ ����Ʈ �������� �Ӹ��� �� ������ �߰�
    Range("preset_list")(search_row).Offset(0, 5).Value = varCategoryRow
    
    Call search_list(act_sheet_name)
    
    Call update_end
    
End Sub

Sub test()

    Dim data_range As ListObject
    
    Set data_range = Sheets("������2").ListObjects(1)
    
    MsgBox data_range.ListRows.Count
    data_range.HeaderRowRange.Select
End Sub
