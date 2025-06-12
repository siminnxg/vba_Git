VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ItemDropRate"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8385.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'======================================================================================================
'����� ���� [�⺻ ��] ��ư Ŭ�� �� ����

Private Sub Button_BaseRate_Click()
    
    Me.�����1.Value = "0.81"
    Me.�����2.Value = "0.11"
    Me.�����3.Value = "0.0081"
    Me.�����4.Value = "0.00081"
    Me.�����5.Value = "0.0000331"
    Me.�����6.Value = "1"
    
End Sub

'======================================================================================================
'���� ����Ʈ [���� ����] ��ư Ŭ�� �� ����

Private Sub Button_FileLoad_Click()
    
    Dim strFileName As String '---������ ������ ��� �� �̸� ���� ����
    
    '���� Ž���� ����
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "��������", "*.xls; *.xlsx; *.xlsm" '---���� �������� ����
        .Show
        
        '���� �� ���� �� ���� ó��
        If .SelectedItems.Count = 0 Then
        
            MsgBox "������ �������� �ʾҽ��ϴ�."
            
            Exit Sub
            
        End If
        
        'etc ��Ʈ '�����̸�' ���� �ʱ�ȭ
        Range("�����̸�").ClearContents
        ThisWorkbook.Names("�����̸�").RefersTo = Sheets("etc").Range("B3")
                        
        '���� ����Ʈ�� ���õ� ���� ��� �Է�
        For i = 1 To .SelectedItems.Count

            Me.List_File.AddItem .SelectedItems(i)

        Next
        
        'etc ��Ʈ '�����̸�' ������ ��θ� ������ ���� �̸��� �Է�
        For i = 0 To Me.List_File.ListCount - 1
            
            strFileName = InStrRev(Me.List_File.List(i), "\") '---���� ��ο� �и�
            
            Range("�����̸�").Offset(i, 0).Value = Mid(Me.List_File.List(i), strFileName + 1) '---���� �̸� �Է�
            
        Next
        
        '�����̸�' ���� ������
        ThisWorkbook.Names("�����̸�").RefersTo = Range("�����̸�").Resize(i + 1, 1)
              
    End With
    
End Sub

'======================================================================================================
'���� ����Ʈ [�ʱ�ȭ] ��ư Ŭ�� �� ����

Private Sub CommandButton2_Click()
    
    Me.List_File.Clear
    
End Sub

'======================================================================================================
'[��ȸ] ��ư Ŭ�� �� ����

Private Sub CommandButton1_Click()
    
    Dim cntResult As Variant
    
    Dim objDB As Object
    Dim obj As Object
    Dim strSql As String
    
    Dim strLevel As String
    Dim varDropRate As Variant
    Dim path As String
        
    Dim ��� As Range '---������� �ʰ��� �����͸� ����� ��ġ
    Dim ���2 As Range '---������ ����� ��ġ���� �ʴ� ����Ʈ�� ����� ��ġ
    Dim ���ϸ� As Range
    
    '#���� ����#
    
    '���� �߻� �� ����
    On Error Resume Next
    
    k = 1
    
    '���� ���� ���� Ȯ��
    If IsNull(Me.List_File.List(0)) Then
    
        MsgBox "���õ� ������ �����ϴ�."
        Exit Sub
        
    End If
    
    '����� ���� �� ���� ���� Ȯ��
    For i = 1 To 6
        
        Set control = Me.Controls("�����" & i)
        
        If control = "" Then
            
            MsgBox "������� �Է����ּ���."
            Exit Sub
            
        ElseIf IsNumeric(control) = False Then
            
            MsgBox "������� ���� �������� �Է����ּ���."
            Exit Sub
            
        End If
        
    Next
    
    '����� ���� ���� ���� Ȯ��
    If IsNumeric(Me.���������) = False Then
        
        MsgBox "����� ���� ��ġ�� ���� �������� �Է����ּ���."
        Exit Sub
            
    End If
    
    'ȭ�� ������Ʈ ����
    Call UpdateStart
    
    'Main ��Ʈ �ʱ�ȭ
    Call ClearMain
    
    Set ��� = Sheets("Main").Range("B4") '---�ٿ� ���� ���� ����
        
    Set ���2 = Sheets("��޿���").Range("B2")
    
    Range("���").Copy Destination:=���2.Offset(-1, 0) '--�Ӹ��� �� �߰�
    
    For i = 0 To Me.List_File.ListCount - 1
        
        cntResult = 0
                
        path = Me.List_File.List(i) '---���� ����Ʈ ���������� ��� ����
        
        Set ���ϸ� = ���.Offset(-1, -1)
        
        ���ϸ� = Range("�����̸�")(i + 1) '---��ȸ�� ���ϸ� �Է�
        
        ���2.Offset(0, -1).Value = Range("�����̸�")(i + 1)
        
        'OLEDB ����
        Set objDB = CreateObject("ADODB.Connection")
        objDB.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                  "Data Source=" & path & ";" & _
                  "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
                
        '��� �� ��ŭ �ݺ� (6��)
        For j = 1 To Range("���").Cells.Count
            
            strLevel = Range("���")(j).Value '---��� ������ ����
            
            Set control = Me.Controls("�����" & j) '---����� �ؽ�Ʈ �ڽ� ���� ����
            
            varDropRate = control.Value * 0.01 '---����� ������ ����
            
            '(������ Ÿ��, ��� ��ġ, ������� �ʰ�) ������ �����ϴ� ������ ��ȸ
            strSql = " SELECT * FROM [Data$] WHERE F10 = '������' AND F8 LIKE '%" & strLevel & "%' AND F9 = '" & strLevel & "' AND F12 LIKE '%Grade" & j & "%' AND F17 > " & varDropRate
            
            Set obj = objDB.Execute(strSql)
            
            '��ȸ�� �����Ͱ� ������ ���� ������� �̵�
            If obj.EOF Then
                
                GoTo �����ȸ
            
            '��ȸ�� �����Ͱ� �ִ� ��� Main ��Ʈ�� �Է�
            Else
                
                ���.CopyFromRecordset obj '---��ȸ�� ������ �Է�
                
                Range("���").Copy Destination:=���.Offset(-1, 0) '---�Ӹ��� �Է�
                
                '����� ���� Ȯ��
                Call CheckRate(���.CurrentRegion, varDropRate)
                
                Set ��� = ���.End(xlDown).Offset(3, 0) '---��� ���� ������
                
                cntResult = 1 '---��ȸ�� ������ ���� Ȯ��
                
                Application.CutCopyMode = False
                
            End If
            
�����ȸ:
            '������ Ÿ�� �� ����� ��ġ���� �ʴ� ���� ��ȸ
            strSql = " SELECT * FROM [Data$] WHERE F10 = '������' AND F8 LIKE '%" & strLevel & "%' AND (F9 <> '" & strLevel & "' OR F12 NOT LIKE '%Grade" & j & "%')"
            
            Set obj = objDB.Execute(strSql)
            
            '��ȸ�� �����Ͱ� ������ ���� ������� �̵�
            If obj.EOF Then
                
                GoTo �������
            
            '��ȸ�� �����Ͱ� �ִ� ��� ��޿��� ��Ʈ�� �Է�
            Else
                
                ���2.CopyFromRecordset obj '---��ȸ�� ������ �Է�
                
                Set ���2 = ���2.End(xlDown).Offset(1, 0)
                
                Application.CutCopyMode = False
                
            End If
                        
�������:

        Next

        '��ü ���� ����
        obj.Close
        objDB.Close
        Set obj = Nothing
        Set objDB = Nothing
        
        '��ȸ�� �����Ͱ� ������ ���ϸ� ����
        If cntResult = 0 Then
        
            ���ϸ�.ClearContents
            
        End If
        
        j = j + 1
        
    Next
        
    Columns("R:R").NumberFormatLocal = "0.000000%" '---����� ���� ����
    
    Sheets("Main").UsedRange.Columns.AutoFit '---��� �ڵ� �� ����
    
    '��ȸ ���� �� �˸�
    If Range("B5") = "" Then
        
        MsgBox "��ȸ�� �����Ͱ� �����ϴ�."
        
    Else
        
        MsgBox "��ȸ �Ϸ�Ǿ����ϴ�."
        
    End If

����:
    Call UpdateEnd
    
End Sub

'======================================================================================================
'����� ���� Ȯ��

Public Function CheckRate(rngResult As Range, varDropRate As Variant)
    
    Dim varRate As Variant '---����� ���� ��ġ ���� ����
    
    '���� ��ȸ�� ������ �׵θ� ����
    With rngResult.Borders
        
        .LineStyle = xlContinuous
        .Color = rgbGainsboro
        
    End With
    
    '��ġ �ԷµǾ� �ִ� ��� ����
    If Me.���������.Value <> "" Then
        
        varRate = Me.���������.Value * 0.01 '---����� ���� ��ġ ������ �Է�
        
        '��ȸ ��� ������ ��ȸ
        For Each cell In rngResult
                        
            '����� ǥ�� ��, ����� ���� ��ġ + ����� ��ġ ���� ū ��� �� ���� ����
            If cell.Column = 18 And cell.Value > (varRate + varDropRate) Then
            
                Range(Cells(cell.Row, 2), Cells(cell.Row, cell.Column)).Interior.ColorIndex = 6
                
            End If
                    
        Next
        
    End If

End Function
