VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchForm 
   Caption         =   "��� �˻�"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10185
   OleObjectBlob   =   "SearchForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '������ ���
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
    
    For i = 0 To Range(�˻�Ű����_����, �˻�Ű����_��).Cells.Count
        Me.List_Category.AddItem �˻�Ű����_����.Offset(0, i).Value
    Next
    
End Sub

'=====================================================================
'�ڵ� ä���
Private Sub Button_AutoFill_Click()
    
    On Error Resume Next
    
    '���� ���� üũ
    If IsNull(Me.List_Category.Value) Then
        MsgBox "���õ� ���� �����ϴ�."
        Exit Sub
    End If
    
    Call SetRange
    Call AutoFill(Me.List_Category.Value)
    
End Sub

Private Sub Button_Refresh_Click()
    
    On Error Resume Next
    
    Call SetRange
    
    '������ �̸����� ������ ��Ʈ�� listobject ���ΰ�ħ
    Sheets(����������.Value).ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
    
    �˻���_����.Value = ""
    
End Sub


Private Sub Button_UniqueLoad_Click()
    
    Dim varSelCol As Variant
    Dim objData As Object
    Dim varUnique As Variant
    
    On Error Resume Next
    
    '���� ���� üũ
    If IsNull(Me.List_Category.Value) Then
        MsgBox "���õ� ���� �����ϴ�."
        Exit Sub
    End If
        
    Call SetRange
    
    Set objData = Sheets(����������.Value).ListObjects(1)
    varSelCol = objData.ListColumns(Me.List_Category.Value).Index
    
    varUnique = Application.WorksheetFunction.Unique(Sheets(����������.Value).Columns(varSelCol)) '---���õ� ���� ������ ����
    
    With Me.List_Search
    .List = varUnique '---����Ʈ �ڽ��� ǥ��
    .List(0) = Me.List_Category.ListIndex & " : " & .List(0)
    End With
End Sub

'=====================================================================
'����Ʈ �ڽ� ���� �ʱ�ȭ
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
    
    '���� �� Ȯ��
    With Me.List_Search
    
        varSelCol = Split(.List(0), " ") '---ù��° �� �������� �����Ͽ� �� ���� ã��
              
        For i = 1 To .ListCount - 1
            If .Selected(i) = True Then
                
                '���� �� �� �������� ����
                If strSearch = vbNullString Then
                
                    strSearch = .List(i)
                    
                Else
                
                    strSearch = strSearch & "," & .List(i)
                    
                End If
            End If
        Next
    End With
    
    '���� �� ���� ��� ó��
    If strSearch = vbNullString Then
        MsgBox "���õ� ���� �����ϴ�."
        Exit Sub
    End If
    
    �˻���_����.Offset(0, varSelCol(0)) = strSearch
    
End Sub

