VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "TOOL"
   ClientHeight    =   2685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5895
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'##########################################################
'������ ���� ��� Ȯ�� �� �����ϱ�
'##########################################################


Private Sub Btn_FileAdr_Click()
    
    If MsgBox("���� ������ ��δ� '" & ThisWorkbook.CustomDocumentProperties("�������").Value & "' �Դϴ�." & vbCrLf & vbCrLf & _
        "��θ� �����Ͻðڽ��ϱ�?", vbQuestion + vbOKCancel) _
        = 1 Then
    
        Call SearchFolder
        Unload Me
    
    End If
    
End Sub


'##########################################################
'Result ����, ��� ǥ�� ������ �Լ� ��ü �Է�
'##########################################################


Private Sub Btn_ShtRefresh_Click()
    
    Dim row As Variant
    
    Call Unprotect
    
    row = Range("������Է�")(1).row
    
    Range("���").Formula2 = "=���Ȯ��($B" & row & ":$E" & row & ",$H" & row & "#)"
    
    Range("���ǥ��").Formula2 = "=FILTER(GachaElement,(GachaElement[[long_2]]=ID����($C" & row & "))*(GachaElement[[long_1]]=Main!$B" & row & "), """")"
    
    Sheets("Main").Protect
    
    MsgBox "���� ���ΰ�ħ�� �Ϸ�Ǿ����ϴ�."
    
    Unload Me
    
End Sub


'##########################################################
'�޸� ���� �� �߰�
'##########################################################

Private Sub Btn_AddRow_Click()
    
    Call Unprotect
    
    With Sheets("Main")
        .Rows(Range("MEMO").row).Insert
        .Protect
    End With
    
End Sub


'##########################################################
'�޸� ���� �� ����
'##########################################################


Private Sub Btn_DelRow_Click()
    
    Dim row As Variant
    
    row = Range("MEMO").row
    
    If row < 6 Then
        MsgBox "������ �޸� ������ �����մϴ�."
        Exit Sub
    End If
    
    Call Unprotect
        
    With Sheets("Main")
        .Rows(row - 1).Delete
        .Protect
    End With
    
End Sub
