Attribute VB_Name = "ModuleCommon"
Option Explicit

'###���� ����###

'���� ����
Public ���ϰ�� As Range '---���� ��� ����
Public ���ϸ�, ���ϸ�1 As Range '---���� �̸� ����
Public ��Ʈ�� As Range '---���� ��Ʈ ����
Public �˻��� As Range '---�˻��� �Է� ����
Public �˻���� As Range '---�˻��� ��� ����
Public �Ӹ��� As Range
Public ������Ʈ As Range

'���� ����
Public i, j, k As Variant '---�ݺ��� ��� ����
Public dateTime As Date '---�ð� üũ�� ����
Public rngTemp As Range '---���� ���� ���� ����

'=====================================================================
'���� ����
Public Sub SetRange()
    
    With Sheets("Main")
                
        '���� �̸� ���� ����
        Set ���ϸ� = .Range("C7")
        
        '���� �̸� ���� �Է� �� ó��
        If ���ϸ�.Offset(1, 0) <> "" Then
            
            Set ���ϸ� = Range(���ϸ�, ���ϸ�.Offset(-1, 0).End(xlDown))
            
        End If
        
        Set ���ϰ�� = ���ϸ�.Offset(0, -1) '---���� ��� ���� ����
        
        Set ��Ʈ�� = ���ϸ�.Offset(0, 1) '---��Ʈ�� ���� ����
        
        Set �Ӹ��� = ���ϸ�.Offset(0, 2) '--- �Ӹ��� �� ���� ����
    
        Set �˻��� = .Range("B21") '---�˻� �� ���� ����
        
        Set �˻���� = .Range("B24") '---�˻� ��� ǥ�� ���� ����
        
    End With
    
    With Sheets("etc")
        
        Set ������Ʈ = .Range("A1") '---ȣ��� ���� ����Ʈ ���� ��ġ
        
    End With
    
End Sub

'=====================================================================
'ȭ�� ������Ʈ ���� (���� �ӵ� ����)
Public Sub UpdateStart()

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
End Sub

'=====================================================================
'ȭ�� ������Ʈ ����
Public Sub UpdateEnd()
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub

'=====================================================================
'etc ��Ʈ '������Ʈ' ������ ���� ��� ����
Public Sub ObjectList(strFile)
    
    Dim varObjCnt As Variant '---������Ʈ ���� üũ
    
    If ������Ʈ = "" Then
        Range("������Ʈ")(i) = strFile
    
    Else
        varObjCnt = Application.WorksheetFunction.CountIf(Range("������Ʈ"), strFile) '---�ߺ� üũ
        
        '�ߺ��� ������Ʈ�� ������ �߰�
        If varObjCnt = 0 Then
        
            ������Ʈ.Offset(Range("������Ʈ").count, 0) = strFile
            ThisWorkbook.Names("������Ʈ").RefersTo = Range(Range("������Ʈ"), ������Ʈ.End(xlDown)) '---���� ������ 2�� �̻��� ��� '������Ʈ' ���� ������
                        
        End If
        
    End If
    
    For k = 1 To Range("������Ʈ").count
        
    Next
End Sub

Sub test()


End Sub


