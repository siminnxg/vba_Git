Attribute VB_Name = "ModuleCommon"
Option Explicit

'###���� ����###

'���� ����
Public ���ϰ�� As Range '---���� ��� ����
Public ���ϸ�, ���ϸ�1 As Range '---���� �̸� ����
Public ��Ʈ�� As Range '---���� ��Ʈ ����
Public �˻��� As Range '---�˻��� �Է� ����
Public �˻���� As Range '---�˻��� ��� ����
Public ������Ʈ As Range

'���� ����
Public i, j, k As Variant '---�ݺ��� ��� ����
Public dateTime As Date '---�ð� üũ�� ����
Public rngTemp As Range '---���� ���� ���� ����

'=====================================================================
'���� ����
Public Sub SetRange()
    
    With Sheets("Main")
    
        Set ���ϰ�� = .Range("B5") '---���� ��� ���� ����
        
        '���� �̸� ���� ����
        Set ���ϸ� = .Range("B7")
        
        '���� �̸� ���� �Է� �� ó��
        If ���ϸ�.Offset(1, 0) <> "" Then
            
            Set ���ϸ� = Range(���ϸ�, ���ϸ�.Offset(-1, 0).End(xlDown))
            
        End If
        
        Set ��Ʈ�� = .Range("C7") '---��Ʈ�� ���� ����
    
        Set �˻��� = .Range("B19") '---�˻� �� ���� ����
        
        Set �˻���� = .Range("B22") '---�˻� ��� ǥ�� ���� ����
        
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
