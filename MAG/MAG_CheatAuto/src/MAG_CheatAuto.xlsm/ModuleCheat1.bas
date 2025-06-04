Attribute VB_Name = "ModuleCheat1"
Option Explicit
'###################################################
'Cheat1 RequestCreateItem 치트키 관련 모듈
'###################################################

'======================================================================================================
'치트키1 [Cheat 생성] 버튼 클릭 시 동작


Public Sub Cheat1()
    
    '# 변수 선언 #
    Dim strFileName As Variant '---조회해야할 문서명 저장 배열
    Dim rngFindCell As Range '---key로 조회된 셀 위치 저장
    Dim rngRune As Range '---룬 문서에서 조회된 셀 위치 저장
    
    '# 동작 시작 #
    Call UpdateStart
    Call SetRange
    
    '선택된 키가 없을 때 종료
    If 검색목록_시작.Value = "" Then
    
        MsgBox "선택된 Key가 없습니다."
        
        GoTo 종료
        
    End If
    
    '조회할 문서명 배열에 저장
    With 타입.ListColumns("문서").DataBodyRange
    
        ReDim strFileName(1 To .Cells.Count)
        
        For i = 1 To UBound(strFileName)
        
            strFileName(i) = 타입.ListColumns("문서").DataBodyRange(i).Value
            
        Next
        
    End With
    
    '선택된 KEY 개수만큼 반복
    For Each cell In 검색목록
        
        If cell.Offset(0, 2) = "" Then
            '문서 개수만큼 반복
            For i = 1 To UBound(strFileName)
                
                '룬 조회
                If strFileName(i) = "RuneUIData" Then
                    
                    Set rngFindCell = Sheets(strFileName(i)).UsedRange.Find(cell.Value, Lookat:=xlWhole)
                    
                    If Not rngFindCell Is Nothing Then
                    
                        Set rngRune = Sheets("RuneData").UsedRange.Find(rngFindCell.Offset(0, -1).Value, Lookat:=xlWhole)
                        
                        If Not rngRune Is Nothing Then
                        
                            cell.Offset(0, 2).Value = rngRune.Offset(0, 1).Value '---아이템 TID 입력
                            
                            cell.Offset(0, 3).Value = strFileName(i) '---아이템 타입에 문서명 입력
                            
                            GoTo 다음셀
                        
                        End If
                    End If
                Else
                    
                    Set rngFindCell = Sheets(strFileName(i)).UsedRange.Find(cell.Value, Lookat:=xlWhole) '---key 값으로 각 문서 조회
                    
                    '조회되었을 때 동작
                    If Not rngFindCell Is Nothing Then
                    
                        cell.Offset(0, 2).Value = rngFindCell.Offset(0, -1).Value '---아이템 TID 입력
                        
                        cell.Offset(0, 3).Value = strFileName(i) '---아이템 타입에 문서명 입력
                        
                        GoTo 다음셀
                        
                    End If
                End If
            Next
        End If
        
다음셀:
    Next
    
    '치트키 생성
    Call CheatCreatItem
    
    치트키_시작.Offset(-1, 0).Value = "일괄 입력 희망 시 [메모장 입력] 버튼을 클릭해주세요." '---상단에 안내 문구 표시
    
종료:

    Call UpdateEnd
    
End Sub

'======================================================================================================
'치트키 생성


Public Sub CheatCreatItem()
    
    '# 변수 선언 #
    
    Dim InItemType As Variant '---아이탬 타입 저장 변수
    Dim InTemplateId As Variant '---아이템 TID 저장 변수
    Dim InCount As Variant '---아이템 개수 저장 변수
    Dim InLevel As Variant '---아이템 레벨 저장 변수
        
        
    '# 동작 시작 #
    
    '고정 영역 호출
    Call SetRange
    
    치트키.ClearContents '---치트키 영역 초기화
    
    '선택된 KEY 개수만큼 반복
    For i = 0 To 검색목록.Cells.Count - 1
        
        With 검색목록(i + 1)
            
            InTemplateId = .Offset(0, 2).Value '---TID 값 저장
            
            '문서에 따라 아이템 타입 저장
            '무기, 외장부품, 반응로
            If .Offset(0, 3).Value = "RangedWeaponData" Or .Offset(0, 2).Value = "AccessoryData" Or .Offset(0, 2).Value = "ReactorData" Then
            
                InItemType = 2
            
            '재료
            ElseIf .Offset(0, 3).Value = "ConsumableItemData" Then
            
                InItemType = 3
            
            '룬
            ElseIf .Offset(0, 3).Value = "RuneUIData" Then
            
                InItemType = 4
             
            '커스터마이징
            ElseIf .Offset(0, 3).Value = "CustomizingItemData" Then
            
                InItemType = 7
            
            '아르케 조율 아이템
            ElseIf .Offset(0, 3).Value = "TuningBoardJewelData" Then
                
                InItemType = 14
                
            End If
            
            InCount = .Offset(0, 4).Value '---아이템 수량 저장
            
            '공백 시 1개 입력
            If InCount = 0 Then
            
                InCount = 1
                
            End If
                        
            InLevel = .Offset(0, 5).Value '---아이템 레벨 설정
            
            '공백 시 레벨 100 입력
            If InLevel = 0 Then
            
                InLevel = 100
                
            End If
        
        End With
        
        '아이템 ID 공백 시 안내 문구 표시
        If InTemplateId = 0 Then
        
            치트키_시작.Offset(i, 0).Value = "조회된 TID가 존재하지 않습니다."
        
        '치트키 입력
        Else
        
            치트키_시작.Offset(i, 0).Value = "RequestCreateItem " & InItemType & " " & InTemplateId & " " & _
                                        InCount & " " & InLevel
        End If
    
    Next
    
    '현재 프리셋 리스트 표시
    Call LoadTxt
    
End Sub
