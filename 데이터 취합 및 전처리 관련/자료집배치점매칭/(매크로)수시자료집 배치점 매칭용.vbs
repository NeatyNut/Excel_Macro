Sub find()
Dim find_h As String
Dim find_hrow_h As Integer
Dim find_hcol_h As Integer
Dim find_hrow_d As Integer
Dim find_h_col_d As Integer
Dim tbbd_row As Integer
Dim tbbd_col As Integer
Dim mcount As Integer
Dim Ucount As Integer
Dim nic As String
Dim ch_cnt As Integer
Dim ch_wd As String
Dim d_count As Integer


For aa = 1 To 33 Step 1

    If Columns(aa).Hidden = True Then
        Columns(aa).Hidden = False
    End If
Next aa
'찾는 키워드
find_h = "모집단위"

'위치값 찾기
For i = 1 To 50 Step 1
    For i2 = 1 To 50 Step 1
        If Cells(i, i2).Value = find_h Then: find_hrow_h = i: find_hcol_h = i2: Exit For
    Next i2
    If find_hcol_h <> 0 Then
        Exit For
    End If
Next i

If find_hcol_h = 0 Then
 MsgBox "인원표를 안붙인 듯 합니다."
 Exit Sub
End If


Ucount = 0

Do While (Cells(find_hrow_h, find_hcol_h + Ucount).Value <> "계열")
    Ucount = Ucount + 1
Loop

find_hrow_d = find_hrow_h + 1
Do While (Cells(find_hrow_d, find_hcol_h).Value = 0)
    find_hrow_d = find_hrow_d + 1
Loop

If Ucount > 1 Then
    For a4 = find_hrow_d To Cells(find_hrow_h, find_hcol_h).CurrentRegion.Rows.Count + find_hrow_h - 1
        '한열로 되어있을때
        If Cells(a4, find_hcol_h).MergeCells And Cells(a4, find_hcol_h + Ucount - 1).MergeCells Then
            Cells(a4, find_hcol_h + Ucount - 1).UnMerge
            nic = Cells(a4, find_hcol_h).Value
            
            For a5 = 1 To Ucount
                Cells(a4, find_hcol_h + a5 - 1).Value = nic
            Next a5
        Else
        End If
    Next a4
    
        
        For a3 = 1 To Ucount - 1
            Columns(find_hcol_h).Delete
            Columns("Z").Insert
        Next a3
    
    Cells(find_hrow_h, find_hcol_h).Value = "모집단위"
End If

'끝줄 찾기
find_hrow_d = find_hrow_h + 1

Do While (Cells(find_hrow_d, find_hcol_h).Value = 0)
    find_hrow_d = find_hrow_d + 1
Loop

'끝줄
find_hrow_d = find_hrow_d - 1

'끝열 찾기
find_hcol_d = find_hcol_h + 1

Do While (Cells(find_hrow_h, find_hcol_d).Interior.ColorIndex = 16)
    find_hcol_d = find_hcol_d + 1
Loop

'끝열
find_hcol_d = find_hcol_d - 1

'내용 줄 갯수
tbbd_row = Cells(find_hrow_h, find_hcol_h).CurrentRegion.Rows.Count - (find_hrow_d - find_hrow_h + 1)

'*************학과명 정리
For i3 = 1 To tbbd_row Step 1
    mcount = 1
    '학과명 합치기
    If Cells(find_hrow_d + i3, find_hcol_h + 1).MergeCells And Cells(find_hrow_d + i3, find_hcol_h).MergeCells <> True Then
        Do While (Cells(find_hrow_d + i3 + mcount, find_hcol_h + 1).Value = 0 And Cells(find_hrow_d + i3 + mcount, find_hcol_h).Value <> 0)
            mcount = mcount + 1
        Loop
        
        For i4 = 1 To mcount - 1 Step 1
            Cells(find_hrow_d + i3, find_hcol_h).Value = Cells(find_hrow_d + i3, find_hcol_h).Value & Cells(find_hrow_d + i3 + i4, find_hcol_h).Value
        Next i4
        
        Range(Rows(find_hrow_d + i3 + 1), Rows(find_hrow_d + i3 + mcount - 1)).Delete
        tbbd_row = tbbd_row - mcount + 1
    '의미없는 병합 없애기
    ElseIf Cells(find_hrow_d + i3, find_hcol_h + 1).MergeCells And Cells(find_hrow_d + i3, find_hcol_h).MergeCells Then
        Do While (Cells(find_hrow_d + i3 + mcount, find_hcol_h + 1).Value = 0 And Cells(find_hrow_d + i3 + mcount, find_hcol_h).Value = 0)
            If i3 + mcount - 1 <= tbbd_row Then
                mcount = mcount + 1
            Else
                Exit For
            End If
        Loop
        
        Range(Rows(find_hrow_d + i3 + 1), Rows(find_hrow_d + i3 + mcount - 1)).Delete
        tbbd_row = tbbd_row - mcount + 1
    '계열명 합치기
    ElseIf Cells(find_hrow_d + i3, find_hcol_h + 1).MergeCells <> True And Cells(find_hrow_d + i3, find_hcol_h).MergeCells Then
        Do While (Cells(find_hrow_d + i3 + mcount, find_hcol_h).Value = 0 And Cells(find_hrow_d + i3 + mcount, find_hcol_h + 1).Value <> 0)
            If i3 + mcount - 1 <= tbbd_row Then
                mcount = mcount + 1
            Else
                Exit For
            End If
        Loop
        
        For i4 = 1 To mcount - 1 Step 1
            Cells(find_hrow_d + i3, find_hcol_h + 1).Value = Cells(find_hrow_d + i3, find_hcol_h + 1).Value & Cells(find_hrow_d + i3 + i4, find_hcol_h + 1).Value
        Next i4
        
        Range(Rows(find_hrow_d + i3 + 1), Rows(find_hrow_d + i3 + mcount - 1)).Delete
        tbbd_row = tbbd_row - mcount + 1
        
    End If
    
Next i3
'새 내용 줄 갯수
tbbd_row = Cells(find_hrow_h, find_hcol_h).CurrentRegion.Rows.Count - (find_hrow_d - find_hrow_h + 1)

'빈칸(특정 키워드 가능) 찾아 지우기
Range(Cells(find_hrow_d + 1, find_hcol_h), Cells(find_hrow_d + tbbd_row, find_hcol_h)).Select

    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

ch_cnt = Sheets("찾아바꾸기").Range("a1").CurrentRegion.Rows.Count

For b1 = 1 To ch_cnt Step 1
    ch_wd = Sheets("찾아바꾸기").Cells(b1, 1).Value
    Sheets("붙이는시트").Range(Cells(find_hrow_d + 1, find_hcol_h), Cells(find_hrow_d + tbbd_row, find_hcol_h)).Select

    Selection.Replace What:=ch_wd, Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Next b1
    
    

Sheets("붙이는시트").Activate



'표 위치 정리하기
If find_hrow_h <> 1 Then
    Range(Cells(1, 1), Cells(find_hrow_h - 1, find_hcol_d)).Delete
    find_hrow_d = find_hrow_d - find_hrow_h + 1
    find_hrow_h = 1
End If

If find_hcol_h <> 1 Then
    For A1 = 1 To find_hcol_h - 1 Step 1
        Columns(find_hcol_d + 2).Insert
    Next A1
    Range(Columns(1), Columns(find_hcol_h - 1)).Delete
    Columns(find_hcol_d + 2).Delete
    find_hcol_d = find_hcol_d - find_hcol_h + 2
    find_hcol_h = 2
Else
    find_hcol_d = find_hcol_d + 1
    find_hcol_h = 2
End If

'매칭값 만들기
Columns(1).Insert

For i5 = 1 To tbbd_row Step 1
    If Cells(find_hrow_d + i5, 3).Value <> "예체능" Then
        Cells(find_hrow_d + i5, 1).Value = "=B" & find_hrow_d + i5 & "&C" & find_hrow_d + i5
        '영어등급이나 배치표에 대한 수정 필요
        Cells(find_hrow_d + i5, find_hcol_d - 1).Value = "=VLOOKUP(A" & find_hrow_d + i5 & ",OFFSET(대학별배치표!$A$1,0,0,COUNTA(대학별배치표!$A:$A),COUNTA(대학별배치표!$1:$1)),4,0)"
        Cells(find_hrow_d + i5, find_hcol_d).Value = "=VLOOKUP(A" & find_hrow_d + i5 & ",OFFSET(대학별배치표!$A$1,0,0,COUNTA(대학별배치표!$A:$A),COUNTA(대학별배치표!$1:$1)),5,0)"
'        Cells(find_hrow_d + i5, find_hcol_d + 1).Value = "=VLOOKUP(A" & find_hrow_d + i5 & ",OFFSET(대학별배치표!$A$1,0,0,COUNTA(대학별배치표!$A:$A),COUNTA(대학별배치표!$1:$1)),3,0)=C" & find_hrow_d + i5
    Else
        Cells(find_hrow_d + i5, find_hcol_d - 1).Value = "-"
        Cells(find_hrow_d + i5, find_hcol_d).Value = "-"
'        Cells(find_hrow_d + i5, find_hcol_d + 1).Value = "=VLOOKUP(A" & find_hrow_d + i5 & ",OFFSET(대학별배치표!$A$1,0,0,COUNTA(대학별배치표!$A:$A),COUNTA(대학별배치표!$1:$1)),3,0)=C" & find_hrow_d + i5
    End If

Next i5

'밑에 셀 합병 시 정리하기
If Cells(tbbd_row + find_hrow_d, find_hcol_h).MergeCells Then
    For b2 = find_hcol_h To find_hcol_d
        Cells(tbbd_row + find_hrow_d, b2).UnMerge
    Next b2
    
    Do While (Cells(tbbd_row + find_hrow_d, find_hcol_h).Value = "")
    Rows(tbbd_row + find_hrow_d).Delete
    tbbd_row = tbbd_row - 1
    Loop
End If

'열 접기
Range(Columns(find_hcol_h + 2), Columns(find_hcol_d - 2)).Hidden = True

End Sub
