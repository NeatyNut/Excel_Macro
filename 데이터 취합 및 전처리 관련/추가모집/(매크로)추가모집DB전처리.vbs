Sub 분리()
Dim row As Integer
Dim column As Integer
Dim rownum As Integer

ThisWorkbook.Sheets("정리").Cells.Select
Selection.ClearContents



row = ThisWorkbook.Sheets("내용").Range("A1").CurrentRegion.Rows.Count
column = ThisWorkbook.Sheets("내용").Range("A1").CurrentRegion.Columns.Count
rownum = 2
            
ThisWorkbook.Sheets("정리").Range("A1") = "대학교"
ThisWorkbook.Sheets("정리").Range("B1") = "캠퍼스"
ThisWorkbook.Sheets("정리").Range("C1") = "설립구분"
ThisWorkbook.Sheets("정리").Range("D1") = "내외"
ThisWorkbook.Sheets("정리").Range("E1") = "모집시기"
ThisWorkbook.Sheets("정리").Range("F1") = "전형유형"
ThisWorkbook.Sheets("정리").Range("G1") = "전형명(대)"
ThisWorkbook.Sheets("정리").Range("H1") = "전형명(중)"
ThisWorkbook.Sheets("정리").Range("I1") = "전형명(소)"
ThisWorkbook.Sheets("정리").Range("J1") = "모집단위"
ThisWorkbook.Sheets("정리").Range("K1") = "인원"


For i = 3 To row Step 1
    If Right(ThisWorkbook.Sheets("내용").Cells(i, 1).Value, 2) = "소계" Or Right(ThisWorkbook.Sheets("내용").Cells(i, 1).Value, 2) = "합계" Then '소계, 합계 무시
    Else
        If IsNumeric(ThisWorkbook.Sheets("내용").Cells(i, column - 3).Value) Or IsNumeric(ThisWorkbook.Sheets("내용").Cells(i, column - 2).Value) Or ThisWorkbook.Sheets("내용").Cells(i, column - 1).Value = 0 Then '미지정이 있다면
            un_summary = WorksheetFunction.Sum(ThisWorkbook.Sheets("내용").Cells(i, column - 3), ThisWorkbook.Sheets("내용").Cells(i, column - 2))
            
            If un_summary <> ThisWorkbook.Sheets("내용").Cells(i, column - 1).Value Then  '배정된 인원 + 미지정이라면
                For ii = 1 To column - 7 Step 1
                    ThisWorkbook.Sheets("정리").Cells(rownum, ii) = ThisWorkbook.Sheets("내용").Cells(i, ii).Value
                Next ii
                
                hackwa = Split(ThisWorkbook.Sheets("내용").Cells(i, column - 6), ",")
                addnum = 1
                
                For a = LBound(hackwa) To UBound(hackwa) Step 1
                    If IsNumeric(Right(hackwa(a), 1)) Then ':0 구조라면
                        If Mid(hackwa(a), WorksheetFunction.Find(":", hackwa(a)) + 1, Len(hackwa(a)) - WorksheetFunction.Find(":", hackwa(a))) <> 0 Then ':1 이상이면
                            For ii = 1 To column - 7 Step 1
                                ThisWorkbook.Sheets("정리").Cells(rownum + addnum, ii) = ThisWorkbook.Sheets("내용").Cells(i, ii).Value
                            Next ii
                            
                            inwon = Split(hackwa(a), ":")
                            ThisWorkbook.Sheets("정리").Cells(rownum + addnum, column - 6) = inwon(LBound(inwon))
                            ThisWorkbook.Sheets("정리").Cells(rownum + addnum, column - 5) = inwon(UBound(inwon))
                            
                            addnum = addnum + 1
                        Else ':0이라면
                            If IsEmpty(ThisWorkbook.Sheets("정리").Cells(rownum, column - 6)) Then
                                ThisWorkbook.Sheets("정리").Cells(rownum, column - 6) = hackwa(a)
                            Else
                                ThisWorkbook.Sheets("정리").Cells(rownum, column - 6) = ThisWorkbook.Sheets("정리").Cells(rownum, column - 6).Value & "," & hackwa(a)
                            End If
                        End If
                        
                    Else '학과명이 짤린거라면
                        hackwa(a + 1) = hackwa(a) & "," & hackwa(a + 1)
                    End If
                Next a
                
                ThisWorkbook.Sheets("정리").Cells(rownum, column - 6) = Replace(ThisWorkbook.Sheets("정리").Cells(rownum, column - 6), ":0", "")
                ThisWorkbook.Sheets("정리").Cells(rownum, column - 5) = un_summary
                ThisWorkbook.Sheets("정리").Cells(rownum, column - 4) = "통합"
                rownum = rownum + addnum
            Else '그냥 미지정이라면
                For ii = 1 To column - 7 Step 1
                    ThisWorkbook.Sheets("정리").Cells(rownum, ii) = ThisWorkbook.Sheets("내용").Cells(i, ii).Value
                Next ii
                
                ThisWorkbook.Sheets("정리").Cells(rownum, column - 6) = Replace(ThisWorkbook.Sheets("내용").Cells(i, column - 6), ":0", "")
                ThisWorkbook.Sheets("정리").Cells(rownum, column - 5) = un_summary
                ThisWorkbook.Sheets("정리").Cells(rownum, column - 4) = "통합"
                rownum = rownum + 1
            End If
            
            
        Else '미지정이 없다면
            hackwa = Split(ThisWorkbook.Sheets("내용").Cells(i, column - 6), ",")
            
            For a = LBound(hackwa) To UBound(hackwa) Step 1
                If IsNumeric(Right(hackwa(a), 1)) Then ':0 구조라면
                    For ii = 1 To column - 7 Step 1
                        ThisWorkbook.Sheets("정리").Cells(rownum, ii) = ThisWorkbook.Sheets("내용").Cells(i, ii).Value
                    Next ii
                    
                    inwon = Split(hackwa(a), ":")
                    ThisWorkbook.Sheets("정리").Cells(rownum, column - 6) = inwon(LBound(inwon))
                    ThisWorkbook.Sheets("정리").Cells(rownum, column - 5) = inwon(UBound(inwon))
                    
                    rownum = rownum + 1
                Else
                    hackwa(a + 1) = hackwa(a) & "," & hackwa(a + 1)
                End If
            Next a
        End If
    End If
Next i


End Sub
