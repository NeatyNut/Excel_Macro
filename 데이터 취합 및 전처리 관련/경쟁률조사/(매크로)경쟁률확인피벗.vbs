Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
    ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Sub 파일다운()

savefile = Environ("USERPROFILE") & "\Desktop\경쟁률원본.xls"
chkfile = Dir(savefile)
If (Len(chkfile) > 0) Then
Application.Workbooks.Open FileName:=savefile
ActiveWorkbook.Close
Kill (savefile)
Else
End If

myresult = URLDownloadToFile(0, "관련 URL", savefile, 0, 0)

If myresult <> 0 Then
    MsgBox "다운로드 안됨"
Else
    MsgBox "완료"
End If
End Sub

Sub 데이터가져오기()
  Dim Fileobj As Object '엑셀파일 Object를 생성하기 위한 개체 변수
  Dim filePath As String '불러올 파일 경로를 저장할 String 변수
  Dim sht As Worksheet 'Worksheet 설정을 위한 변수
  Dim col As Integer
  Dim row As Integer
  
Sheets("피벗테이블").Activate
Sheets("피벗테이블").Range("A:I").Select
Selection.Clear
Sheets("피벗테이블").Cells(1, 1).Select

Sheets("데이터").Activate
If ActiveSheet.AutoFilterMode Then ActiveSheet.UsedRange.AutoFilter
  Sheets("데이터").Cells.Select
  Selection.Clear

  '"Sample_Data.xlsx"의 파일경로 및 파일명 변수에 저장
  filePath = Environ("USERPROFILE") & "\Desktop\경쟁률원본.xls"
  
  'GetObject 함수를 통해 엑셀파일 개체 생성
  Set Fileobj = GetObject(filePath)
  
  '위 엑셀파일 개체의 활성화된 시트를 sht 변수에 저장
  Set sht = Fileobj.ActiveSheet
  
  'sht에 접근하기
  With sht
  
    .UsedRange.Copy '각 DATA 시트의 데이터를 모두 복사
    Sheets("데이터").Cells(1, 1).PasteSpecial xlPasteValues '현재 매크로 실행되는 시트에 붙여넣기
    
  End With
  
    col = Sheets("데이터").Cells(1, 1).CurrentRegion.Columns.Count
    row = Sheets("데이터").Cells(1, 1).CurrentRegion.Rows.Count
    
    If row <= 1 Or col <= 1 Then
        MsgBox "데이터가 없습니다"
        Sheets("피벗테이블").Activate
        Exit Sub
    End If
    
    Sheets("데이터").Cells(1, col + 1).Value = "모집인원2"
    Sheets("데이터").Cells(1, col + 2).Value = "지원인원2"

    Sheets("데이터").Cells(2, col + 1) = "=iferror(INT(RC[-7]),0)"
    Sheets("데이터").Cells(2, col + 2) = "=iferror(INT(RC[-7]),0)"
    Sheets("데이터").Range(Cells(2, col + 1), Cells(row, col + 2)).Select
    Selection.FillDown



   ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "데이터!R1C1:R" & row & "C" & col + 2, Version:=6).CreatePivotTable TableDestination:= _
        "피벗테이블!R1C1", TableName:="피벗 테이블6", DefaultVersion:=6
    
    Sheets("피벗테이블").Activate

    With ActiveSheet.PivotTables("피벗 테이블6").PivotFields("학교명")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("피벗 테이블6").PivotFields("내외")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("피벗 테이블6").PivotFields("전형유형")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("피벗 테이블6").PivotFields("전형명")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("피벗 테이블6").PivotFields("기준시간")
        .Orientation = xlRowField
        .Position = 4
    End With
    ActiveSheet.PivotTables("피벗 테이블6").AddDataField ActiveSheet.PivotTables( _
        "피벗 테이블6").PivotFields("모집인원2"), "합계 : 모집인원2", xlSum
    ActiveSheet.PivotTables("피벗 테이블6").AddDataField ActiveSheet.PivotTables( _
        "피벗 테이블6").PivotFields("지원인원2"), "합계 : 지원인원2", xlSum
    Range("A13").Select
    With ActiveSheet.PivotTables("피벗 테이블6")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    ActiveSheet.PivotTables("피벗 테이블6").PivotFields("전형명").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("피벗 테이블6").PivotFields("전형유형").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("피벗 테이블6").PivotFields("내외").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveWindow.SmallScroll Down:=-90
    ActiveSheet.PivotTables("피벗 테이블6").RepeatAllLabels xlRepeatLabels
    Columns("A:F").Select
    Columns("A:F").EntireColumn.AutoFit

MsgBox "완료하였습니다."
    
End Sub
