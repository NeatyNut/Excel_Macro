Option Explicit
Function file_Exists(strRng As String) As Boolean

End Function
Function GetBook() As String
    GetBook = ActiveWorkbook.name
End Function

Sub 단추1_Click()
Dim work_name As String
Dim cor As String
Dim name As String
Dim datevalue As String
Dim datevalue2 As String
Dim SavePath As String
Dim password As String
Dim file_name As String
Dim strName As String
Dim row_num As Integer
Dim a As Integer
Dim i As Integer

work_name = GetBook()

SavePath = Environ("USERPROFILE") & "\Desktop\PDF암호화"

If Dir(SavePath, vbDirectory) = "" Then
    MkDir SavePath
End If


Workbooks(work_name).Sheets("임금명세서_4대보험").Activate

cor = Range("G3").Value
name = Range("I3").Value
datevalue = month(Range("A10").Value) & "." & day(Range("A10").Value)
datevalue2 = month(Range("A14").Value) & "." & day(Range("A14").Value)
file_name = SavePath & "\" & "(" & cor & ")임금명세서_" & name & "_" & datevalue & "~" & datevalue2 & ".pdf"


Application.ScreenUpdating = False
ActiveSheet.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    filename:=file_name, _
    OpenAfterPublish:=False

Application.ScreenUpdating = True


password = Chr(34) & Mid(Range("L3"), 1, 6) & Chr(34)
strName = SavePath & "\" & "PDF_암호.xlsx"
a = 0

If Dir(strName) <> vbNullString Then
    Workbooks.Open strName
    row_num = ActiveWorkbook.Sheets(1).Range("A1").CurrentRegion.Rows.Count
        
    For i = 1 To row_num
        If ActiveWorkbook.Sheets(1).Cells(i, 1).Value = name Then
            a = 1
            Exit For
        Else
        End If
    Next
    
    If a = 0 Then
        ActiveWorkbook.Sheets(1).Cells(row_num + 1, 1) = name
        ActiveWorkbook.Sheets(1).Cells(row_num + 1, 2) = password
        ActiveWorkbook.Save
        ActiveWorkbook.Close
    Else
        ActiveWorkbook.Close
        a = 0
    End If

Else
    Workbooks.Add
    ActiveWorkbook.Sheets(1).Range("A1") = name
    ActiveWorkbook.Sheets(1).Range("B1") = password
    ActiveWorkbook.SaveAs SavePath & "\" & "PDF_암호.xlsx"
    ActiveWorkbook.Close

End If

End Sub