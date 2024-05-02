Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
    ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Sub make_sheets()
Dim xWs As Worksheet
Dim uni_name As String
Dim rownum As Integer

Application.ScreenUpdating = False
Application.DisplayAlerts = False

For Each xWs In Application.ActiveWorkbook.Worksheets
    If xWs.Name <> "출처" And xWs.Name <> "목록" And xWs.Name <> "현황" Then
        xWs.Delete
    End If
Next

rownum = ActiveWorkbook.Sheets("목록").Range("a1").CurrentRegion.Rows.Count

For i = 2 To rownum Step 1
    Sheets.Add(after:=Sheets("목록")).Name = ActiveWorkbook.Sheets("목록").Cells(i, 1).Value
Next i

Sheets("목록").Activate

Application.DisplayAlerts = True
Application.ScreenUpdating = True

MsgBox ("완료하였습니다")

End Sub

Sub 작업()
Dim rownum As Integer
Dim fst_num As Integer
Dim col As Integer
Dim appIE As Object
Dim text As Object
Dim Dir As String


Dir = ThisWorkbook.Path

'Application.ScreenUpdating = False
'Application.DisplayAlerts = False

rownum = Sheets("목록").Cells(1, 1).CurrentRegion.Rows.Count
col = 1

For TT = 2 To rownum Step 1
    If Sheets("목록").Cells(TT, 5).Value = "X" Then
        fst_num = 1
        Sheets(Sheets("목록").Cells(TT, 1).Value).Activate
        Range("A1:X100").Select
        Selection.ClearContents
        
        Dim shp As Shape
        For Each shp In ActiveSheet.Shapes
            shp.Delete
        Next shp
        
        Set IE = CreateObject("internetexplorer.application")
        With IE
            .navigate (Sheets("목록").Cells(TT, 2).Value)
            .Visible = True
        End With
    
        Do While (IE.readyState <> READYSTATE_COMPLETE Or IE.Busy = True)
            DoEvents
            Application.Wait DateAdd("s", 3, Now)
        Loop
    
        Select Case Sheets("목록").Cells(TT, 3).Value
        
        Case "class"
            For Each text In IE.Document.getElementsByClassName(Sheets("목록").Cells(TT, 4).Value)
                Sheets(Sheets("목록").Cells(TT, 1).Value).Cells(fst_num, col) = text.innertext
                fst_num = fst_num + 1
            Next
        Case "image"
            Dim objXMLHTTP As Object
            Dim objADOStream As Object

            Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
            Set objADOStream = CreateObject("ADODB.Stream")

            objXMLHTTP.Open "GET", Sheets("목록").Cells(TT, 2).Value, False
            objXMLHTTP.Send

            If objXMLHTTP.Status = 200 Then
                objADOStream.Open
                objADOStream.Type = 1 'adTypeBinary

                objADOStream.Write objXMLHTTP.ResponseBody
                objADOStream.SaveToFile Dir & "\image.png", 2 'adSaveCreateOverWrite
                objADOStream.Close
            End If

            Set objADOStream = Nothing
            Set objXMLHTTP = Nothing
            
            With ThisWorkbook.Sheets(Sheets("목록").Cells(TT, 1).Value).Pictures.Insert(Dir & "\image.png")
                .Left = 10
                .Top = 10
            End With
        Case "imagenum"
            Set imgElements = IE.Document.getElementsByTagName("img")

            For i = 0 To imgElements.Length - 1
                Set imgElement = imgElements(i)
                imgsrc = imgElement.src
                URLDownloadToFile 0, imgsrc, Dir & "\image" & Str(i) & ".jpg", 0, 0
            Next i
            
            'Set imgClass = IE.Document.getElementsByClassName(Sheets("목록").Cells(TT, 4).Value)
            'imgsrc = imgClass(0).getAttribute("src")
            'URLDownloadToFile 0, imgsrc, Dir & "\image.png", 0, 0
              
            With ThisWorkbook.Sheets(Sheets("목록").Cells(TT, 1).Value).Pictures.Insert(Dir & "\image" & Str(Sheets("목록").Cells(TT, 4).Value) & ".jpg")
                .Left = 10
                .Top = 10
            End With
            
            For ii = 0 To imgElements.Length - 1
                'If ii <> Sheets("목록").Cells(TT, 4).Value Then
                    Kill Dir & "\image" & Str(ii) & ".jpg"
                'End If
            Next ii
            
        Case "id"
            Sheets(Sheets("목록").Cells(TT, 1).Value).Cells(fst_num, col) = IE.Document.getelementbyid(Sheets("목록").Cells(TT, 4).Value).innertext

        End Select
        
        IE.Quit
        Set IE = Nothing
    Else
    End If
    
Next TT
MsgBox "완료"
Sheets("목록").Activate

'Application.DisplayAlerts = True
'Application.ScreenUpdating = True

End Sub

