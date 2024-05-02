Sub WebScraping()
    Dim ie As InternetExplorer
    Dim htmlDoc As HTMLDocument
    Dim htmlDoc2 As HTMLDocument
    Dim htmlElements As IHTMLElementCollection
    Dim htmlElement As IHTMLElement
    Dim a As Integer
    Dim uni As String '대학명
    

    For a = 1 To ThisWorkbook.ActiveSheet.Cells(1, 1).CurrentRegion.Rows.Count Step 1
            
        If ThisWorkbook.ActiveSheet.Cells(a, 2) = "" Then
            uni = ThisWorkbook.ActiveSheet.Cells(a, 1)
        
            Set ie = New InternetExplorer
            ie.Visible = True
            ie.navigate "https://www.naver.com"
        
            Do While ie.readyState <> READYSTATE_COMPLETE Or ie.Busy = True
                DoEvents
            Loop
            
            Set htmlDoc = ie.document
            htmlDoc.getElementById("query").Value = uni
            htmlDoc.getElementById("search-btn").Click
    
            Do While ie.readyState <> READYSTATE_COMPLETE Or ie.Busy = True
                DoEvents
            Loop
        
            Application.Wait (Now + TimeValue("0:00:05"))
    
            Set htmlDoc2 = ie.document
            Set htmlElements = htmlDoc2.getElementsByClassName("text")
            
            On Error Resume Next
            ThisWorkbook.ActiveSheet.Cells(a, 2) = htmlElements.Item(2).href
            
            ie.Quit
            Set ie = Nothing
            Set htmlDoc = Nothing
            Set htmlDoc2 = Nothing
            Set htmlElements = Nothing
            Set htmlElement = Nothing
        Else
        End If
    Next a
    
    
End Sub
