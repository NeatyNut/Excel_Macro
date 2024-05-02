Sub 메일발송하기()
Dim dirpath As String
Dim filename As String
Dim adrrows As Integer
Dim mailshname As String
Dim madr As String
Dim attach As String
Dim objOutlook As Object
Set objOutlook = CreateObject("Outlook.Application")
    
' CREATE EMAIL OBJECT.
Dim objEmail As Object
Set objEmail = objOutlook.CreateItem(olMailItem)

Dim message As String

mailshname = "이메일주소, 최저시급"
adrrows = Sheets(mailshname).Range("a1").CurrentRegion.Rows.Count

dirpath = Environ("USERPROFILE") & "\Desktop\PDF암호화\"

If Dir(dirpath, vbDirectory) = "" Then
    MkDir dirpath
    MsgBox "기존 폴더가 없어서 폴더를 생성한 뒤 종료합니다."
    Exit Sub
End If

filename = Dir(dirpath & "*.PDF")

i2 = 0

Do While (filename <> "")
    i2 = i2 + 1
    filename = Dir()
Loop

dirpath = Environ("USERPROFILE") & "\Desktop\PDF암호화\"
filename = Dir(dirpath & "*.PDF")

For i3 = 1 To i2 Step 1
    For i = 1 To adrrows Step 1
        If Sheets(mailshname).Cells(i, 1).Value = Mid(filename, InStr(filename, "_") + 1, InStrRev(filename, "_") - InStr(filename, "_") - 1) Then
            madr = Sheets(mailshname).Cells(i, 2).Value
            attach = Environ("USERPROFILE") & "\Desktop" & "\PDF암호화\" & filename
            Set objEmail = objOutlook.CreateItem(olMailItem)
            With objEmail
                .To = madr
                .Subject = Sheets(mailshname).Range("S3").Value
                .Body = Sheets(mailshname).Range("S5").Value & vbNewLine & Sheets(mailshname).Range("S6").Value & vbNewLine & Sheets(mailshname).Range("S7").Value & vbNewLine & Sheets(mailshname).Range("S8").Value
                .attachments.Add attach
                .send        ' send the message in Outlook.
            End With
            madr = "": attach = ""
            message = message & vbNewLine & Mid(filename, InStr(filename, "_") + 1, InStrRev(filename, "_") - InStr(filename, "_") - 1) & " : 발송완료"
            
            filename = Dir()
            Exit For
        ElseIf i = adrrows Then
            message = message & vbNewLine & Mid(filename, InStr(filename, "_") + 1, InStrRev(filename, "_") - InStr(filename, "_") - 1) & " : 이메일 데이터 없음"
        Else
        End If
    Next i
Next i3

If message <> "" Then
    MsgBox message
Else
    MsgBox "임금명세서가 없습니다."
End If
End Sub