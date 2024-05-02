Sub 파일정리()
Dim dirpath As String
Dim filename As String
Dim i As String

dirpath = Environ("USERPROFILE") & "\Desktop\PDF암호화\"

If Dir(dirpath, vbDirectory) = "" Then
    MkDir dirpath
    MsgBox "기존 폴더가 없어서 생성합니다."
    Exit Sub
End If


filename = Dir(dirpath & "*.PDF")

Do While (filename <> "")
    Kill dirpath & filename
    filename = Dir()
Loop

On Error Resume Next
    Kill dirpath & "PDF_암호.xlsx"

MsgBox "폴더 정리 완료하였습니다."

End Sub