Sub PDF암호화()

dirpath = Environ("USERPROFILE") & "\Desktop\PDF암호화\"

If Dir(dirpath, vbDirectory) = "" Then
    MkDir dirpath
    MsgBox "기존 폴더가 없어서 생성합니다. 자체 암호화 프로그램 구비 후 다시 실행해주시기 바랍니다."
End If

'암호화파일 경로
sProgpath = Environ("USERPROFILE") & "\Desktop\PDF암호화\pdf암호화_바탕화면PDF암호화 폴더.exe"

If Dir(sProgpath, vbDirectory) <> "" Then
    Proc = Shell(sProgpath, vbNormalFocus) '파일 실행!!
Else
    MsgBox "자체 암호화 프로그램이 존재하지 않습니다. 구비 후 다시 실행해주시기 바랍니다."
End If

End Sub