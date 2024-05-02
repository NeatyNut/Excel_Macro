Sub Save_as()

    Dim F_Dir As FileDialog
    Dim F_Path, F_Name, C_Name As String
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim SN As String
    Dim Schoolname As String
    
    Set F_Dir = Application.FileDialog(msoFileDialogFolderPicker)
    F_Dir.Show
    F_Path = F_Dir.SelectedItems(1)
    
    F_Name = Dir(F_Path & "\*.xls")
    Application.ScreenUpdating = False
    Do While F_Name <> ""
        If Right(F_Name, 3) = "xls" Then
            C_Name = Left(F_Name, Len(F_Name) - 4)
            Set WB = Workbooks.Open(F_Path & "\" & F_Name)
            Set WS = WB.ActiveSheet
            SN = WS.Cells(Cells(100000, 1).End(xlUp).Row - 2, 1).Value
            Schoolname = (Mid(SN, 6, Len(SN) - 7))
            ActiveWorkbook.SaveAs filename:= _
            F_Path & "\" & Schoolname & ".xlsx" _
            , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            WB.Close
        End If
        F_Name = Dir()
    Loop
    Application.ScreenUpdating = True

    
    If MsgBox(".xls 파일을 삭제하겠습니까?", vbYesNo) = 6 Then
        F2_Name = Dir(F_Path & "\*.xls")
        Do While F2_Name <> ""
            If Right(F2_Name, 3) = "xls" Then
                Kill F_Path & "\" & F2_Name
            End If
            F2_Name = Dir()
        Loop
        MsgBox ("완료하였습니다.")
    Else
        MsgBox ("완료하였습니다.")
    End If

End Sub
