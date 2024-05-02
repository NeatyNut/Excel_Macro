Function makedate(날짜 As Range, 돈 As Range)
Dim F_makedate As String
Dim srtdate As Date
Dim enddate As Date
Dim mon As Integer

F_makedate = ""
srtdate = 0
enddate = 0
mon = 0

For i = 1 To 날짜.Columns.Count Step 1
    If 돈(1, i).Value > 0 And IsNumeric(돈(1, i).Value) Then
        If srtdate = 0 Then
            srtdate = 날짜(1, i)
            enddate = 날짜(1, i)
        Else
            enddate = 날짜(1, i)
        End If
        
    ElseIf srtdate = 0 Then
        Debug.Print "nothing"
    ElseIf srtdate = enddate Then
        If F_makedate = "" Then
            F_makedate = month(srtdate) & "." & day(srtdate)
            mon = month(srtdate)
        Else
            If month(srtdate) = mon Then '기존 문구와 month 비교
                F_makedate = F_makedate & ", " & day(srtdate)
            Else
                F_makedate = F_makedate & ", " & month(srtdate) & "." & day(srtdate)
                mon = month(srtdate)
            End If
        End If
        
        srtdate = 0
        enddate = 0
    Else
        If F_makedate = "" Then
            If month(srtdate) = month(enddate) Then
                F_makedate = month(srtdate) & "." & day(srtdate) & "~" & day(enddate)
                mon = month(srtdate)
            Else
                F_makedate = month(srtdate) & "." & day(srtdate) & "~" & month(enddate) & "." & day(enddate)
                mon = month(enddate)
            End If
        Else
            If month(srtdate) = month(enddate) Then
                If month(srtdate) = mon Then
                    F_makedate = F_makedate & ", " & day(srtdate) & "~" & day(enddate)
                Else
                    F_makedate = F_makedate & ", " & month(srtdate) & "." & day(srtdate) & "~" & day(enddate)
                End If
            Else
                F_makedate = F_makedate & ", " & month(srtdate) & "." & day(srtdate) & "~" & month(enddate) & "." & day(enddate)
            End If
        End If
        
        srtdate = 0
        enddate = 0
    End If
Next i

'한번더
If srtdate = enddate And srtdate <> 0 Then
    If F_makedate = "" Then
        F_makedate = month(srtdate) & "." & day(srtdate)
    Else
        If month(srtdate) = mon Then '기존 문구와 month 비교
            F_makedate = F_makedate & ", " & day(srtdate)
        Else
            F_makedate = F_makedate & ", " & month(srtdate) & "." & day(srtdate)
        End If
    End If
Else
    If srtdate <> 0 Or enddate <> 0 Then
        If F_makedate = "" Then
            If month(srtdate) = month(enddate) Then
                F_makedate = month(srtdate) & "." & day(srtdate) & "~" & day(enddate)
            Else
                F_makedate = month(srtdate) & "." & day(srtdate) & "~" & month(enddate) & "." & day(enddate)
            End If
        Else
            If month(srtdate) = month(enddate) Then
                If month(srtdate) = mon Then
                    F_makedate = F_makedate & ", " & day(srtdate) & "~" & day(enddate)
                Else
                    F_makedate = F_makedate & ", " & month(srtdate) & "." & day(srtdate) & "~" & day(enddate)
                End If
            Else
                F_makedate = F_makedate & ", " & month(srtdate) & "." & day(srtdate) & "~" & month(enddate) & "." & day(enddate)
            End If
        End If
    End If
End If
makedate = F_makedate

End Function