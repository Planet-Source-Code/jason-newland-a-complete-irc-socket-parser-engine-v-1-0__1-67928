Attribute VB_Name = "modAPI"
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Function GetDuration(ByVal Duration As Long) As String
    On Error Resume Next
    Dim strDuration As String
    Dim Sec As Long
    Dim Min As Long
    Dim Hour As Long
    Dim Day As Long
    Dim Week As Long
    Sec = Duration Mod 60
    If Sec >= 60 Then
        Sec = Sec Mod 60
        Min = Int(Duration / 60)
    End If
    If Min < 60 Then
        Min = Int(Duration / 60)
    End If
    If Min >= 60 Then
        Hour = Int(Min / 60)
        Min = Min Mod 60
    End If
    If Hour >= 24 Then
        Day = Int(Hour / 24)
        Hour = Hour Mod 24
    End If
    If Day >= 7 Then
        Week = Int(Day / 7)
        Day = Day Mod 7
    End If
    strDuration = " " & Format(Week, "0") & "weeks " & Format(Day, "0") & "days " & Format(Hour, "0") & "hours " & Format(Min, "0") & "mins " & Format(Sec, "0") & "secs"
    strDuration = Replace(strDuration, " 0weeks", vbNullString)
    strDuration = Replace(strDuration, " 0days", vbNullString)
    strDuration = Replace(strDuration, " 0hours", vbNullString)
    strDuration = Replace(strDuration, " 0mins", vbNullString)
    strDuration = Replace(strDuration, " 0secs", vbNullString)
    strDuration = Replace(strDuration, " 1weeks", " 1week")
    strDuration = Replace(strDuration, " 1days", " 1day")
    strDuration = Replace(strDuration, " 1hours", " 1hour")
    strDuration = Replace(strDuration, " 1mins", " 1min")
    strDuration = Replace(strDuration, " 1secs", " 1sec")
    If LenB(strDuration) = 0 Then strDuration = "0secs"
    GetDuration = Trim$(strDuration)
End Function

