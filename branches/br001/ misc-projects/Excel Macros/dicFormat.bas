Attribute VB_Name = "dicFormat"
Const FL_SPLIT = 32
Const FL_DEFAULT = 16
Const FL_POST = 8
Const FL_ROUND = 4
Const FL_SIGNED = 2
Const FL_BYPASS = 1
Public Function RawToAscii(ChannelID, Avg, Cnt, Flags, Offset, Mult, Div, Min, Max As Variant, ChanName, Text, Fmt As String, d As Long, Optional ReturnFloat As Boolean) As String

Rem Function requires the list of parameters as supplied by the dic command.
Rem There is an optional argument that specifies whether the returned value is formatted or not.
Rem the returned value is a string.

Dim sbuf As String
Dim str, stmp As String
Dim PreOffset As Long
Dim PostOffset As Long
Dim PreOff, PostOff As Long
Dim dFloat As Single

Rem Check for ReturnFloat flag
If IsMissing(ReturnFloat) Then ReturnFloat = False

Rem Keep floating point version of data to enable a more precise value to be returned.
dFloat = d

If (Flags And FL_BYPASS) = 0 Then
    PreOff = 0
    PostOff = 0
    If (Flags And FL_SPLIT) = 0 Then
        If (Flags And FL_POST) Then
            PostOff = Offset
        Else
            PreOff = Offset
        End If
        d = d + PreOff
        dFloat = dFloat + PreOff
    End If
    d = d * Mult
    dFloat = dFloat * Mult
    If (Flags And FL_ROUND) Then
        If (d >= 0) Then
            d = d + Div / 2
            dFloat = dFloat + Div / 2
        Else
            d = d - Div / 2
            dFloat = dFloat - Div / 2
        End If
    End If
    If (Div <> 0) Then
        d = d / Div
        dFloat = dFloat / Div
    End If
    If (Flags And FL_SPLIT) = 0 Then
        d = d + PostOff
        dFloat = dFloat + PostOff
    End If
    If (d < Min) Then
        d = Min
        dFloat = Min
    End If
    If (Max <> Min) And (d > Max) Then
        d = Max
        dFloat = Max
    End If
End If

If ReturnFloat Then
    RawToAscii = dFloat
    Exit Function
End If
    

If (Fmt <> "") Then
    If (Flags And FL_SPLIT) Then
        Div = Offset
        If (Div <= 0) Then Div = 1
        remainder = Abs(d Mod Div)
        str = Int(d / Div)
        RawToAscii = printf(Fmt, CStr(str) & "," & CStr(remainder))
        If (d < 0 And (d / Div) = 0) Then
            RawToAscii = "-" & RawToAscii
        End If
    Else
        If (Fmt = "%4s") Then
            RawToAscii = StrReverse(printf(Fmt, CStr(d)))
        Else
            RawToAscii = printf(Fmt, CStr(d))
        End If
    End If
Else
    RawToAscii = printf(Fmt, CStr(d))
End If

End Function
Function printf(FmtSpec As String, Args As String) As String

Dim Values() As String

Values = SPLIT(Args, ",")
NumArgs = StringCountOccurrences(FmtSpec, "%")

before = FmtSpec
after = FmtSpec
Start = 0

Rem deal with integers first
For i = Start To UBound(Values)
    after = Replace(after, "%d", Values(i), 1, 1)
    Rem increase starting arg if one has been used
    If before <> after Then
        Start = Start + 1
        before = after
    End If
    'Debug.Print after
Next i

Rem deal with decimals with padding
For i = Start To UBound(Values)
    For j = 1 To 9
        numchars = j - Len(CStr(Values(i)))
        If numchars < 0 Then numchars = 0
        after = Replace(after, "%0" & j & "d", String(numchars, "0") & Values(i), 1, 1)
        If before <> after Then
            Start = Start + 1
            before = after
        End If
    Next j
Next i

Rem deal with hex with padding
For i = Start To UBound(Values)
    For j = 1 To 9
        numchars = j - Len(CStr(Hex(Values(i))))
        If numchars < 0 Then numchars = 0
        after = Replace(after, "0x%0" & j & "X", String(numchars, "0") & Hex(Values(i)), 1, 1)
        If before <> after Then
            Start = Start + 1
            before = after
        End If
    Next j
Next i

Rem deal with strings
For i = Start To UBound(Values)
    For j = 1 To 9
        numchars = j - Len(CStr(Hex(Values(i))))
        If numchars < 0 Then numchars = 0
        after = Replace(after, "%" & j & "s", hex2ascii(Hex(Values(i))), 1, 1)
        If before <> after Then
            Start = Start + 1
            before = after
        End If
    Next j
Next i

printf = after

End Function
Function StringCountOccurrences(strText As String, strFind As String, _
                                Optional lngCompare As VbCompareMethod) As Long
' Counts occurrences of a particular character or characters.
' If lngCompare argument is omitted, procedure performs binary comparison.
'Testcases:
'?StringCountOccurrences("","") = 0
'?StringCountOccurrences("","a") = 0
'?StringCountOccurrences("aaa","a") = 3
'?StringCountOccurrences("aaa","b") = 0
'?StringCountOccurrences("aaa","aa") = 1
Dim lngPos As Long
Dim lngTemp As Long
Dim lngCount As Long
    If Len(strText) = 0 Then Exit Function
    If Len(strFind) = 0 Then Exit Function
    lngPos = 1
    Do
        lngPos = InStr(lngPos, strText, strFind, lngCompare)
        lngTemp = lngPos
        If lngPos > 0 Then
            lngCount = lngCount + 1
            lngPos = lngPos + Len(strFind)
        End If
    Loop Until lngPos = 0
    StringCountOccurrences = lngCount
End Function
Function hex2ascii(ByVal hextext As String) As String
    
For y = 1 To Len(hextext)
    num = Mid(hextext, y, 2)
    Value = Value & Chr(Val("&h" & num))
    y = y + 1
Next y

hex2ascii = Value
End Function
Sub test()
    Debug.Print RawToAscii(3, 27, 8, 20, 0, 47, 400, 0, 0, "VOUT", "Output relay drive voltage", "%dV", 408)
    Debug.Print RawToAscii(3, 27, 8, 20, 0, 47, 400, 0, 0, "VOUT", "Output relay drive voltage", "%dV", 408, True)
End Sub

