Attribute VB_Name = "ModOther"

Public Function FindFile(lzFile As String) As Boolean
    If Dir(lzFile) <> "" Then FindFile = True Else FindFile = False
End Function

Public Function GetTagData1(StrString As String, StartPos As String, EndPos As String) As String
Dim ipos As Long, lPos As Long
    If InStr(1, StrString, StartPos, vbTextCompare) > 0 Then
        ipos = InStr(1, StrString, StartPos, vbTextCompare) + Len(StartPos)
        lPos = InStr(ipos + 1, StrString, EndPos, vbTextCompare)
    End If

    If (ipos > 0) And (lPos > 0) Then
        GetTagData1 = Mid(StrString, ipos, lPos - ipos)
    Else
        GetTagData1 = ""
    End If
    
    ipos = 0
    lPos = 0
    
End Function

Public Function FixWebSlash(lzStr As String) As String
    FixWebSlash = Replace(lzStr, "/", "\")
End Function

Public Function FixFontSize(nSize As Integer) As Integer

    Select Case nSize
        Case 0
            FixFontSize = Def_FontSize
        Case 1
            FixFontSize = Def_FontSize
        Case 2
            FixFontSize = 9
        Case 3
            FixFontSize = 10
        Case 4
            FixFontSize = 14
        Case 5
            FixFontSize = 16
        Case 6
            FixFontSize = 18
        Case 7
            FixFontSize = 20
        Case Else
            FixFontSize = Def_FontSize
    End Select
    
End Function

Public Function HexToLong(StrHex As String) As Long
Dim Red As Integer, Green As Integer, Blue As Integer
Dim HexTmp As String

    HexTmp = StrHex
    If Len(HexTmp) = 0 Then HexToLong = 0: Exit Function
    
    If Left(StrHex, 1) = "#" Then
        HexTmp = Right(StrHex, Len(StrHex) - 1)
    End If
    
    Red = CInt("&H" & Mid(HexTmp, 1, 2))
    Green = CInt("&H" & Mid(HexTmp, 3, 2))
    Blue = CInt("&H" & Mid(HexTmp, 5, 2))
    
    HexToLong = RGB(Red, Green, Blue)
    HexTmp = ""
    Red = 0
    Green = 0
    Blue = 0
    
End Function

Public Function FormatSpecialTags(lzStr As String) As String
Dim StrA As String
Dim Counter As Integer

    StrA = lzStr
    
    StrA = Replace(StrA, "&amp;", "&")
    StrA = Replace(StrA, "&copy;", "©")
    StrA = Replace(StrA, "&reg;", "®")
    StrA = Replace(StrA, "&pound;", "£")
    StrA = Replace(StrA, "&yen;", "¥")
    StrA = Replace(StrA, "&euro;", "€")
    StrA = Replace(StrA, "&laquo;", "«")
    StrA = Replace(StrA, "&raquo;", "»")
    StrA = Replace(StrA, "&iquest;", "¿")
    StrA = Replace(StrA, "&para;", "¶")
    StrA = Replace(StrA, "&not;", "¬")
    StrA = Replace(StrA, "&plusmn;", "±")
    StrA = Replace(StrA, "&deg;", "°")
    StrA = Replace(StrA, "&Agrave;", "À")
    
    StrA = Replace(StrA, "&nbsp;", Chr(32))
    StrA = Replace(StrA, "&quot;", Chr(34))
    
    StrA = Replace(StrA, "&#9;", vbTab)
    StrA = Replace(StrA, "&#130;", 130)
    StrA = Replace(StrA, "&#140;", Chr(140))
    StrA = Replace(StrA, "&#137;", Chr(137))
    StrA = Replace(StrA, "&#147;", Chr(147))
    StrA = Replace(StrA, "&#149;", Chr(149))
    StrA = Replace(StrA, "&#150;", Chr(150))
    StrA = Replace(StrA, "&#151;", Chr(151))
    StrA = Replace(StrA, "&#153;", Chr(153))
    
    FormatSpecialTags = StrA
    StrA = ""
    lzStr = ""
End Function

Public Function RemoveJunk(StrString As String) As String
Dim StrA As String
    StrA = StrString
    StrA = Replace(StrA, Chr(34), "")
    StrA = Replace(StrA, "<", "")
    StrA = Replace(StrA, ">", "")
    RemoveJunk = StrA
    StrA = ""
End Function

Public Function StripHomePath(lzFilename As String)
Dim i As Integer, ipos As Integer
    For i = 1 To Len(lzFilename)
        If Mid(lzFilename, i, 1) = "\" Then ipos = i
    Next
    
    i = 0
    StripHomePath = Trim(Mid(lzFilename, 1, ipos))
    ipos = 0
End Function

Public Sub PhaseFonts(lzText As String)
Dim nFontColor As Long
Dim nFontSize As Integer
Dim nFontName As String

    nFontColor = HexToLong(RemoveJunk(GetTagData1(lzText, "font color=", Chr(34)))) ' Get Font Color
    nFontSize = Val(RemoveJunk(GetTagData1(lzText, "size=", Chr(34)))) ' Get the Font Size
    If nFontSize = 0 Then nFontSize = Def_FontSize
    nFontName = RemoveJunk(GetTagData1(lzText, "face=", Chr(34))) ' Get font name
    If Len(nFontName) = 0 Then nFontName = Def_FontName
    
    tHtmlDoc.tFontName = nFontName
    tHtmlDoc.tFontColor = nFontColor
    tHtmlDoc.tFontSize = FixFontSize(nFontSize)
    
    lzText = ""
End Sub

