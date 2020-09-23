Attribute VB_Name = "Module1"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public EmbedFlashObjCnt As Integer
Public Const RefreshInterval As Integer = 1000
Public Const LineBreakHeight As Integer = 8
Public Const Def_FontColor = vbBlack
Public Const Def_FontSize = 8
Public Const Def_FontName = "MS Sans Serif"

Function OpenFile(lzFile As String) As String
Dim iFile As Long, StrB As String
    iFile = FreeFile
    Open lzFile For Binary As #iFile
        StrB = Space(LOF(iFile))
        Get #iFile, , StrB
    Close #iFile
    
    OpenFile = StrB
    StrB = ""
End Function

Function FixPath(lzPath As String) As String
    If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function
