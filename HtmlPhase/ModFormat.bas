Attribute VB_Name = "ModFormat"
Enum TextStyles
    mBold = 1
    mItalic
    mUnderline
End Enum

Enum AlignType
    mleft = 1
    mCenter
    mRight
End Enum

Public Type HtmlDoc
    tFontColor As Long
    tFontSize As Integer
    tFontName As String
    AlignText As Boolean
    AlignOption As AlignType
End Type

Public tHtmlDoc As HtmlDoc
Public Ypos As Long, XPos As Long

Function FormatTextStyle(HtmDoc As PictureBox, tStyle As TextStyles, mEnable As Boolean)
' Function used to set text styles
    Select Case tStyle
        Case mBold
            HtmDoc.FontBold = mEnable
        Case mItalic
            HtmDoc.FontItalic = mEnable
        Case mUnderline
            HtmDoc.FontUnderline = mEnable
    End Select

End Function

Function SetAlignment(HtmDoc As PictureBox, mAlign As AlignType, Optional TextLength As Integer)
    Select Case mAlign
        Case mleft
            HtmDoc.CurrentX = XPos
        Case mCenter
            HtmDoc.CurrentX = (HtmDoc.ScaleWidth - TextLength - 5) / 2
        Case mRight
            HtmDoc.CurrentX = (HtmDoc.ScaleWidth - TextLength * 6)
    End Select
    
End Function

Function AddHozLine(HtmDoc As PictureBox)
    HtmDoc.Line (XPos, Ypos + LineBreakHeight)-(HtmDoc.ScaleWidth, Ypos + LineBreakHeight), &H808080
    HtmDoc.Line (XPos, Ypos + LineBreakHeight + 1)-(HtmDoc.ScaleWidth, Ypos + LineBreakHeight + 1), &HC8D0D4
End Function

Function AddNewLine(HtmDoc As PictureBox)
    Ypos = Ypos + LineBreakHeight * 2
    HtmDoc.CurrentY = Ypos
    HtmDoc.CurrentX = XPos
End Function

Function DisplayText(lzText As String, HtmlDoc As PictureBox)
    
    HtmlDoc.CurrentY = Ypos
    HtmlDoc.Font.Name = tHtmlDoc.tFontName
    HtmlDoc.Font.Size = tHtmlDoc.tFontSize
    HtmlDoc.ForeColor = tHtmlDoc.tFontColor
    SetAlignment HtmlDoc, tHtmlDoc.AlignOption, Len(lzText)
    HtmlDoc.Print lzText
    
End Function

Sub AddImage(HtmlDoc As PictureBox, PictureSrc As PictureBox, imgWidth As Integer, imgHeight As Integer)
    Select Case tHtmlDoc.AlignOption
        Case mCenter
           XPos = (HtmlDoc.ScaleWidth - imgWidth - 5) / 2
        Case mRight
            XPos = (HtmlDoc.ScaleWidth - imgWidth)
    End Select
    
    BitBlt HtmlDoc.hDC, XPos, Ypos, imgWidth, imgHeight, PictureSrc.hDC, 0, 0, vbSrcCopy
    HtmlDoc.Refresh
    
End Sub

Sub SetTitle(Frm As Form, sTitle As String)
    Frm.Caption = sTitle
End Sub

Sub SetLeftMargin(HtmDoc As PictureBox, MarginSize As Integer)
    HtmDoc.CurrentX = XPos
End Sub

Function SetupHtmlDOC(HtmDoc As PictureBox, Optional LeftMarginSize As Integer = 5, Optional TopMarginSize As Integer = 10)
    tHtmlDoc.tFontColor = Def_FontColor
    tHtmlDoc.tFontSize = Def_FontSize
    tHtmlDoc.tFontName = Def_FontName
    tHtmlDoc.AlignOption = mleft
    
    XPos = LeftMarginSize
    Ypos = TopMarginSize
    
    HtmDoc.CurrentX = XPos
    HtmDoc.CurrentY = Ypos
    Set HtmDoc = Nothing
End Function
