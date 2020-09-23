VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   490
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   706
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   330
      Left            =   8325
      TabIndex        =   8
      Top             =   6885
      Width           =   1080
   End
   Begin VB.CommandButton cmdCodeView 
      Caption         =   "View Code"
      Height          =   330
      Left            =   7125
      TabIndex        =   7
      Top             =   6900
      Width           =   1080
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   0
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   585
      TabIndex        =   5
      Top             =   6915
      Width           =   5655
   End
   Begin VB.PictureBox PicImg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9015
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   3
      Top             =   7560
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox PicTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9585
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   330
      Left            =   6360
      TabIndex        =   1
      Top             =   6915
      Width           =   645
   End
   Begin VB.PictureBox RenderDC 
      AutoRedraw      =   -1  'True
      Height          =   6795
      Left            =   0
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   701
      TabIndex        =   0
      Top             =   0
      Width           =   10575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Page:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   6945
      Width           =   420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type HtmlBody
    BodyBackColor As Long
    BodyTextColor As Long
    BkGoundImg As String
    tMarginSize As Integer
End Type

Dim FileName As String
Dim Htm_Body As HtmlBody
Dim nRefreshTime As Integer


Private Sub HtmlDocLoadPage(FileName As String)
Dim HtmData As String

    If FindFile(FileName) = False Then Exit Sub
    
    HtmData = OpenFile(FileName)
    txtCode.Text = HtmData
    PhaseBody HtmData ' Get the main body data
    SetupHtmlDOC RenderDC, Htm_Body.tMarginSize
    SetTitle Form1, PhasePageTitle(HtmData)
    PhaseHtml HtmData, RenderDC
    
End Sub
Private Sub TileBkImage()
Dim i As Long, J As Long
    PicTile.Picture = LoadPicture(Htm_Body.BkGoundImg)
    For i = 0 To (RenderDC.ScaleWidth / PicTile.Width)
        For J = 0 To (RenderDC.ScaleHeight / PicTile.Height)
        BitBlt RenderDC.hDC, PicTile.Width * i, PicTile.Height * J, PicTile.Width, PicTile.Height, PicTile.hDC, 0, 0, vbSrcCopy
        Next
    Next
    RenderDC.Refresh
    i = 0
    J = 0
End Sub

Sub GetImgProp(lzText As String)
Dim lFilename As String
Dim mHeight As Integer, mWidth As Integer


    lFilename = FixPath(StripHomePath(FileName)) & FixWebSlash(RemoveJunk(GetTagData1(lzText, "src=", Chr(34))))
    
    
    mHeight = Val(RemoveJunk(GetTagData1(lzText, "height=", Chr(34))))
    mWidth = Val(RemoveJunk(GetTagData1(lzText, "width=", Chr(34))))
    
    If Not FindFile(lFilename) Then Exit Sub
    PicImg.Picture = LoadPicture(lFilename)

    AddImage RenderDC, PicImg, mWidth, mHeight
End Sub

Private Sub PhaseHtml(lzStr As String, PicDc As PictureBox)
Dim StrB As String, TheTag As String, sHtml As String, NextTag As String
Dim ipos As Long, lPos As Long, nPos As Long
Dim CanRefresh As Boolean, nPageToLoad As String



    StrB = lzStr
    StrB = Replace(StrB, vbCrLf, "")
    lPos = 1
    Do
        DoEvents
        ipos = InStr(lPos, StrB, "<", vbBinaryCompare)
        If ipos = 0 Then Exit Do
        
        sHtml = LTrim(RemoveJunk(Mid(StrB, lPos + 1, ipos - lPos))) ' Html Text
        
        If Len(sHtml) > 0 Then
            DisplayText FormatSpecialTags(sHtml), PicDc
        End If
  
        lPos = InStr(ipos, StrB, ">", vbBinaryCompare)
        If lPos = 0 Then Exit Do
        
        TheTag = LCase(Mid(StrB, ipos, lPos - ipos + 1))
        
        Select Case TheTag
            Case "<br>", "<p>", "</p>"
                AddNewLine PicDc
            Case "<hr>"
                AddNewLine PicDc
                AddHozLine PicDc
            Case "<b>"
                FormatTextStyle PicDc, mBold, True
            Case "</b>"
                FormatTextStyle PicDc, mBold, False
            Case "<i>"
                FormatTextStyle PicDc, mItalic, True
            Case "</i>"
                FormatTextStyle PicDc, mItalic, False
            Case "<u>"
                FormatTextStyle PicDc, mUnderline, True
            Case "</u>"
                FormatTextStyle PicDc, mUnderline, False
            Case "</div>"
                tHtmlDoc.AlignOption = mleft
        End Select
        
        nPos = InStr(1, TheTag, Chr(32), vbBinaryCompare)
        
        If nPos > 0 Then
            NextTag = Trim(Mid(TheTag, 1, nPos))
            
            Select Case NextTag
                Case "<font"
                    PhaseFonts TheTag
                Case "<img"
                    GetImgProp TheTag
                Case "<meta"
                    Select Case LCase(RemoveJunk(GetTagData1(TheTag, "http-equiv=", Chr(34))))
                        Case "refresh"
                            CanRefresh = True
                            nRefreshTime = Val(RemoveJunk(GetTagData1(TheTag, "content=", Chr(34))))
                            FileName = FixPath(CurDir(FileName)) & RemoveJunk(GetTagData1(TheTag, "URL=", Chr(34)))
                        Case Else
                            CanRefresh = False
                    End Select
                    
                Case "<div"
                    Select Case LCase(RemoveJunk(GetTagData1(TheTag, "align=", Chr(34))))
                        Case "left"
                            tHtmlDoc.AlignOption = mleft
                        Case "center"
                            tHtmlDoc.AlignOption = mCenter
                        Case "right"
                            tHtmlDoc.AlignOption = mRight
                    End Select
            End Select
        End If
    Loop

    nPos = 0
    NextTag = ""
    StrB = ""
    sHtml = ""
  
  
    If CanRefresh Then
        Sleep RefreshInterval * nRefreshTime
        HtmlDocLoadPage FileName
    End If
    
                        

    
    
End Sub



Sub PhaseBody(lzText As String)
Dim lPos As Long, hpos As Long
Dim sBody As String, sTemp As String, StrA As String, sPath As String
Dim UseBackGoundImg As Boolean

    sBody = lzText
    
    sBody = Replace(sBody, "<body", "<BODY")
    sBody = Replace(sBody, "<font", "<FONT")
    
    lPos = InStr(1, sBody, "<BODY", vbTextCompare)
    hpos = InStr(lPos + 1, sBody, ">", vbBinaryCompare)
    
    If (lPos > 0) And (hpos > 0) Then
        sTemp = Mid(sBody, lPos + 5, hpos - lPos - 4)
        ' phase out the rest of the body and get background and text color properties
        StrA = RemoveJunk(GetTagData1(sTemp, "bgcolor=", Chr(34))) ' Get page bk colour
        Htm_Body.BodyBackColor = HexToLong(StrA) ' Convert hex color to rgb and store it
        StrA = ""
        
        StrA = RemoveJunk(GetTagData1(sTemp, "text=", Chr(34))) ' extract text color
        Htm_Body.BodyTextColor = HexToLong(StrA) ' Convert hex color to rgb and store it
        StrA = ""
        
        StrA = RemoveJunk(GetTagData1(sTemp, "leftmargin=", Chr(34)))
        Htm_Body.tMarginSize = Val(StrA)
        StrA = ""
        
        StrA = RemoveJunk(GetTagData1(sTemp, "background=", Chr(34))) ' Get the bk image
   
        
        If Len(StrA) = 0 Then
            UseBackGoundImg = False
        End If
        
        Htm_Body.BkGoundImg = FixPath(StripHomePath(FileName)) & FixWebSlash(StrA)
  
    

        If Len(StrA) = 0 Or (FindFile(Htm_Body.BkGoundImg) = False) Then
            UseBackGoundImg = False
        Else
            UseBackGoundImg = True
        End If
        
        lPos = 0
        hpos = 0
        sTemp = ""
        sPath = ""
        sBody = ""
    End If

    If UseBackGoundImg Then
        TileBkImage
    Else
        RenderDC.BackColor = Htm_Body.BodyBackColor
    End If

    XPos = Htm_Body.tMarginSize
    RenderDC.ForeColor = Htm_Body.BodyTextColor
    
    ipos = 0
    hpos = 0
    sTemp = ""
    StrA = ""
    sPath = ""
    sBody = ""
End Sub

Function PhasePageTitle(lzData As String) As String
Dim ipos As Integer
Dim lPos As Integer
Dim StrTemp As String

    ipos = InStr(1, lzData, "<title>", vbTextCompare)
    lPos = InStr(ipos + 1, lzData, "</title>", vbTextCompare)
    
    If (ipos > 0) And (lPos > 0) Then
        StrTemp = Mid(lzData, ipos, lPos - 7)
        lzData = Replace(lzData, StrTemp, "")
        PhasePageTitle = GetTagData1(StrTemp, "<title>", "</title>")
    End If
    
    ipos = 0
    lPos = 0
    StrTemp = ""
    
End Function

Private Sub cmdCodeView_Click()
Dim ViewCode As Boolean
Dim sData As String

    ViewCode = Not ViewCode
    If cmdCodeView.Caption = "View Code" Then
        cmdCodeView.Caption = "Hide Code"
        ViewCode = True
    Else
        cmdCodeView.Caption = "View Code"
        ViewCode = False
    End If
    
    RenderDC.Visible = Not ViewCode
    txtCode.Visible = ViewCode
    sData = txtCode.Text
    
    PhaseBody sData ' Get the main body data
    SetupHtmlDOC RenderDC, Htm_Body.tMarginSize
    SetTitle Form1, PhasePageTitle(sData)
    PhaseHtml sData, RenderDC
    
End Sub

Private Sub Command1_Click()
    FileName = Text1
    HtmlDocLoadPage Text1.Text
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
Dim i As Integer

    Text1.Text = App.Path & "\index.html"
    txtCode.Width = (RenderDC.Width - txtCode.Left)
    txtCode.Height = (RenderDC.Height - txtCode.Top)
    txtCode.Visible = False

    
End Sub

