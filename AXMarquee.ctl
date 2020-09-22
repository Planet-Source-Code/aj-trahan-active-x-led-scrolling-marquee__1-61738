VERSION 5.00
Begin VB.UserControl AXMarquee 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10410
   PropertyPages   =   "AXMarquee.ctx":0000
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   694
   ToolboxBitmap   =   "AXMarquee.ctx":0016
   Begin VB.PictureBox picBlankCol 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   3
      Left            =   9120
      Picture         =   "AXMarquee.ctx":0328
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox picBlankCol 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   2
      Left            =   9000
      Picture         =   "AXMarquee.ctx":05AA
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox picBlankCol 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   1
      Left            =   8880
      Picture         =   "AXMarquee.ctx":082C
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox picCaps 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   3
      Left            =   240
      Picture         =   "AXMarquee.ctx":0AAE
      ScaleHeight     =   35.752
      ScaleMode       =   0  'User
      ScaleWidth      =   1714.229
      TabIndex        =   5
      Top             =   2760
      Width           =   25725
   End
   Begin VB.PictureBox picCaps 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   2
      Left            =   240
      Picture         =   "AXMarquee.ctx":2DEE0
      ScaleHeight     =   35.752
      ScaleMode       =   0  'User
      ScaleWidth      =   1714.229
      TabIndex        =   4
      Top             =   2160
      Width           =   25725
   End
   Begin VB.PictureBox picCaps 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   1
      Left            =   240
      Picture         =   "AXMarquee.ctx":5B312
      ScaleHeight     =   35.752
      ScaleMode       =   0  'User
      ScaleWidth      =   1714.229
      TabIndex        =   3
      Top             =   1560
      Width           =   25725
   End
   Begin VB.PictureBox picBlankCol 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   8760
      Picture         =   "AXMarquee.ctx":88744
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Timer tAni 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   3000
   End
   Begin VB.PictureBox picMsg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   120
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   1
      Top             =   2880
      Width           =   1170
   End
   Begin VB.PictureBox picCaps 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   240
      Picture         =   "AXMarquee.ctx":889C6
      ScaleHeight     =   35.752
      ScaleMode       =   0  'User
      ScaleWidth      =   1714.229
      TabIndex        =   0
      Top             =   960
      Width           =   25725
   End
End
Attribute VB_Name = "AXMarquee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum ScrollModeValue
    R_to_L = 0
    L_to_R = 1
End Enum

Enum ForeColorValue
    Red = 0
    Green = 1
    Blue = 2
    Purple = 3
End Enum

'Vars for tracking BMP size and position
Private lBMPWidth   As Long     'Total width of the Message Bitmap to be drawn on the background
Private bRestart    As Boolean
Private lCtlWidth   As Long
Const SRC_Y = 0
Const CTL_HEIGHT = 540
Const m_def_ScrollMode = R_to_L
Const m_def_ForeColor = 0
Const m_def_Text = "Scrolling Marquee OCX .... Created by:  James Miller (2005)"
Const m_def_Scrolling = False
Dim m_ScrollMode As ScrollModeValue
Dim m_ForeColor As ForeColorValue
Dim m_Text As String
Dim m_Scrolling As Boolean

Private Sub tAni_Timer()
    Static lX           As Long
    Static lX2          As Long
    Static lSrcOffset   As Long
    Static lSrcWidth    As Long

    If bRestart Then
        If m_ScrollMode = R_to_L Then
            lX = lCtlWidth - BULB_WIDTH
            lSrcOffset = 0
            lSrcWidth = BULB_WIDTH
        Else
            lX = BULB_WIDTH
            lSrcOffset = BULB_WIDTH
            lSrcWidth = BULB_WIDTH
        End If
        bRestart = False
    End If
    If m_ScrollMode = R_to_L Then
        If lX > 0 Then
            lX2 = lX
            If lCtlWidth - lX <= lBMPWidth Then
                lSrcWidth = lCtlWidth - lX
            Else
                lSrcWidth = lBMPWidth
            End If
        Else
            lX2 = 0
            lSrcOffset = Abs(lX)
            lSrcWidth = lBMPWidth - lSrcOffset
        End If
    Else
        If lX < lCtlWidth Then
            If lX <= lBMPWidth Then
                lX2 = 0
                lSrcWidth = lX
                lSrcOffset = lBMPWidth - lX
            Else
                lX2 = lX2 + BULB_WIDTH
                lSrcWidth = lBMPWidth
                lSrcOffset = 0
            End If
        Else
            If lX > lBMPWidth Then
                lX2 = lX2 + BULB_WIDTH
                lSrcWidth = lBMPWidth
            Else
                lSrcOffset = lBMPWidth - lX
                lSrcWidth = lCtlWidth
            End If
        End If
    End If
    UserControl.PaintPicture picMsg.Picture, lX2, SRC_Y, , , _
        lSrcOffset, , lSrcWidth, , _
        vbSrcCopy
    If m_ScrollMode = R_to_L Then
        If lSrcOffset + BULB_WIDTH = lBMPWidth Then
            bRestart = True
        Else
            lX = lX - BULB_WIDTH
        End If
    Else
        If lX2 + BULB_WIDTH = lCtlWidth Then
            bRestart = True
        Else
            lX = lX + BULB_WIDTH
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    InitBMPStruct
End Sub
Private Sub UserControl_InitProperties()
    m_ScrollMode = m_def_ScrollMode
    m_Text = m_def_Text
    m_Scrolling = True
    m_ForeColor = Red
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ScrollMode = PropBag.ReadProperty("ScrollMode", m_def_ScrollMode)
    Text = PropBag.ReadProperty("Text", m_def_Text)
    Scrolling = PropBag.ReadProperty("Scrolling", m_def_Scrolling)
    ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ScrollMode", m_ScrollMode, m_def_ScrollMode)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Scrolling", m_Scrolling, m_def_Scrolling)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = CTL_HEIGHT
    lCtlWidth = UserControl.ScaleWidth - UserControl.ScaleWidth Mod 5
    DrawBackground
End Sub

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    PropertyChanged "Text"
    If m_Scrolling Then
        tAni.Enabled = False
        bRestart = True
        DrawBackground
        BuildTheBmp (m_Text)
        tAni.Enabled = True
    Else
        tAni.Enabled = False
        bRestart = False
    End If
End Property

Public Property Get Scrolling() As Boolean
    Scrolling = m_Scrolling
End Property

Public Property Let Scrolling(ByVal bScrolling As Boolean)
    m_Scrolling = bScrolling
    PropertyChanged "Scrolling"
    If m_Scrolling Then
        DrawBackground
        BuildTheBmp (m_Text)
        tAni.Enabled = True
    Else
        tAni.Enabled = False
        bRestart = False
    End If
End Property

Public Property Get ForeColor() As ForeColorValue
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As ForeColorValue)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    If m_Scrolling Then
        tAni.Enabled = False
        bRestart = True
        DrawBackground
        BuildTheBmp (m_Text)
        tAni.Enabled = True
    Else
        tAni.Enabled = False
        bRestart = True
    End If
End Property

Public Property Get ScrollMode() As ScrollModeValue
    ScrollMode = m_ScrollMode
End Property

Public Property Let ScrollMode(ByVal New_ScrollMode As ScrollModeValue)
    m_ScrollMode = New_ScrollMode
    PropertyChanged "ScrollMode"
    If m_Scrolling Then
        tAni.Enabled = False
        bRestart = True
        DrawBackground
        BuildTheBmp (m_Text)
        tAni.Enabled = True
    Else
        tAni.Enabled = False
        bRestart = False
    End If
End Property

Private Sub DrawBackground()
    Dim lColX As Long
    Dim lCol As Integer
    
    Select Case m_ForeColor
        Case Red: lCol = 0
        Case Green: lCol = 1
        Case Blue:  lCol = 2
        Case Purple: lCol = 3
    End Select
    With UserControl
        .AutoRedraw = True
        For lColX = 0 To .ScaleWidth Step 5
            .PaintPicture picBlankCol(lCol).Picture, lColX, 0, _
                aCharSpace.Width, , _
                aCharSpace.Left, 0, _
                aCharSpace.Width
        Next lColX
        .AutoRedraw = False
    End With
End Sub

Private Function BuildTheBmp(sText As String) As Long
    Dim lChar     As Long
    Dim lOffset   As Long
    Dim lCharVal  As Long
    Dim lCounter  As Long
    Dim lMsgLength As Long
    Dim lForeColor As Integer
    
    Select Case m_ForeColor
        Case Red: lForeColor = 0
        Case Green: lForeColor = 1
        Case Blue: lForeColor = 2
        Case Purple:  lForeColor = 3
    End Select
    sText = UCase$(sText)
    lMsgLength = Len(sText)
    With picMsg
        .AutoRedraw = True
        For lChar = 1 To lMsgLength
            lCharVal = Asc(Mid$(sText, lChar, 1))
            If lCharVal = 32 Then
                For lCounter = 1 To 4
                    lOffset = lOffset + aCharSpace.Width
                Next lCounter
            ElseIf lCharVal >= 33 And lCharVal <= 90 Then
                lOffset = lOffset + aChars(lCharVal).Width
            End If
        Next lChar
        .Width = lOffset + aCharSpace.Width
        lOffset = 0
        For lChar = 1 To lMsgLength
            lCharVal = Asc(Mid$(sText, lChar, 1))
            If lCharVal = 32 Then
                For lCounter = 1 To 4
                    .PaintPicture picCaps(lForeColor).Picture, lOffset, 0, _
                        aCharSpace.Width, , _
                        aCharSpace.Left, 0, _
                        aCharSpace.Width
                    lOffset = lOffset + aCharSpace.Width
                Next lCounter
            ElseIf lCharVal >= 33 And lCharVal <= 90 Then
                .PaintPicture picCaps(lForeColor).Picture, lOffset, 0, _
                      aChars(lCharVal).Width, , _
                      aChars(lCharVal).Left, 0, _
                      aChars(lCharVal).Width
                lOffset = lOffset + aChars(lCharVal).Width
            Else
                Debug.Print "Unsupported character entered - " & Mid$(sText, lChar, 1) & "ASCII = " & Asc(Mid$(sText, lChar, 1))
            End If
        Next lChar
        .PaintPicture picCaps(lForeColor).Picture, lOffset, 0, _
            aCharSpace.Width, , _
            aCharSpace.Left, 0, _
            aCharSpace.Width
        lOffset = lOffset + aCharSpace.Width
        .AutoRedraw = False
        .Picture = picMsg.Image
    End With
    lBMPWidth = lOffset
    BuildTheBmp = 0
    bRestart = True
End Function
