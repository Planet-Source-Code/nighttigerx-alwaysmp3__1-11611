VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1950
      Picture         =   "Form1.frx":D222
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   23
      Top             =   1260
      Width           =   495
   End
   Begin VB.PictureBox picbalance 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   435
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   88
      TabIndex        =   17
      Top             =   4080
      Width           =   1350
   End
   Begin VB.PictureBox picpuse 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   4500
      ScaleHeight     =   570
      ScaleWidth      =   420
      TabIndex        =   16
      ToolTipText     =   "Puse"
      Top             =   4155
      Width           =   450
   End
   Begin VB.PictureBox pbuttons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   405
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   12
      Top             =   1440
      Width           =   1350
   End
   Begin VB.PictureBox pallbuttons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   240
      Picture         =   "Form1.frx":DAEC
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   11
      Top             =   7680
      Width           =   6330
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      Caption         =   "X"
      Height          =   360
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "-"
      Height          =   360
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   255
   End
   Begin VB.Timer tm3 
      Interval        =   1
      Left            =   4200
      Top             =   7320
   End
   Begin VB.Timer tm2 
      Left            =   3840
      Top             =   7320
   End
   Begin VB.PictureBox ipos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1920
      Picture         =   "Form1.frx":1A00E
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   3
      Top             =   4155
      Width           =   315
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2040
      ScaleHeight     =   345
      ScaleWidth      =   2505
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
      Begin VB.Label lb 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "NightTigerX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Left            =   840
         TabIndex        =   2
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.Timer tm1 
      Interval        =   60
      Left            =   3480
      Top             =   7320
   End
   Begin MSComctlLib.ImageList img 
      Left            =   3480
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483625
      ImageWidth      =   53
      ImageHeight     =   71
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A407
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D0BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1FD6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":22A23
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":256D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2838B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B03F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txcu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Computer"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   390
      Left            =   2160
      TabIndex        =   7
      Text            =   "0"
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Mute"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   360
      ScaleHeight     =   1425
      ScaleWidth      =   1425
      TabIndex        =   13
      Top             =   2280
      Width           =   1455
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "Volume"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Top             =   1200
         Width           =   855
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   240
         X2              =   240
         Y1              =   360
         Y2              =   1200
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFF80&
         BorderWidth     =   2
         X1              =   1200
         X2              =   360
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   1200
         X2              =   1200
         Y1              =   840
         Y2              =   120
      End
      Begin VB.Label labplus 
         BackColor       =   &H80000008&
         Caption         =   "+"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "High"
         Top             =   0
         Width           =   255
         WordWrap        =   -1  'True
      End
      Begin VB.Label labne 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "-"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "low"
         Top             =   1080
         Width           =   255
      End
      Begin VB.Image im 
         Appearance      =   0  'Flat
         Height          =   1215
         Left            =   360
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2640
      Picture         =   "Form1.frx":2DCF3
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lbtime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "0:0 min."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3240
      TabIndex        =   8
      Top             =   1320
      Width           =   825
   End
   Begin VB.Image imgicon 
      Height          =   480
      Left            =   4200
      Picture         =   "Form1.frx":2E5BD
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   360
      MouseIcon       =   "Form1.frx":2E8C7
      MousePointer    =   15  'Size All
      TabIndex        =   21
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Time"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFC0C0&
      X1              =   24
      X2              =   120
      Y1              =   248
      Y2              =   248
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   9
      X1              =   24
      X2              =   120
      Y1              =   248
      Y2              =   248
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "sec."
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   4320
      Width           =   615
   End
   Begin MediaPlayerCtl.MediaPlayer mp9 
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   5880
      Width           =   3135
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   0
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -1
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lbdur 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Computer"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   4365
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   4320
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim v, v1, v2
Dim title, vol, dm, vol1, volrc
Dim t1, t2, t3, t4, t5
Dim vp1, vp2, c3, pos
Dim b, cb, mu

Private Sub Command1_Click()
mu = mu + 1
  If mu = 1 Then
    mp9.Mute = True
  End If
  If mu = 2 Then
    mp9.Mute = False
    mu = 0
   End If

End Sub

Private Sub Command4_Click()
Me.Hide
End Sub
Private Sub Command5_Click()
End
End Sub
Private Sub Form_Load()
Dim dm
Dim dummy

AutoFormShape Form1, RGB(255, 0, 255)

v = 4
im.Picture = img.ListImages(4).Picture
volrc = Hex((32767))
 vol1 = "&h" & Trim((volrc)) & Trim((volrc))
  
dm = waveOutSetVolume(0, vol1)
'................................'
With nid
        .cbSize = Len(nid)
        .hWnd = Me.hWnd
        .uId = vbNull
       .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
       .hIcon = Me.Icon
     
       End With
       Shell_NotifyIcon NIM_ADD, nid
  '.....................................'
  dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 40, pallbuttons.hdc, 0, 0, SRCCOPY)
  dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 40, pallbuttons.hdc, 61, 0, SRCCOPY)
  dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 40, pallbuttons.hdc, 121, 0, SRCCOPY)
  dummy = BitBlt(picpuse.hdc, 0, 0, 30, 40, pallbuttons.hdc, 181, 0, SRCCOPY)
  dummy = BitBlt(picbalance.hdc, 0, 0, 30, 40, pallbuttons.hdc, 241, 0, SRCCOPY)
  dummy = BitBlt(picbalance.hdc, 31, 0, 30, 40, pallbuttons.hdc, 331, 0, SRCCOPY)
  dummy = BitBlt(picbalance.hdc, 61, 0, 30, 40, pallbuttons.hdc, 361, 0, SRCCOPY)

mp9.AutoRewind = True
mp9.AutoStart = True


End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Res As Long
Dim msg As Long
    
     If Me.ScaleMode = vbPixels Then
         msg = X
     Else
         msg = X / Screen.TwipsPerPixelX
     End If
    
     Select Case msg
                
        Case WM_LBUTTONUP
             Res = SetForegroundWindow(Me.hWnd)
            Me.Show
         Case WM_RBUTTONUP
          Unload Me
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Shell_NotifyIcon NIM_DELETE, nid
End Sub


Private Sub imgmove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim result As Long

If Button = 1 Then
  ReleaseCapture  ' This releases the mouse communication with the form so it can communicate with the operating system to move the form
  result& = SendMessage(Me.hWnd, &H112, &HF012, 0)  ' This tells the OS to pick up the form to be moved
DoEvents


End If
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim result As Long

If Button = 1 Then
  ReleaseCapture  ' This releases the mouse communication with the form so it can communicate with the operating system to move the form
  result& = SendMessage(Me.hWnd, &H112, &HF012, 0)  ' This tells the OS to pick up the form to be moved
DoEvents


End If
End Sub

Private Sub labne_Click()
v = v + 1
If v = 0 Then
 v = 1
 End If
 
If v = 1 Then
im.Picture = img.ListImages(1).Picture
 volrc = Hex((65535))
End If
If v = 2 Then
im.Picture = img.ListImages(2).Picture
 volrc = Hex((54612.5))
End If
If v = 3 Then
im.Picture = img.ListImages(3).Picture
volrc = Hex((43690))
End If
If v = 4 Then
im.Picture = img.ListImages(4).Picture
volrc = Hex((32767))
End If
If v = 5 Then
im.Picture = img.ListImages(5).Picture
volrc = Hex((21845))
End If
If v = 6 Then
im.Picture = img.ListImages(6).Picture
volrc = Hex((10922))
End If
If v = 7 Then
im.Picture = img.ListImages(7).Picture
volrc = Hex((0))
End If
If v = 8 Then
v = 7
End If
 vol1 = "&h" & Trim((volrc)) & Trim((volrc))
  
 dm = waveOutSetVolume(0, vol1)
End Sub

Private Sub labplus_Click()
v = v - 1
If v = 1 Then
im.Picture = img.ListImages(1).Picture
 volrc = Hex((65535))
End If
If v = 2 Then
im.Picture = img.ListImages(2).Picture
 volrc = Hex((54612.5))
End If
If v = 3 Then
im.Picture = img.ListImages(3).Picture
volrc = Hex((43690))
End If
If v = 4 Then
im.Picture = img.ListImages(4).Picture
volrc = Hex((32767))
End If
If v = 5 Then
im.Picture = img.ListImages(5).Picture
volrc = Hex((21845))
End If
If v = 6 Then
im.Picture = img.ListImages(6).Picture
volrc = Hex((10922))
End If
If v = 7 Then
im.Picture = img.ListImages(7).Picture
volrc = Hex((0))
End If
If v = 8 Then
v = 7
End If
 vol1 = "&h" & Trim((volrc)) & Trim((volrc))
  
 dm = waveOutSetVolume(0, vol1)
End Sub
Public Sub openf()

Dim filebox As OPENFILENAME
Dim fname As String
Dim retval As Long


' Configure how the dialog box will look
filebox.lStructSize = Len(filebox)
filebox.hwndOwner = Me.hWnd
filebox.lpstrTitle = "Open File"
' The next line sets up the file types drop-box
filebox.lpstrFilter = "MP3 Files" & vbNullChar & "*.mp3" & vbNullChar & vbNullChar
filebox.lpstrFile = Space(255)
filebox.nMaxFile = 255
filebox.lpstrFileTitle = Space(255)
filebox.nMaxFileTitle = 255
' Allow only existing files and hide the read-only check box
'filebox.flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY


retval = GetOpenFileName(filebox)

If retval <> 0 Then
  fname = Left(filebox.lpstrFile, InStr(filebox.lpstrFile, vbNullChar) - 1)
End If
filename = fname
title = fname
mp9.filename = fname
mp9.Stop


End Sub

 
Private Sub pbuttons_Click()
If ctr = 1 Then
  openf
lb.Caption = filename

t1 = mp9.Duration
t2 = (t1 \ 60)
t3 = t1 Mod 60
t4 = (t3 / 100)
t5 = Mid(t4, 3, 20)

If t3 = 0 Then
     t5 = "00"
          End If
        If t3 Mod 10 = 0 Then
            t5 = t3
         End If
      lbtime.Caption = t2 & ":" & t5 & " " & "min."
  
  lbdur.Caption = (t2 + (t3 / 60)) * 60

  End If

If ctr = 2 Then
  Dim dm
    If filename = "" Then Exit Sub
      mp9.Play
     tm2.Interval = 1
    ipos.Left = 130
  tm3.Enabled = True
  
End If

If ctr = 3 Then
  mp9.Stop
   mp9.CurrentPosition = 0
    ipos.Left = 130
    v1 = 0: v2 = 0
  
 End If
End Sub

Private Sub pbuttons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dummy As Long

If Button = 1 Then
  If X >= 0 And X <= 30 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 40, pallbuttons.hdc, 30, 0, SRCCOPY)
    dummy = BitBlt(pbuttons.hdc, 30, 0, 30, 40, pallbuttons.hdc, 61, 0, SRCCOPY)
    dummy = BitBlt(pbuttons.hdc, 60, 0, 30, 40, pallbuttons.hdc, 120, 0, SRCCOPY)
           
   ctr = 1
  End If
  If X >= 31 And X <= 60 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 40, pallbuttons.hdc, 90, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 40, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 40, pallbuttons.hdc, 121, 0, SRCCOPY)
          
   ctr = 2
  End If
  If X >= 61 And X <= 90 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 40, pallbuttons.hdc, 150, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 40, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 40, pallbuttons.hdc, 61, 0, SRCCOPY)
           'd
   ctr = 3
  End If
  End If

End Sub

Private Sub pbuttons_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dummy As Long

If Button = 1 Then
  If ctr = 1 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 40, pallbuttons.hdc, 0, 0, SRCCOPY)
          
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 40, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 40, pallbuttons.hdc, 121, 0, SRCCOPY)
           
  End If
  If ctr = 2 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 40, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 40, pallbuttons.hdc, 0, 0, SRCCOPY)
          
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 40, pallbuttons.hdc, 121, 0, SRCCOPY)
          

  End If
  If ctr = 3 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 40, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 40, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 40, pallbuttons.hdc, 61, 0, SRCCOPY)
           

     End If
  End If
End Sub


Private Sub picbalance_Click()
If cb = 1 Then
 mp9.Balance = -10000
 End If
 If cb = 2 Then
 mp9.Balance = 0
 End If
 If cb = 3 Then
  mp9.Balance = 10000
  End If
End Sub

Private Sub picbalance_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dummy As Long

If Button = 1 Then
  If X >= 0 And X <= 30 Then
    picbalance.Cls
    dummy = BitBlt(picbalance.hdc, 0, 0, 30, 40, pallbuttons.hdc, 271, 0, SRCCOPY)
    dummy = BitBlt(picbalance.hdc, 31, 0, 30, 40, pallbuttons.hdc, 301, 0, SRCCOPY)
    dummy = BitBlt(picbalance.hdc, 61, 0, 30, 40, pallbuttons.hdc, 361, 0, SRCCOPY)
           
   cb = 1
  End If
  If X >= 31 And X <= 60 Then
    picbalance.Cls
    dummy = BitBlt(picbalance.hdc, 0, 0, 30, 40, pallbuttons.hdc, 241, 0, SRCCOPY)
    dummy = BitBlt(picbalance.hdc, 31, 0, 30, 40, pallbuttons.hdc, 331, 0, SRCCOPY)
    dummy = BitBlt(picbalance.hdc, 61, 0, 30, 40, pallbuttons.hdc, 361, 0, SRCCOPY)
             
          
          
   cb = 2
  End If
If X >= 61 And X <= 90 Then
    picbalance.Cls
    dummy = BitBlt(picbalance.hdc, 0, 0, 30, 40, pallbuttons.hdc, 241, 0, SRCCOPY)
    dummy = BitBlt(picbalance.hdc, 31, 0, 30, 40, pallbuttons.hdc, 301, 0, SRCCOPY)
    dummy = BitBlt(picbalance.hdc, 61, 0, 30, 40, pallbuttons.hdc, 391, 0, SRCCOPY)
             
          
          
   cb = 3
  End If

End If
End Sub

Private Sub picpuse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dummy

If Button = 1 Then
  
   picpuse.Cls
    dummy = BitBlt(picpuse.hdc, 0, 0, 30, 40, pallbuttons.hdc, 211, 0, SRCCOPY)
End If
End Sub

Private Sub picpuse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dummy As Long

If Button = 1 Then
  
   picpuse.Cls
    dummy = BitBlt(picpuse.hdc, 0, 0, 30, 40, pallbuttons.hdc, 181, 0, SRCCOPY)
If mp9.CurrentPosition <= 0 Then Exit Sub
  mp9.Pause
End If
End Sub

Private Sub tm2_Timer()
txcu = Int(mp9.CurrentPosition)

End Sub

Private Sub tm1_Timer()
lb.Left = lb.Left - 50
If lb.Left < -lb.Width Then
lb.Left = pic1.ScaleWidth


End If
End Sub





Private Sub ipos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

c3 = 1
pos = X / 15
tm3.Enabled = False
If filename = "" Then Exit Sub

End Sub


Private Sub ipos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


If c3 = 1 Then
  If ipos.Left >= 275 Then
    ipos.Left = 275
    Exit Sub
  End If
  If ipos.Left <= 130 Then
    ipos.Left = 130
    Exit Sub
  End If
  ipos.Left = ipos.Left + ((X / 15) - pos)
  mp9.CurrentPosition = (mp9.Duration * ipos.Left) / 240
End If



End Sub


Private Sub ipos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

c3 = 0
tm1.Enabled = True
If filename = "" Then Exit Sub

End Sub

Private Sub tm3_Timer()

If mp9.CurrentPosition <= 0 Then
  Exit Sub
End If

vp1 = ((mp9.CurrentPosition / mp9.Duration)) * 100
vp2 = (Int(vp1 * 155) / 100)
ipos.Left = vp2 + 130


End Sub



















