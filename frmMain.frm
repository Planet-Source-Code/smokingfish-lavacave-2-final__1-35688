VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LavaCave"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3420
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":0E42
   MousePointer    =   99  'Custom
   ScaleHeight     =   2685
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picdummi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4000
      Left            =   1920
      ScaleHeight     =   3945
      ScaleWidth      =   3945
      TabIndex        =   33
      Top             =   2880
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   2760
      Top             =   360
   End
   Begin VB.PictureBox picenemymask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2400
      Picture         =   "frmMain.frx":170C
      ScaleHeight     =   180
      ScaleWidth      =   720
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   2400
      Picture         =   "frmMain.frx":1A9D
      ScaleHeight     =   2685
      ScaleWidth      =   3420
      TabIndex        =   31
      Top             =   2640
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   -240
      Picture         =   "frmMain.frx":2138
      ScaleHeight     =   2685
      ScaleWidth      =   3420
      TabIndex        =   30
      Top             =   2520
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Timer Timer4 
      Interval        =   3000
      Left            =   600
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Picture         =   "frmMain.frx":2F9D
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   27
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picPlayermask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2160
      Picture         =   "frmMain.frx":55BD
      ScaleHeight     =   180
      ScaleWidth      =   1080
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.ListBox lstY 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   23
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer3 
      Interval        =   150
      Left            =   600
      Top             =   120
   End
   Begin VB.PictureBox pic200mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      Picture         =   "frmMain.frx":601F
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picenemyslowmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      Picture         =   "frmMain.frx":6079
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picenemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2160
      Picture         =   "frmMain.frx":60CB
      ScaleHeight     =   180
      ScaleWidth      =   720
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox picenemyslow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      Picture         =   "frmMain.frx":67CD
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic200 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      Picture         =   "frmMain.frx":6883
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picPlayer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2160
      Picture         =   "frmMain.frx":6976
      ScaleHeight     =   180
      ScaleWidth      =   1080
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1320
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   1200
      Top             =   120
   End
   Begin VB.PictureBox pic200pmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      Picture         =   "frmMain.frx":6E1D
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pic500pmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      Picture         =   "frmMain.frx":6E77
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picspeedmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      Picture         =   "frmMain.frx":6ED5
      ScaleHeight     =   180
      ScaleWidth      =   300
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picwallsmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      Picture         =   "frmMain.frx":6F1A
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picspeedslowmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      Picture         =   "frmMain.frx":7260
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picwallbackmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      Picture         =   "frmMain.frx":72CF
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pickillmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      Picture         =   "frmMain.frx":7615
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pickill 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      Picture         =   "frmMain.frx":7684
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picwallback 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      Picture         =   "frmMain.frx":7723
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picspeedslow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      Picture         =   "frmMain.frx":77EF
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picWalls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      Picture         =   "frmMain.frx":78F6
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picSpeed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1320
      Picture         =   "frmMain.frx":79D1
      ScaleHeight     =   180
      ScaleWidth      =   300
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pic500p 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      Picture         =   "frmMain.frx":7AE6
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pic200p 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      Picture         =   "frmMain.frx":7C30
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin MediaPlayerCtl.MediaPlayer sound 
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
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
      CurrentPosition =   -1
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
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "End Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "-?-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   3525
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UpDown As Boolean
Dim UpDown2 As Boolean
Dim EnemyTrue1 As Boolean
Dim EnemyTrue2 As Boolean
Dim EnemyTrue3 As Boolean
Dim Score1 As Integer
Dim Score2 As Integer
Dim Track As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If Label1.Visible = True Then
Label1_Click
End If
End If
If KeyCode = vbKeyEscape Then
If Label1.Visible = True Then
Unload Me
End If
If Label1.Visible = False Then
Game.SnakeOver = True
End If
End If
End Sub

Private Sub Form_Load()
Game.BonusItem = 1
sound.AutoRewind = True
sound.AutoStart = True
sound.PlayCount = 999
sound.Filename = App.Path & "\mid\fast_like.Mid"
sound.Play
Track = 1
UpDown2 = True
Me.Picture = Picture1.Picture
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Game.EnemyFrame = 0
Game.SideMove = 0
Game.SnakeAcceleration = 0
Game.SnakeOver = False
Game.SnakePositionX = 0
Game.SnakePositionY = 0
Game.SnakeSpeed = 0
Game.WorldGravity = 0
Game.WorldWidth = 0
EnemyTrue1 = False
EnemyTrue2 = False
EnemyTrue3 = False
UpDown = False
Score1 = 0
Score2 = 0
Me.Show
End Sub

Public Sub MainLoop()
Dim TEMPa As Integer
Dim TEMPb As Integer
Do
DoEvents
Me.Cls
Game.SnakeSpeed = Game.SnakeSpeed + Game.WorldGravity

If Game.SnakePositionY + Game.SnakeSpeed + 180 > 2500 Then
    Game.SnakeSpeed = -Int(Game.SnakeSpeed * 0.5)
    Game.SnakePositionY = 2500 - 180
Else
    Game.SnakePositionY = Game.SnakePositionY + Game.SnakeSpeed
End If

If Game.SnakeDirection = True Then
    If Game.SnakePositionX + 230 + Game.SnakeAcceleration > 3200 Then
        Game.SnakeDirection = False
        Game.SnakePositionX = 3200 - 230
        If UpDown2 = True Then
        Game.SnakeAcceleration = Game.SnakeAcceleration + 4
        Else
        Game.SnakeAcceleration = Game.SnakeAcceleration - 4
        End If
        If Game.SnakeAcceleration = 52 Then UpDown2 = False
        If Game.SnakeAcceleration = 30 Then UpDown2 = True
    Else
    Game.SnakePositionX = Game.SnakePositionX + Game.SnakeAcceleration
    End If
End If

If Game.SnakeDirection = False Then
    If Game.SnakePositionX - Game.SnakeAcceleration < 254 Then
        Game.SnakeDirection = True
        Game.SnakePositionX = 254
        If UpDown2 = True Then
        Game.SnakeAcceleration = Game.SnakeAcceleration + 4
        Else
        Game.SnakeAcceleration = Game.SnakeAcceleration - 4
        End If
        If Game.SnakeAcceleration = 52 Then UpDown2 = False
        If Game.SnakeAcceleration = 30 Then UpDown2 = True
    Else
        Game.SnakePositionX = Game.SnakePositionX - Game.SnakeAcceleration
    End If
End If

If Game.SnakePositionY <= 254 Then
Game.SnakePositionY = 254
Game.SnakeSpeed = -Int(Game.SnakeSpeed * 0.5)
End If

If GetAsyncKeyState(vbKeySpace) Then
Game.SnakeSpeed = Game.SnakeSpeed - 4
End If

Game.DrawitemsShadow
Game.DrawEnemys
Game.DrawWorld
Game.CheckEnemys
Game.CheckItems

If frmMain.Point(Game.SnakePositionX, Game.SnakePositionY) = RGB(2, 2, 2) Then
If Game.BonusItem = 1 Then
ITEM1
GoTo OUT
End If
If Game.BonusItem = 2 Then
ITEM2
GoTo OUT
End If
If Game.BonusItem = 3 Then
ITEM3
GoTo OUT
End If
If Game.BonusItem = 4 Then
ITEM4
GoTo OUT
End If
If Game.BonusItem = 5 Then
ITEM5
GoTo OUT
End If
If Game.BonusItem = 6 Then
ITEM6
GoTo OUT
End If
If Game.BonusItem = 7 Then
ITEM7
GoTo OUT
End If
If Game.BonusItem = 8 Then
ITEM8
GoTo OUT
End If
If Game.BonusItem = 9 Then
ITEM9
GoTo OUT
End If
End If

If frmMain.Point(Game.SnakePositionX + 170, Game.SnakePositionY) = RGB(2, 2, 2) Then
If Game.BonusItem = 1 Then
ITEM1
GoTo OUT
End If
If Game.BonusItem = 2 Then
ITEM2
GoTo OUT
End If
If Game.BonusItem = 3 Then
ITEM3
GoTo OUT
End If
If Game.BonusItem = 4 Then
ITEM4
GoTo OUT
End If
If Game.BonusItem = 5 Then
ITEM5
GoTo OUT
End If
If Game.BonusItem = 6 Then
ITEM6
GoTo OUT
End If
If Game.BonusItem = 7 Then
ITEM7
GoTo OUT
End If
If Game.BonusItem = 8 Then
ITEM8
GoTo OUT
End If
If Game.BonusItem = 9 Then
ITEM9
GoTo OUT
End If
End If

If frmMain.Point(Game.SnakePositionX, Game.SnakePositionY + 170) = RGB(2, 2, 2) Then
If Game.BonusItem = 1 Then
ITEM1
GoTo OUT
End If
If Game.BonusItem = 2 Then
ITEM2
GoTo OUT
End If
If Game.BonusItem = 3 Then
ITEM3
GoTo OUT
End If
If Game.BonusItem = 4 Then
ITEM4
GoTo OUT
End If
If Game.BonusItem = 5 Then
ITEM5
GoTo OUT
End If
If Game.BonusItem = 6 Then
ITEM6
GoTo OUT
End If
If Game.BonusItem = 7 Then
ITEM7
GoTo OUT
End If
If Game.BonusItem = 8 Then
ITEM8
GoTo OUT
End If
If Game.BonusItem = 9 Then
ITEM9
GoTo OUT
End If
End If

If frmMain.Point(Game.SnakePositionX + 170, Game.SnakePositionY + 170) = RGB(2, 2, 2) Then
If Game.BonusItem = 1 Then
ITEM1
GoTo OUT
End If
If Game.BonusItem = 2 Then
ITEM2
GoTo OUT
End If
If Game.BonusItem = 3 Then
ITEM3
GoTo OUT
End If
If Game.BonusItem = 4 Then
ITEM4
GoTo OUT
End If
If Game.BonusItem = 5 Then
ITEM5
GoTo OUT
End If
If Game.BonusItem = 6 Then
ITEM6
GoTo OUT
End If
If Game.BonusItem = 7 Then
ITEM7
GoTo OUT
End If
If Game.BonusItem = 8 Then
ITEM8
GoTo OUT
End If
If Game.BonusItem = 9 Then
ITEM9
GoTo OUT
End If
End If

If Not frmMain.Point(Game.SnakePositionX, Game.SnakePositionY) = vbBlack Then
Game.SnakeOver = True
End If
If Not frmMain.Point(Game.SnakePositionX + 170, Game.SnakePositionY) = vbBlack Then
Game.SnakeOver = True
End If
If Not frmMain.Point(Game.SnakePositionX, Game.SnakePositionY + 170) = vbBlack Then
Game.SnakeOver = True
End If
If Not frmMain.Point(Game.SnakePositionX + 170, Game.SnakePositionY + 170) = vbBlack Then
Game.SnakeOver = True
End If
 
OUT:
'frmMain.PSet (Game.SnakePositionX, Game.SnakePositionY)
'frmMain.PSet (Game.SnakePositionX + 170, Game.SnakePositionY)
'frmMain.PSet (Game.SnakePositionX, Game.SnakePositionY + 170)
'frmMain.PSet (Game.SnakePositionX + 170, Game.SnakePositionY + 170)

Game.DrawSnake
Game.DrawSnakeShadow

Game.DrawItems


Score1 = Score1 + 1
frmMain.CurrentX = 1100
frmMain.CurrentY = 2200
frmMain.Print "Score: " & Score1
frmMain.CurrentX = 900
frmMain.CurrentY = -40
frmMain.Print "Highscore: " & Score2

'StretchBlt picGame.HDC, 0, 0, picGame.Width, picGame.Height, frmMain.HDC, 0, 0, 3500, 3100, SRCCOPY
Loop Until Game.SnakeOver = True
BitBlt picdummi.HDC, 0, 0, frmMain.Width, frmMain.Height, frmMain.HDC, 0, 0, SRCCOPY
frmMain.Cls
'For i = 1 To 80
'Sleep 5
'Game.GradientCircle frmMain, Game.SnakePositionX, Game.SnakePositionY, i * 5, 200, 50, 50, 5, False, False
'frmMain.Refresh
'Next i
BitBlt frmMain.HDC, 0, 0, frmMain.Width, frmMain.Height, picdummi.HDC, 0, 0, SRCCOPY
StopMP3
PlayMP3 App.Path & "\Sound\blast6.mp3"
For i = 1 To 10
Sleep 10
Game.GradientCircle frmMain, lstX.List(0), lstY.List(0), i * 5, 200, 50, 50, 3, True, False
frmMain.Refresh
Next i
For i = 1 To 20
Sleep 10
Game.GradientCircle frmMain, lstX.List(1), lstY.List(1), i * 5, 200, 50, 50, 3, True, False
frmMain.Refresh
Next i
For i = 1 To 30
Sleep 10
Game.GradientCircle frmMain, lstX.List(2), lstY.List(2), i * 5, 200, 50, 50, 3, True, False
frmMain.Refresh
Next i
For i = 1 To 40
Sleep 10
Game.GradientCircle frmMain, lstX.List(3), lstY.List(3), i * 5, 200, 50, 50, 3, True, False
frmMain.Refresh
Next i
For i = 1 To 50
Sleep 10
Game.GradientCircle frmMain, lstX.List(4), lstY.List(4), i * 5, 200, 50, 50, 3, True, False
frmMain.Refresh
Next i
For i = 1 To 60
Sleep 10
Game.GradientCircle frmMain, Game.SnakePositionX, Game.SnakePositionY, i * 5, 200, 50, 50, 3, True, False
frmMain.Refresh
Next i
Sleep 1000
frmMain.Cls
frmMain.CurrentX = 1100
frmMain.CurrentY = 500
frmMain.Print "Game Over!!"
frmMain.CurrentX = 1300
frmMain.CurrentY = 2200
frmMain.Print "Score: " & Score1
StopMP3
PlayMP3 App.Path & "\sound\kinderschrei6.mp3"
If Score2 < Score1 Then
Open App.Path & "\Score.Dat" For Binary As #1
Put #1, , Score1
Close #1
MsgBox "NEW HIGHSCORE!!!"
End If
Me.Refresh
Sleep 3000
'Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
Me.Cls
Form_Load
End Sub

Private Sub ITEM1()
Randomize Timer
StopMP3
PlayMP3 App.Path & "\sound\wizzloop1.mp3"
Score1 = Score1 + 200
Game.ItemY = 200
Game.ItemX = Int(Rnd * 2500) + 300
Game.BonusItem = Int(Rnd * 9) + 1
End Sub
Private Sub ITEM2()
Randomize Timer
StopMP3
PlayMP3 App.Path & "\sound\kinderschrei7.mp3"
Score1 = Score1 - 200
Game.ItemY = 200
Game.ItemX = Int(Rnd * 2500) + 300
Game.BonusItem = Int(Rnd * 9) + 1
End Sub
Private Sub ITEM3()
Randomize Timer
StopMP3
PlayMP3 App.Path & "\sound\wizzloop1.mp3"
Score1 = Score1 + 500
Game.ItemY = 200
Game.ItemX = Int(Rnd * 2500) + 300
Game.BonusItem = Int(Rnd * 9) + 1
End Sub
Private Sub ITEM4()
Randomize Timer
StopMP3
PlayMP3 App.Path & "\sound\wizzloop1.mp3"
'Game.WorldWidth = Game.WorldWidth + 20
If Game.WorldWidth < 180 Then
UpDown = False
End If
Game.ItemY = 200
Game.ItemX = Int(Rnd * 2500) + 300
Game.BonusItem = Int(Rnd * 9) + 1
End Sub
Private Sub ITEM5()
Randomize Timer
StopMP3
PlayMP3 App.Path & "\sound\kinderschrei7.mp3"
UpDown = True
'Game.WorldWidth = Game.WorldWidth - 10
Game.ItemY = 200
Game.ItemX = Int(Rnd * 2500) + 300
Game.BonusItem = Int(Rnd * 9) + 1
End Sub
Private Sub ITEM6()
Randomize Timer
StopMP3
PlayMP3 App.Path & "\sound\kinderschrei7.mp3"
Game.SnakeOver = True
Game.ItemY = 200
Game.ItemX = Int(Rnd * 2500) + 300
Game.BonusItem = Int(Rnd * 9) + 1
End Sub
Private Sub ITEM7()
Randomize Timer
StopMP3
PlayMP3 App.Path & "\sound\kinderschrei7.mp3"
Game.SnakeAcceleration = Game.SnakeAcceleration + 4
Game.ItemY = 200
Game.ItemX = Int(Rnd * 2500) + 300
Game.BonusItem = Int(Rnd * 9) + 1
End Sub
Private Sub ITEM8()
Randomize Timer
StopMP3
PlayMP3 App.Path & "\sound\wizzloop1.mp3"
If Game.SnakeAcceleration > 10 Then
Game.SnakeAcceleration = Game.SnakeAcceleration - 8
End If
Game.ItemY = 200
Game.ItemX = Int(Rnd * 2500) + 300
Game.BonusItem = Int(Rnd * 9) + 1
End Sub
Private Sub ITEM9()
Randomize Timer
StopMP3
PlayMP3 App.Path & "\sound\wizzloop1.mp3"
Game.EnemyStop = True
Timer5.Enabled = True
Game.ItemY = 200
Game.ItemX = Int(Rnd * 2500) + 300
Game.BonusItem = Int(Rnd * 9) + 1
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.Visible = True Then
Label1.BackColor = vbRed
Label2.BackColor = vbRed
Label3.BackColor = vbRed
Label4.BackColor = vbRed
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
StopMP3
sound.Stop
End
End Sub

Private Sub Label1_Click()
Me.Picture = Picture2.Picture
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
lstX.List(0) = 0
lstX.List(1) = 0
lstX.List(2) = 0
lstX.List(3) = 0
lstX.List(4) = 0
lstY.List(0) = 0
lstY.List(1) = 0
lstY.List(2) = 0
lstY.List(3) = 0
lstY.List(4) = 0
Score1 = 0
Score2 = 0
If FileExists(App.Path & "\Score.Dat") = True Then
Open App.Path & "\Score.Dat" For Binary As #1
Get #1, , Score2
Close #1
End If
Game.SnakePositionX = 400
Game.SnakePositionY = 1250
Game.SnakeSpeed = 0
Game.SnakeDirection = True
Game.WorldGravity = 1.6
Game.DummiWorld
Game.SnakeAcceleration = 10
Game.SideMove = 25
Game.WorldWidth = 150
UpDown = True
Call MainLoop
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.Visible = True Then
Label1.BackColor = &H404040
End If
End Sub

Private Sub Label2_Click()
MsgBox "2002 by SmokingFish!" & vbCrLf & "mail@SmokingFish.de" & vbCrLf & "Grafix by VaterUnser", vbInformation
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.Visible = True Then
Label2.BackColor = &H404040
End If
End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.Visible = True Then
Label3.BackColor = &H404040
End If
End Sub

Private Sub Label4_Click()
MsgBox "Use the Space Key to accelerate!" & vbCrLf & "Dont touch the Red and the Blue Wall or the Enemys!" & vbCrLf & "And remember , This Game is Truecolor only!", vbInformation
End Sub

Private Sub MediaPlayer1_DVDNotify(ByVal EventCode As Long, ByVal EventParam1 As Long, ByVal EventParam2 As Long)

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.Visible = True Then
Label4.BackColor = &H404040
End If
End Sub

Private Sub Timer1_Timer()
If UpDown = True Then
Game.WorldWidth = Game.WorldWidth - 1
Else
Game.WorldWidth = Game.WorldWidth + 1
End If
If Game.WorldWidth = 80 Then UpDown = False
If Game.WorldWidth = 180 Then UpDown = True
End Sub

Private Sub Timer3_Timer()
lstX.AddItem Game.SnakePositionX
lstY.AddItem Game.SnakePositionY
If lstX.ListCount = 6 Then
lstX.RemoveItem 0
lstY.RemoveItem 0
End If
If Game.EnemyFrame = 0 Then
Game.EnemyFrame = 1
Exit Sub
End If
If Game.EnemyFrame = 1 Then
Game.EnemyFrame = 2
Exit Sub
End If
If Game.EnemyFrame = 2 Then
Game.EnemyFrame = 3
Exit Sub
End If
If Game.EnemyFrame = 3 Then
Game.EnemyFrame = 0
Exit Sub
End If
End Sub

Private Sub Timer4_Timer()
If Track = 1 Then
PlayMP3 App.Path & "\Sound\loop1.mp3"
Track = 2
Exit Sub
End If
End Sub

Private Sub Timer5_Timer()
Game.EnemyStop = False
Timer5.Enabled = False
End Sub
