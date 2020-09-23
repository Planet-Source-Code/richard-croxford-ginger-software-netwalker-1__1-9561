VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2085
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   195
      Left            =   3960
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
      ExtentX         =   661
      ExtentY         =   344
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MessengerProgram.CoolTitlebar BAR 
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   476
      Caption         =   " About"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconVisible     =   0   'False
      MinVisible      =   0   'False
      MaxVisible      =   0   'False
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      X1              =   1260
      X2              =   3480
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.come.to/croxford/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1260
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Click to go to my web site"
      Top             =   1800
      Width           =   2220
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This Software is freeware and we encorage you to distribute it"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1260
      Width           =   2715
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright Ginger Software 2000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1260
      TabIndex        =   3
      Top             =   900
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright Richard Croxford 2000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1260
      TabIndex        =   2
      Top             =   660
      Width           =   3195
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ginger Software NetWalker Message System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1260
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   60
      Picture         =   "frmAbout.frx":030A
      ToolTipText     =   "The Ginger Software Logo"
      Top             =   360
      Width           =   1140
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4560
      Y1              =   270
      Y2              =   270
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BAR_CloseClick()
Unload Me
End Sub


Private Sub Form_Load()
BAR.Parent_Hwnd = Me.hwnd
BAR.PosistionButtons
Set BAR.Picture = frmMain.BAR.Picture
End Sub

Private Sub lblURL_Click()
    WB.Navigate2 "http://www.come.to/croxford/", , "_New"
End Sub
