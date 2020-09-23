VERSION 5.00
Begin VB.Form frmSend 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   420
      TabIndex        =   4
      Top             =   300
      Width           =   3135
   End
   Begin VB.TextBox txtMessage 
      Height          =   2175
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   2820
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2460
      TabIndex        =   1
      Top             =   2820
      Width           =   1095
   End
   Begin MessengerProgram.CoolTitlebar BAR 
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   476
      Caption         =   " Send Message"
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TO:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   300
      Width           =   315
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BAR_CloseClick()
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
frmMain.WinPop.SendToWinPopUp frmMain.lblUser.Caption, UCase(Trim(txtUser.Text)), txtMessage.Text
Unload Me
End Sub

Private Sub Form_Load()
    BAR.PosistionButtons
    BAR.Parent_Hwnd = Me.hwnd
    Set BAR.Picture = frmMain.BAR.Picture
End Sub
