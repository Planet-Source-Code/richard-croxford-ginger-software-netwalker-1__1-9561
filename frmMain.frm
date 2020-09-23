VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4275
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
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
      Left            =   2460
      TabIndex        =   6
      ToolTipText     =   "About the program"
      Top             =   300
      Width           =   675
   End
   Begin MessengerProgram.WinPoPupEmulator WinPop 
      Left            =   1500
      Top             =   1080
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin MessengerProgram.cSysTray SysTray 
      Left            =   2100
      Top             =   1080
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "frmMain.frx":0442
      TrayTip         =   ""
   End
   Begin MessengerProgram.CoolTitlebar BAR 
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   476
      Caption         =   ""
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
      Picture         =   "frmMain.frx":0894
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
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
      Left            =   3180
      TabIndex        =   4
      ToolTipText     =   "Send a message to a computer"
      Top             =   300
      Width           =   555
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
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
      Left            =   3780
      TabIndex        =   3
      ToolTipText     =   "Close the application"
      Top             =   300
      Width           =   435
   End
   Begin VB.TextBox txtMessage 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMain.frx":1DD6E
      Top             =   600
      Width           =   4155
   End
   Begin VB.Label lblUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1140
      TabIndex        =   1
      Top             =   300
      Width           =   1275
   End
   Begin VB.Label STATIC0 
      Caption         =   "Current User"
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
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Private Sub BAR_CloseClick()
Me.WindowState = 1
Me.Visible = False
SysTray.InTray = True
End Sub

Private Sub BAR_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmSend.Top = Me.Top
    frmSend.Left = Me.Left + Me.Width
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub cmdQuit_Click()
If MsgBox("If you close the program you will not recieve any messages" + vbCrLf + "Are you sure you wish to continue", vbYesNo + vbQuestion, "Are you sure?") = vbYes Then
    WinPop.CloseSimulator
    End
End If
End Sub

Private Sub cmdSend_Click()
    frmSend.Show
    frmSend.Top = Me.Top
    frmSend.Left = Me.Left + Me.Width
End Sub

Private Sub Form_Load()
App.Title = "Ginger Software - NetWalker V" & App.Major & "." & App.Revision
SysTray.TrayTip = App.Title
BAR.Caption = App.Title
    BAR.PosistionButtons
    BAR.Parent_Hwnd = Me.hwnd
WinPop.Initialisation
If WinPop.MailSlotHandle = -1 Then
    MsgBox "There is a error with the network to fix this problem log off and back on", vbCritical
    WinPop.CloseSimulator
    'End
End If

    Dim szBuffer As String * 20
    GetUserName szBuffer, 20
    lblUser.Caption = Trim(szBuffer)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmSend.Top = Me.Top
    frmSend.Left = Me.Left + Me.Width
End Sub

Private Sub SysTray_MouseDblClick(Button As Integer, Id As Long)
    Me.Visible = True
    Me.WindowState = 0
    SysTray.InTray = False
    Me.SetFocus
End Sub

Private Sub WinPop_MessageWaiting(NbrMessageWaiting As Integer)
Beep
Beep

Dim x As Integer
SysTray_MouseDblClick 0, 0
For x = 1 To NbrMessageWaiting
    txtMessage.Text = "Message From " & WinPop.MessageFrom(1) & vbCrLf & WinPop.MessageText(1)
    WinPop.ClearMessage (1)
Next

Me.SetFocus
End Sub
