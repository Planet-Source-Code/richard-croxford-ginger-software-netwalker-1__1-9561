VERSION 5.00
Begin VB.UserControl CoolTitlebar 
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   PaletteMode     =   4  'None
   Picture         =   "CoolTitlebar.ctx":0000
   ScaleHeight     =   54
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ToolboxBitmap   =   "CoolTitlebar.ctx":1AE8
   Begin VB.Image cmdClose 
      Appearance      =   0  'Flat
      Height          =   210
      Index           =   2
      Left            =   915
      Picture         =   "CoolTitlebar.ctx":1DFA
      Top             =   525
      Width           =   240
   End
   Begin VB.Image cmdClose 
      Appearance      =   0  'Flat
      Height          =   210
      Index           =   1
      Left            =   915
      Picture         =   "CoolTitlebar.ctx":20DC
      Top             =   315
      Width           =   240
   End
   Begin VB.Image cmdMax 
      Height          =   210
      Index           =   2
      Left            =   600
      Picture         =   "CoolTitlebar.ctx":23BE
      Top             =   525
      Width           =   240
   End
   Begin VB.Image cmdMax 
      Height          =   210
      Index           =   1
      Left            =   600
      Picture         =   "CoolTitlebar.ctx":26A0
      Top             =   315
      Width           =   240
   End
   Begin VB.Image cmdRestore 
      Height          =   210
      Index           =   2
      Left            =   330
      Picture         =   "CoolTitlebar.ctx":2982
      Top             =   525
      Width           =   240
   End
   Begin VB.Image cmdRestore 
      Height          =   210
      Index           =   1
      Left            =   330
      Picture         =   "CoolTitlebar.ctx":2C64
      Top             =   315
      Width           =   240
   End
   Begin VB.Image cmdMin 
      Height          =   210
      Index           =   2
      Left            =   60
      Picture         =   "CoolTitlebar.ctx":2F46
      Top             =   525
      Width           =   240
   End
   Begin VB.Image cmdMin 
      Height          =   210
      Index           =   1
      Left            =   60
      Picture         =   "CoolTitlebar.ctx":3228
      Top             =   315
      Width           =   240
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titlebar Caption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   270
      TabIndex        =   0
      Top             =   45
      Width           =   1140
   End
   Begin VB.Image imaIcon 
      Height          =   240
      Left            =   15
      Picture         =   "CoolTitlebar.ctx":350A
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image cmdMax 
      Height          =   210
      Index           =   0
      Left            =   4845
      Picture         =   "CoolTitlebar.ctx":394C
      Top             =   30
      Width           =   240
   End
   Begin VB.Image cmdMin 
      Height          =   210
      Index           =   0
      Left            =   4605
      Picture         =   "CoolTitlebar.ctx":3C2E
      Top             =   30
      Width           =   240
   End
   Begin VB.Image cmdRestore 
      Height          =   210
      Index           =   0
      Left            =   4845
      Picture         =   "CoolTitlebar.ctx":3F10
      Top             =   30
      Width           =   240
   End
   Begin VB.Image cmdClose 
      Appearance      =   0  'Flat
      Height          =   210
      Index           =   0
      Left            =   5145
      Picture         =   "CoolTitlebar.ctx":41F2
      Top             =   30
      Width           =   240
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "&Move"
      End
      Begin VB.Menu mnuSize 
         Caption         =   "&Size"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Mi&nimize"
      End
      Begin VB.Menu mnuMaximize 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "CoolTitlebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2


'Default Property Values:
'Const m_def_SysMenu = 0
Const m_def_Parent_Hwnd = 0
Const m_def_IconVisible = True
Const m_def_MinVisible = True
Const m_def_MaxVisible = True
Const m_def_RestoreVisible = False
Const m_def_CloseVisible = True
'Property Variables:
Dim m_MouseIcon As Picture
'Dim m_SysMenu As Menu
Dim m_Parent_Hwnd As Long
Dim m_IconVisible As Boolean
Dim m_MinVisible As Boolean
Dim m_MaxVisible As Boolean
Dim m_RestoreVisible As Boolean
Dim m_CloseVisible As Boolean
'Event Declarations:
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event BarDblClick() 'MappingInfo=lblCaption,lblCaption,-1,DblClick
Attribute BarDblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MinClick()
Event MaxClick()
Event RestoreClick()
Event CloseClick()
Event IconClick() 'MappingInfo=imaIcon,imaIcon,-1,Click
Attribute IconClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
'Event MinClick() 'MappingInfo=cmdMin(0),cmdMin,0,Click
'Event MaxClick() 'MappingInfo=cmdMax(0),cmdMax,0,Click
'Event RestoreClick() 'MappingInfo=cmdRestore(0),cmdRestore,0,Click
'Event CloseClick() 'MappingInfo=cmdClose(0),cmdClose,0,Click
'Event IconClick()

Sub PosistionButtons()
    cmdClose(0).Left = UserControl.ScaleWidth - 18
    cmdClose(0).Visible = m_CloseVisible
    
    If m_CloseVisible Then
        cmdRestore(0).Left = UserControl.ScaleWidth - 36
        cmdRestore(0).Visible = m_RestoreVisible
    Else
        cmdRestore(0).Left = UserControl.ScaleWidth - 18
        cmdRestore(0).Visible = m_RestoreVisible
    End If
    
    If m_CloseVisible Then
        cmdMax(0).Left = UserControl.ScaleWidth - 36
        cmdMax(0).Visible = m_MaxVisible
    Else
        cmdMax(0).Left = UserControl.ScaleWidth - 18
        cmdMax(0).Visible = m_MaxVisible
    End If
    
    cmdMin(0).Left = cmdMax(0).Left - 16
    cmdMin(0).Visible = m_MinVisible
    
    imaIcon.Visible = m_IconVisible
    
    If m_IconVisible Then
        lblCaption.Left = imaIcon.Left + 17
    Else
        lblCaption.Left = 1
    End If
End Sub




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get CaptionFont() As Font
Attribute CaptionFont.VB_Description = "Returns a Font object."
    Set CaptionFont = lblCaption.Font
End Property

Public Property Set CaptionFont(ByVal New_CaptionFont As Font)
    Set lblCaption.Font = New_CaptionFont
    PropertyChanged "CaptionFont"
End Property
'
'Private Sub cmdMin_Click(Index As Integer)
'    RaiseEvent MinClick
'End Sub
'
'Private Sub cmdMax_Click(Index As Integer)
'    RaiseEvent MaxClick
'End Sub
'
'Private Sub cmdRestore_Click(Index As Integer)
'    RaiseEvent RestoreClick
'End Sub
'
'Private Sub cmdClose_Click(Index As Integer)
'    RaiseEvent CloseClick
'End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get IconVisible() As Boolean
    IconVisible = m_IconVisible
    PosistionButtons
End Property

Public Property Let IconVisible(ByVal New_IconVisible As Boolean)
    m_IconVisible = New_IconVisible
    PropertyChanged "IconVisible"
    PosistionButtons
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MinVisible() As Boolean
    MinVisible = m_MinVisible
    PosistionButtons
End Property

Public Property Let MinVisible(ByVal New_MinVisible As Boolean)
    m_MinVisible = New_MinVisible
    PropertyChanged "MinVisible"
    PosistionButtons
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MaxVisible() As Boolean
    MaxVisible = m_MaxVisible
    PosistionButtons
End Property

Public Property Let MaxVisible(ByVal New_MaxVisible As Boolean)
    m_MaxVisible = New_MaxVisible
    PropertyChanged "MaxVisible"
    PosistionButtons
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get RestoreVisible() As Boolean
    RestoreVisible = m_RestoreVisible
    PosistionButtons
End Property

Public Property Let RestoreVisible(ByVal New_RestoreVisible As Boolean)
    m_RestoreVisible = New_RestoreVisible
    PropertyChanged "RestoreVisible"
    PosistionButtons
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get CloseVisible() As Boolean
    CloseVisible = m_CloseVisible
    PosistionButtons
End Property

Public Property Let CloseVisible(ByVal New_CloseVisible As Boolean)
    m_CloseVisible = New_CloseVisible
    PropertyChanged "CloseVisible"
    PosistionButtons
End Property




Private Sub cmdClose_Click(index As Integer)
    RaiseEvent CloseClick
End Sub

Private Sub cmdClose_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set cmdClose(0).Picture = cmdClose(2).Picture
End Sub
Private Sub cmdClose_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set cmdClose(0).Picture = cmdClose(1).Picture
End Sub

Private Sub cmdMax_Click(index As Integer)
    RaiseEvent MaxClick
End Sub

Private Sub cmdMax_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set cmdMax(0).Picture = cmdMax(2).Picture
End Sub
Private Sub cmdMax_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set cmdMax(0).Picture = cmdMax(1).Picture
End Sub

Private Sub cmdMin_Click(index As Integer)
    RaiseEvent MinClick
End Sub

Private Sub cmdMin_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set cmdMin(0).Picture = cmdMin(2).Picture
End Sub

Private Sub cmdMin_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Set cmdMin(0).Picture = cmdMin(1).Picture
End Sub

Private Sub cmdRestore_Click(index As Integer)
    RaiseEvent RestoreClick
End Sub




Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage m_Parent_Hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent BarDblClick
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_IconVisible = m_def_IconVisible
    m_MinVisible = m_def_MinVisible
    m_MaxVisible = m_def_MaxVisible
    m_RestoreVisible = m_def_RestoreVisible
    m_CloseVisible = m_def_CloseVisible
    m_Parent_Hwnd = m_def_Parent_Hwnd
    
    Set m_MouseIcon = LoadPicture("")
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    ReleaseCapture
    SendMessage m_Parent_Hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblCaption.Caption = PropBag.ReadProperty("Caption", "Titlebar Caption")
    Set lblCaption.Font = PropBag.ReadProperty("CaptionFont", Ambient.Font)
    m_IconVisible = PropBag.ReadProperty("IconVisible", m_def_IconVisible)
    m_MinVisible = PropBag.ReadProperty("MinVisible", m_def_MinVisible)
    m_MaxVisible = PropBag.ReadProperty("MaxVisible", m_def_MaxVisible)
    m_RestoreVisible = PropBag.ReadProperty("RestoreVisible", m_def_RestoreVisible)
    m_CloseVisible = PropBag.ReadProperty("CloseVisible", m_def_CloseVisible)

    m_Parent_Hwnd = PropBag.ReadProperty("Parent_Hwnd", m_def_Parent_Hwnd)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
'    Set m_SysMenu = PropBag.ReadProperty("SysMenu", mnu)
    Set m_MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
End Sub

Private Sub UserControl_Resize()
    PosistionButtons
    
    UserControl.Height = 270
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Titlebar Caption")
    Call PropBag.WriteProperty("CaptionFont", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("IconVisible", m_IconVisible, m_def_IconVisible)
    Call PropBag.WriteProperty("MinVisible", m_MinVisible, m_def_MinVisible)
    Call PropBag.WriteProperty("MaxVisible", m_MaxVisible, m_def_MaxVisible)
    Call PropBag.WriteProperty("RestoreVisible", m_RestoreVisible, m_def_RestoreVisible)
    Call PropBag.WriteProperty("CloseVisible", m_CloseVisible, m_def_CloseVisible)
    Call PropBag.WriteProperty("Parent_Hwnd", m_Parent_Hwnd, m_def_Parent_Hwnd)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    'Call PropBag.WriteProperty("SysMenu", m_SysMenu, m_def_SysMenu)
    Call PropBag.WriteProperty("MouseIcon", m_MouseIcon, Nothing)
End Sub

Private Sub imaIcon_Click()
    RaiseEvent IconClick
    PopupMenu mnu
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Parent_Hwnd() As Long
    Parent_Hwnd = m_Parent_Hwnd
End Property

Public Property Let Parent_Hwnd(ByVal New_Parent_Hwnd As Long)
    m_Parent_Hwnd = New_Parent_Hwnd
    PropertyChanged "Parent_Hwnd"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get SysMenu() As Menu
'    SysMenu = m_SysMenu
'End Property
'
'Public Property Let SysMenu(ByVal New_SysMenu As Menu)
'    m_SysMenu = New_SysMenu
'    PropertyChanged "SysMenu"
'End Property

Private Sub lblCaption_DblClick()
    RaiseEvent BarDblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function About()
    frmAbout.Show vbModal
End Function

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = m_MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set m_MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

