VERSION 5.00
Begin VB.UserControl WinPoPupEmulator 
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   3075
   ToolboxBitmap   =   "WinPoPupEmulator.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1860
      Top             =   1440
   End
   Begin VB.Image imaImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   0
      Picture         =   "WinPoPupEmulator.ctx":0312
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "WinPoPupEmulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
  
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Const MAILSLOT_WAIT_FOREVER = (-1)
Const OPEN_EXISTING = 3
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Const GENERIC_EXECUTE = &H20000000
Const GENERIC_ALL = &H10000000
Const INVALID_HANDLE_VALUE = -1
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Private Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFileName As Long, ByVal lpBuff As Any, ByVal nNrBytesToWrite As Long, lpNrOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwAccess As Long, ByVal dwShare As Long, ByVal lpSecurityAttrib As Long, ByVal dwCreationDisp As Long, ByVal dwAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateMailslot Lib "kernel32.dll" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Dim MHandle As Long
'Default Property Values:
Const m_def_MailSlotHandle = 0
Const m_def_MessageFrom = ""
Const m_def_MessageText = ""
'Property Variables:
Dim m_MailSlotHandle As Long
Dim m_MessageFrom() As String
Dim m_MessageText() As String
'Event Declarations:
Event MessageWaiting(NbrMessageWaiting As Integer)
Function SendToWinPopUp(PopFrom As String, PopTo As String, MsgText As String) As Long
    Dim rc As Long
    Dim mshandle As Long
    Dim msgtxt As String
    Dim byteswritten As Long
    Dim mailslotname As String
    ' name of the mailslot
    mailslotname = "\\" + PopTo + "\mailslot\messngr"
    msgtxt = PopFrom + Chr(0) + PopTo + Chr(0) + MsgText + Chr(0)
    mshandle = CreateFile(mailslotname, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, -1)
    rc = WriteFile(mshandle, msgtxt, Len(msgtxt), byteswritten, 0)
    rc = CloseHandle(mshandle)
End Function
Private Function ReadToWinPoPup()
Dim NextSize As Long
Dim Waiting As Long
Dim Buffer() As Byte
Dim ReadSize As Long
Dim FHandle As Long
Dim TempsWaiting As Long
Dim tempo As String
Dim x As Integer
Dim passe As Integer
Dim FromTo As String
Dim message As String
TempsWaiting = MAILSLOT_WAIT_FOREVER
' look for message
 FHandle = GetMailslotInfo(MHandle, 10, NextSize, Waiting, TempsWaiting)
 
 'if message go read this
 If Waiting <> 0 Then
    If m_MessageFrom(UBound(m_MessageFrom)) <> "" Then
        ReDim Preserve m_MessageFrom(UBound(m_MessageFrom) + 1)
        ReDim Preserve m_MessageText(UBound(m_MessageText) + 1)
    End If
    ReDim Buffer(NextSize)
    FHandle = ReadFile(MHandle, Buffer(1), NextSize, ReadSize, ByVal 0&)
    passe = 1
    For x = 1 To UBound(Buffer)
        If Buffer(x) <> 0 Then
            tempo = tempo & Chr(Buffer(x))
        Else
            Select Case passe
                Case 1
                    m_MessageFrom(UBound(m_MessageFrom)) = tempo
                    passe = 2
                Case 2
                    passe = 3
                Case 3
                    m_MessageText(UBound(m_MessageText)) = tempo
            End Select
            tempo = ""
        End If
    Next
    RaiseEvent MessageWaiting(UBound(m_MessageFrom))
End If
End Function
 
Private Sub Timer1_Timer()
    ReadToWinPoPup
End Sub
Public Sub Initialisation()
Dim MaxMessage As Long
Dim MesssageTimer As Long
Dim t As SECURITY_ATTRIBUTES
    t.nLength = Len(t)
    t.bInheritHandle = False
    MaxMessage = 0
    MesssageTimer = MAILSLOT_WAIT_FOREVER
    MHandle = CreateMailslot("\\.\mailslot\messngr", MaxMessage, MesssageTimer, t)
    ReDim m_MessageFrom(1)
    ReDim m_MessageText(1)
    m_MessageFrom(1) = m_def_MessageFrom
    m_MessageText(1) = m_def_MessageText
    m_MailSlotHandle = MHandle
    Timer1.Enabled = True
End Sub

Public Property Get MessageFrom(index As Integer) As String
    MessageFrom = m_MessageFrom(index)
End Property
Public Property Get MessageText(index As Integer) As String
    MessageText = m_MessageText(index)
End Property
Public Sub ClearMessage(index As Integer)
Dim tempo() As String
Dim tempo2() As String
Dim x As Integer
Dim y As Integer
Dim z As Integer
    If UBound(m_MessageFrom) = 1 Then
        m_MessageFrom(1) = ""
        m_MessageText(1) = ""
        Exit Sub
    End If
    ReDim tempo(UBound(m_MessageFrom) - 1)
    ReDim tempo2(UBound(m_MessageFrom) - 1)
    y = 1
    z = 1
    For x = 1 To UBound(m_MessageFrom)
        If x <> index Then
            MsgBox "10"
            MsgBox UBound(m_MessageFrom)
            MsgBox "y=" & y
            tempo(z) = m_MessageFrom(y)
            MsgBox "11"
            tempo2(z) = m_MessageText(y)
            MsgBox "12"
            z = z + 1
        End If
        MsgBox "13"
        y = y + 1
        MsgBox "14"
    Next
    MsgBox "15"
    ReDim m_MessageFrom(UBound(tempo))
    MsgBox "16"
    ReDim m_MessageText(UBound(tempo2))
    MsgBox "17"
    For x = 1 To UBound(tempo)
    MsgBox "18"
        m_MessageFrom(x) = tempo(x)
        MsgBox "19"
        m_MessageText(x) = tempo2(x)
        MsgBox "20"
    Next
End Sub
Public Sub CloseSimulator()
     Timer1.Enabled = False
     CloseHandle MHandle
End Sub

Public Property Get MailSlotHandle() As Long
    MailSlotHandle = m_MailSlotHandle
End Property

    

Private Sub UserControl_Resize()
UserControl.Width = imaImage.Width
UserControl.Height = imaImage.Height
End Sub
