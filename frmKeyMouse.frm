VERSION 5.00
Begin VB.Form frmKeyMouse 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3015
   Icon            =   "frmKeyMouse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrToggle 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer tmrCheck 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Press  ~  to switch modes"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2775
      Begin VB.Label optTM 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Text Mode"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label optMM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mouse Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mouse 2 (ALT)"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mouse 3 (CTRL)"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mouse 1 (SPACE)"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmKeyMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'           ***************************************************
'           *                 Keyboard Mouse                  *
'           *      By Mike Plaehn (LCSBSSRHXXX) 4/27/04       *
'           ***************************************************

Option Explicit
'[Type PointAPI For Mouse Position And Mouse Distance]
Private Type POINTAPI
    X As Long
    Y As Long
End Type
'[Type NotifyIconData For Tray Icon]
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'[Tray Constants]
Const NIM_ADD = &H0 'Add to Tray
Const NIM_MODIFY = &H1 'Modify Details
Const NIM_DELETE = &H2 'Remove From Tray
Const NIF_MESSAGE = &H1 'Message
Const NIF_ICON = &H2 'Icon
Const NIF_TIP = &H4 'TooTipText
Const WM_MOUSEMOVE = &H200 'On Mousemove
Const WM_LBUTTONDBLCLK = &H203 'Left Double Click
Const WM_RBUTTONDOWN = &H204 'Right Button Down
Const WM_RBUTTONUP = &H205 'Right Button Up
Const WM_RBUTTONDBLCLK = &H206 'Right Double Click
'[Mouse Constants]
Const MOUSEEVENTF_LEFTDOWN = &H2 'Mouse 1 Down
Const MOUSEEVENTF_LEFTUP = &H4 'Mouse 1 Up
Const MOUSEEVENTF_RIGHTDOWN = &H8 'Mouse 2 Down
Const MOUSEEVENTF_RIGHTUP = &H10 'Mouse 2 Up
Const MOUSEEVENTF_MIDDLEDOWN = &H20 'Mouse Wheel Down
Const MOUSEEVENTF_MIDDLEUP = &H40 'Mouse Wheel Up
Const MOUSEEVENTF_MOVE = &H1 'Move
'[API]
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'[Variables For Tray Icon]
Dim TrayIco As NOTIFYICONDATA
Dim InTray As Boolean
Public Function CurLeft() As Boolean
    CurLeft = CBool(GetAsyncKeyState(37))
End Function
Public Function CurRight() As Boolean
    CurRight = CBool(GetAsyncKeyState(39))
End Function
Public Function CurUp() As Boolean
    CurUp = CBool(GetAsyncKeyState(38))
End Function
Public Function CurDown() As Boolean
    CurDown = CBool(GetAsyncKeyState(40))
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case InTray
        Case True
            'if me is in in tray and you click me then
            If Button = 1 Then
                'restore the form
                Me.WindowState = vbNormal
                'show the form
                Me.Show
            End If
        'if me isn't in the try and you click me then
        Case False
            'exit sub
            Exit Sub
    End Select
End Sub
Private Sub optMM_Click()
    tmrCheck.Enabled = True
    tmrToggle.Enabled = False
    optMM.BackColor = &HC0FFC0
    optTM.BackColor = &H8000000F
End Sub
Private Sub optTM_Click()
    tmrToggle.Enabled = True
    tmrCheck.Enabled = False
    optTM.BackColor = &HC0FFC0
    optMM.BackColor = &H8000000F
End Sub
Private Sub tmrCheck_Timer()
Dim Posit As POINTAPI
Dim keyresult As Integer
Dim cButt As Long
Dim dwEI As Long

    'if you press ` or ~ then call optTM_Click
    keyresult = GetAsyncKeyState(192)
    If keyresult = -32767 Then Call optTM_Click

    '[Get Cursor Position]
    'get the cursor position
    GetCursorPos Posit
    'display the cursor position on the form's caption
    Me.Caption = "(X: " & Posit.X & ", Y:" & Posit.Y & ")"
    '[Up, Down)
    'if your pressing the up arrow then go up
    If CurUp = True Then Call SetCursorPos(Posit.X, Posit.Y - 10)
    'if your pressing the down arrow then go down
    If CurDown = True Then Call SetCursorPos(Posit.X, Posit.Y + 10)
    '[Left, Right]
    'if your pressing the left arrow then go left
    If CurLeft = True Then Call SetCursorPos(Posit.X - 10, Posit.Y)
    'if your pressing the right arrow then go right
    If CurRight = True Then Call SetCursorPos(Posit.X + 10, Posit.Y)
    '[Diagonal Up]
    'if your pressing the up arrow and the left arrow then go up left
    If CurUp = True And CurLeft = True Then Call SetCursorPos(Posit.X - 10, Posit.Y - 10)
    'if your pressing the up arrow and the right arrow then go up right
    If CurUp = True And CurRight = True Then Call SetCursorPos(Posit.X + 10, Posit.Y - 10)
    '[Diagonal Down]
    'if your pressing the down arrow and the left arrow then go down left
    If CurDown = True And CurLeft = True Then Call SetCursorPos(Posit.X - 10, Posit.Y + 10)
    'if your pressing the down arrow and the right arrow then go down right
    If CurDown = True And CurRight = True Then Call SetCursorPos(Posit.X + 10, Posit.Y + 10)
    '[Mouse 1]
    'if you press space bar then
    keyresult = GetAsyncKeyState(32)
    If keyresult = -32767 Then
        'left click
        'mouse event left mouse down
        mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, cButt, dwEI
        'mouse event left mouse up
        mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, cButt, dwEI
    End If
    '[Mouse 2]
    'if you press alt then
    keyresult = GetAsyncKeyState(18)
    If keyresult = -32767 Then
        'right click
        'mouse event right mouse down
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0&, 0&, cButt, dwEI
        'mouse event right mouse up
        mouse_event MOUSEEVENTF_RIGHTUP, 0&, 0&, cButt, dwEI
    End If
    '[Mouse 3]
    'if you press ctrl then
    keyresult = GetAsyncKeyState(17)
    If keyresult = -32767 Then
        'middle click
        'mouse event middle mouse down
        mouse_event MOUSEEVENTF_MIDDLEDOWN, 0&, 0&, cButt, dwEI
        'mouse event middle mouse up
        mouse_event MOUSEEVENTF_MIDDLEUP, 0&, 0&, cButt, dwEI
    End If
    
    '[Tray]
    'if minimize
    If Me.WindowState = 1 Then
        'make variable InTray = true
        InTray = True
        'Hide frmMain
        Me.Hide
        With TrayIco
            .cbSize = Len(TrayIco)
            'tray icon hwnd = me.hwnd
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            'call back message on mouse move
            .uCallBackMessage = WM_MOUSEMOVE
            'tray icon
            .hIcon = Me.Icon
            'tray ToolTipText
            .szTip = "Keyboard Mouse" & vbNullChar
        End With
        'add tray icon with the properties of TrayIcon
        Shell_NotifyIcon NIM_ADD, TrayIco
    Else
        'remove from tray if the window is not minimized or in the tray
        InTray = False
        Shell_NotifyIcon NIM_DELETE, TrayIco
    End If
End Sub
Private Sub tmrToggle_Timer()
Dim Posit As POINTAPI
Dim keyresult As Integer
    
    'if you press ` or ~ then call optMM_Click
    keyresult = GetAsyncKeyState(192)
    If keyresult = -32767 Then Call optMM_Click
    
    '[Get Cursor Position]
    'get the cursor position
    GetCursorPos Posit
    'display the cursor position on the form's caption
    Me.Caption = "(X: " & Posit.X & ", Y:" & Posit.Y & ")"
End Sub
