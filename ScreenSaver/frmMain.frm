VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "VBSCRNSAVE"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   554
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3600
      Top             =   3480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'It is important that this form has a caption of "VBSCRNSAVE", so that we may
'prevent multiple instances of the program from running. You can change the caption
'so long as this change is updated in modScrnSave AlreadyRunning().

'Remove the code for Timer1 and wipe frmSetup, then save this project
'as a Base Template for your future Screensavers.

Private iMoves

Private Sub Form_Click()
    If Not ThumbView Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Not ThumbView Then Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If ThumbView Then Exit Sub
    'Windows sends a MouseMove message when the ScreenSaver starts, and when the
    'activate period is reached and a second instance attempts to load.
    'So here we will check for actual Mouse Movement.
    iMoves = iMoves + 1
    If iMoves > 3 Then Unload Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If not a thumbnail Preview then check for Password
    If ThumbView Then Exit Sub
    If EndScreenSaver(Me.hwnd) Then Unload Me Else Cancel = True: iMoves = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Unload all forms
    Dim frm As Form
    
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
    
End Sub

Private Sub Timer1_Timer()
    'For your viewing pleasure.
    'OK so it's not very exciting, but it's just for the demo.
    Dim iMax As Integer, iShape As Integer
    
    On Error GoTo ErrorHandle
    
    Timer1.Interval = CInt(GetSetting("ScrnSave Base", "Settings", "Interval", "100"))
    iMax = CInt(GetSetting("ScrnSave Base", "Settings", "Max", "20"))
    iMax = Me.ScaleHeight / 100 * iMax
    iShape = GetSetting("ScrnSave Base", "Settings", "Shape", "0")
    
    Randomize
    Me.FillColor = RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
    Me.FillStyle = vbSolid
    Me.ScaleMode = vbPixels
    
    If iShape = 2 Then iShape = Int(2 * Rnd)
    
    Select Case iShape
        Case 0
            Circle (Rnd * ScaleWidth, Rnd * ScaleHeight), Rnd * (iMax \ 2), Me.FillColor
        Case Else
            Dim StartX As Single, StartY As Single, iSize As Integer
            StartX = Rnd * ScaleWidth: StartY = Rnd * ScaleHeight: iSize = Rnd * iMax
            Line (StartX, StartY)-(StartX + iSize, StartY + iSize), Me.FillColor, BF
    End Select

Exit Sub

ErrorHandle:
    'Just in case
    Err.Clear
    Unload Me
End Sub

