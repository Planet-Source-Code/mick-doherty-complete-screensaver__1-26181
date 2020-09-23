Attribute VB_Name = "modScrnSave"
Option Explicit

Private Declare Function VerifyScreenSavePwd Lib "password.cpl" (ByVal hwnd&) As Boolean
Private Declare Function PwdChangePassword Lib "mpr" Alias "PwdChangePasswordA" (ByVal lpcRegkeyname As String, ByVal hwnd As Long, ByVal uiReserved1 As Long, ByVal uiReserved2 As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long)

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const SPI_SCREENSAVERRUNNING = 97&

Private Const WS_CHILD = &H40000000
Private Const GWL_STYLE = (-16)
Private Const GWL_HWNDPARENT = (-8)

Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_SHOWWINDOW = &H40

Public Const VER_PLATFORM_WIN32_NT = 2

Public ThumbView As Boolean
Public rctThumb As RECT
Public OSVer As OSVERSIONINFO

Sub Main()
    
    On Error GoTo ErrorHandle
    
    GetVersionEx OSVer
    
    Select Case LCase(Left(Command, 2))
    
        Case "/s"   'Screen Saver run or Preview selected
            If Not AlreadyRunning Then
                SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                'WIN NT handles it's own security
                If OSVer.dwPlatformId <> VER_PLATFORM_WIN32_NT Then
                    SystemParametersInfo SPI_SCREENSAVERRUNNING, 1, 0&, 0&
                End If
                ShowCursor 0
                frmMain.Show
            Else
                End
            End If
            Exit Sub
       
        Case "/p"   'Thumbnail Preview
            'Command here was /p xxxx where xxxx is the
            'handle to the thumbnail window(SSDemoParent).
            If GetSetting("ScrnSave Base", "Settings", "Preview", "0") = 1 Then
                MiniMe frmMain, CLng(Right(Command, 4))
                Exit Sub
            Else
                MiniMe frmThumb, CLng(Right(Command, 4))
            End If
            
            
        Case "/a"   'Change Password selected
            PwdChangePassword "SCRSAVE", frmMain.hwnd, 0&, 0&
            
        Case "/c"   'Configure selected
            frmSetup.Show
            
        Case Else   'Configure selected from files right click menu.
            frmSetup.Show
            
    End Select
    
    Unload frmMain
    
Exit Sub

ErrorHandle:
    'I've not seen this, but just in case.
    MsgBox Err.Number & vbCrLf & Err.Description
    Err.Clear
    Unload frmMain
    
End Sub

Public Function EndScreenSaver(hWndForm As Long) As Boolean
    'Win NT handles it's own security, so if not Win NT
    'Check Password. Returns True if Password left blank.
    If OSVer.dwPlatformId <> VER_PLATFORM_WIN32_NT Then
        If VerifyScreenSavePwd(hWndForm) Then
            SystemParametersInfo SPI_SCREENSAVERRUNNING, 0, 0&, 0&
        Else
            Exit Function
        End If
    End If
    
    EndScreenSaver = True
    ShowCursor 1
    
End Function

Private Sub MiniMe(frm As Form, hThumb As Long)
    'Make the form a Child of SSDemoParent.
    Dim lStyle As Long
    
    On Error GoTo ErrorHandle
    
    ThumbView = True
    
    GetClientRect hThumb, rctThumb
    
    lStyle = GetWindowLong(frm.hwnd, GWL_STYLE)
    lStyle = lStyle Or WS_CHILD
    SetWindowLong frm.hwnd, GWL_STYLE, lStyle
    
    SetParent frm.hwnd, hThumb
    SetWindowLong frm.hwnd, GWL_HWNDPARENT, hThumb
    SetWindowPos frm.hwnd, HWND_TOP, 0, 0, rctThumb.Right, rctThumb.Bottom, SWP_SHOWWINDOW
    
Exit Sub

ErrorHandle:
    'I've not seen this, but just in case.
    MsgBox Err.Number & vbCrLf & Err.Description
    Err.Clear
    Unload frmMain
    
End Sub

Private Function AlreadyRunning() As Boolean
    'We can't use App.PrevInstance as there is a previous instance running in the
    'Thumbnail preview window. Besides that, It apparently causes probs with NT.
    'FindWindow does not search Child Windows, so will not find the Thumbnail Window.
    If FindWindow(vbNullString, "VBSCRNSAVE") <> 0 Then
        AlreadyRunning = True
    End If
    
End Function

