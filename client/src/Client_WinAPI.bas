Attribute VB_Name = "Client_WinAPI"
Option Explicit

' This module contains calls to the Windows API for various reasons

' //WIN32API Function
Private Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Boolean

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CYCAPTION = 4
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

Private Const ABM_GETTASKBARPOS = &H5

' General calls
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

' API Declares
Public myHWnd As Long

Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetTopWindow Lib "user32" _
                (ByVal hWnd As Integer) As Integer
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" _
                (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Public Declare Function IsWindowVisible Lib "user32" _
                (ByVal hWnd As Integer) As Boolean
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                (ByVal hWnd As Integer, ByVal lpString As String, _
                 ByVal cch As Integer) As Integer

Public Declare Function ZCompress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Function ZUncompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Public Const WU_LOGPIXELSX = 88
Public Const WU_LOGPIXELSY = 90

Public Const FLASHW_STOP = 0    'Stop flashing. The system restores the window to its original state.

Public Const FLASHW_CAPTION = &H1    'Flash the window caption.

Public Const FLASHW_TRAY = &H2    'Flash the taskbar button.

Public Const FLASHW_ALL = (FLASHW_CAPTION Or FLASHW_TRAY)    'Flash both the window caption and taskbar button. This is equivalent to setting the FLASHW_CAPTION Or FLASHW_TRAY flags.

Public Const FLASHW_TIMER = &H4    'Flash continuously, until the FLASHW_STOP flag is set.

Public Const FLASHW_TIMERNOFG = &HC    'Flash continuously until the window comes to the foreground.

'''''''''''''''''''''''''''''''''''''''''''''''
' KEYBOARD INPUT
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''
' ACCESS TOKEN RELATED JAZZ (Check for admin and stuff)
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Const TOKEN_READ As Long = &H20008 ' Used to read the token data
Private Const TOKEN_INFO_CLASS_TokenElevation As Long = 20 ' Used to check whether token is elevated or not

'''''''''''''''''''''''''''''''''''''''''''''''
' HIGH RESOLUTION TIMER
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private GetSystemTimeOffset As Currency

'''''''''''''''''''''''''''''''''''''''''''''''
' FORM WINDOW RELATED
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClipBox Lib "gdi32" (ByVal hDC As Long, pRect As ClipBoxRect) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, pRect As ClipBoxRect) As Long
Private Declare Function EqualRect Lib "user32" (rc1 As ClipBoxRect, rc2 As ClipBoxRect) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, hDC As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Type ClipBoxRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type FLASHWINFO
    cbSize As Long
    hWnd As Long
    dwFlags As Long
    uCount As Long
    dwTimeout As Long
End Type

Private Type APPBARDATA
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type

Public IsFlashing As Boolean

Public Function GetTaskBarHeight() As Long
    Dim ABD As APPBARDATA

    SHAppBarMessage ABM_GETTASKBARPOS, ABD
    GetTaskBarHeight = ABD.rc.Bottom - ABD.rc.Top
End Function

Public Function GetTaskBarWidth() As Long
    Dim ABD As APPBARDATA

    SHAppBarMessage ABM_GETTASKBARPOS, ABD
    GetTaskBarWidth = ABD.rc.Right - ABD.rc.Left
End Function

Property Get TitleBarHeight() As Long
    TitleBarHeight = GetSystemMetrics(SM_CYCAPTION)
End Property

Public Function IsOnTop(ByVal hWnd As Integer) As Boolean
    Dim i As Integer
    Dim X As Integer
    Dim S As String
    
    X = 1
    i = GetTopWindow(0)

    ' Enumeration
    Do
        i = GetNextWindow(i, 2)  ' Find next window in Z-order
        If i = hWnd Then
            Exit Do
        Else
            If i = 0 Then        ' Never find any window match the input handle
                IsOnTop = False
            End If
        End If

        If IsWindowVisible(i) = True Then
            S = Space(256)
            If GetWindowText(i, S, 255) <> 0 Then
            ' Very important to prevent confusing
            ' of BalloonTips and ContextMenuStrips
                X = X + 1
            End If
        End If
    Loop

    ' x is Z-order number

    If X = 1 Then
        IsOnTop = True
    Else
        IsOnTop = False
    End If
End Function

Public Sub InitTime()
    
    ' Set the high-resolution timer
    timeBeginPeriod 1
    
    ' Get the initial time, time starting from this point will be calculated relative to this value
    GetSystemTime GetSystemTimeOffset

End Sub

Public Function getTime() As Single

    ' The roll over still happens but the advantage is that you don't have to restart your pc, just restart the server
    ' This is getTimeCount starts counting from when your PC has started, but this method starts counting from when the server has started
    
    Dim CurrentTime As Currency

    ' Grab the current time (we have to pass a variable ByRef instead of a function return like the other timers)
    GetSystemTime CurrentTime

    ' Calculate the difference between the 64-bit times, return as a 32-bit time
    getTime = CurrentTime - GetSystemTimeOffset
    

End Function

' Returns the version of Windows that the user is running
Public Function GetWindowsVersion() As String
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = "Win32s on Windows 3.1"
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = "Windows NT"

                Select Case osv.dwVerMajor
                    Case 3
                        GetWindowsVersion = "Windows NT 3.5"
                    Case 4
                        GetWindowsVersion = "Windows NT 4.0"
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows 2000"
                            Case 1
                                GetWindowsVersion = "Windows XP"
                            Case 2
                                GetWindowsVersion = "Windows Server 2003"
                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows Vista/Server 2008"
                            Case 1
                                GetWindowsVersion = "Windows 7/Server 2008 R2"
                            Case 2
                                GetWindowsVersion = "Windows 8 and 10"
                        End Select
                End Select

            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        GetWindowsVersion = "Windows 95"
                    Case 90
                        GetWindowsVersion = "Windows Me"
                    Case Else
                        GetWindowsVersion = "Windows 98"
                End Select
        End Select
    Else
        GetWindowsVersion = "Unable to identify your version of Windows."
    End If
End Function

Public Function IsElevatedAccess() As Boolean
    Dim PID As Long, hToken As Long, Elevated As Long, ReturnLength As Long
    
    PID = GetCurrentProcess ' = -1 always, it's a value that when passed to API maps to the current, correct PID
    
    If OpenProcessToken(PID, TOKEN_READ, hToken) Then
        Call GetTokenInformation(hToken, TOKEN_INFO_CLASS_TokenElevation, Elevated, LenB(Elevated), ReturnLength)
        IsElevatedAccess = Not (Elevated = 0) ' if Elevated = 0 then not elevated
        Call CloseHandle(hToken)
    End If
End Function

Public Function IsWindowObscured(ByVal hWnd As Long) As Boolean
    Dim hDC As Long, rcClip As ClipBoxRect, rcClient As ClipBoxRect

    IsWindowObscured = True

    If hWnd Then
        hDC = GetDC(hWnd)
        If hDC Then
            Select Case GetClipBox(hDC, rcClip)
            Case 0    ' NULLREGION, i.e fully covered
                IsWindowObscured = True
            Case 1  ' SIMPLEREGION,
                Call GetClientRect(hWnd, rcClient)
                If EqualRect(rcClient, rcClip) Then
                    IsWindowObscured = False    ' Fully visible
                Else
                    'IsWindowObscured = False    ' Partially visible
                End If
            Case 2    ' COMPLEXREGION
                IsWindowObscured = False    ' Partially Visible
            Case Else:
                IsWindowObscured = False    ' No fucking clue what happened
            End Select

            Call ReleaseDC(hWnd, hDC)
        End If
    End If
End Function

Public Sub FlashWindow(ByRef Window As Form)
    Dim FlashInfo As FLASHWINFO

    IsFlashing = Not IsFlashing

    With FlashInfo
        .cbSize = Len(FlashInfo)
        If IsFlashing Then
            .dwFlags = FLASHW_TRAY Or FLASHW_TIMER
        Else
            .dwFlags = FLASHW_STOP
        End If
        .dwTimeout = 0
        .hWnd = Window.hWnd
        .uCount = 0
    End With

    FlashWindowEx FlashInfo
End Sub



