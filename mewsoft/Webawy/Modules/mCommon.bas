Attribute VB_Name = "mCommon"
'==========================================================
'           Copyright Information
'==========================================================
'Program Name: Mewsoft Visual Html Editor
'Program Author   : Elsheshtawy, A. A.
'Home Page        : http://www.mewsoft.com
'Copyrights © 2006 Mewsoft Corporation. All rights reserved.
'==========================================================
'==========================================================
Option Explicit

Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Public MemoryInfo As MEMORYSTATUS

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function IsUserAnAdmin Lib "shell32" () As Long

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Public Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

' Declaration for Stay on Top sub
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

' Win32 Declarations for INI Access
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCapture Lib "user32" () As Long

'====================================================================
Public Sub AltLVBackground(lv As ListView, _
    ByVal BackColorOne As OLE_COLOR, _
    ByVal BackColorTwo As OLE_COLOR)
'---------------------------------------------------------------------------------
' Purpose   : Alternates row colors in a ListView control
' Method    : Creates a picture box and draws the desired color scheme in it, then
'             loads the drawn image as the listviews picture.
'---------------------------------------------------------------------------------
Dim lH      As Long
Dim lSM     As Byte
Dim picAlt  As PictureBox
Dim lvParent

    Set lvParent = lv.Parent
    With lv
        
        If .View = lvwReport And .ListItems.Count Then
            Set picAlt = lvParent.Controls.Add("VB.PictureBox", "picAlt")
            lSM = .Parent.ScaleMode
            .Parent.ScaleMode = vbTwips
            .PictureAlignment = lvwTile
            lH = .ListItems(1).Height
            With picAlt
                .BackColor = BackColorOne
                .AutoRedraw = True
                .Height = lH * 2
                .BorderStyle = 0
                .Width = 10 * Screen.TwipsPerPixelX
                picAlt.Line (0, lH)-(.ScaleWidth, lH * 2), BackColorTwo, BF
                Set lv.Picture = .Image
            End With
            Set picAlt = Nothing
            lvParent.Controls.Remove "picAlt"
            lv.Parent.ScaleMode = lSM
        End If
    End With
    Set lvParent = Nothing
End Sub
'====================================================================
Public Function AddProgBar(pb As ProgressBar, sb As StatusBar, lPan As Long)
    ' make sure that when the form is resized that the
    ' statusbar is resized before we continue
    sb.Align = 2
    sb.Refresh
    
    ' set the properties of the progressbar
    ' flat with no border seems to look the best
    ' also set the progressbar to the top of the zorder
    pb.ZOrder 0
    pb.Appearance = ccFlat
    pb.BorderStyle = ccNone
    
    ' now resize the progressbar1 to fit in the statusbar panel
    pb.left = sb.Panels(lPan).left + 25
    pb.Width = sb.Panels(lPan).Width - 45
    pb.top = sb.top + 45
    pb.Height = sb.Height - 75
End Function

' Important note: This works in SDI only not in MDI forms
Public Sub ShowProgressInStatusBar(pb As ProgressBar, sb As StatusBar, _
                    ByVal Pan As Long, ByVal bShowProgressBar As Boolean)

    Dim tRC As RECT
    
    If bShowProgressBar Then
        ' Get the size of the Panel lPan Rectangle from the status bar
        ' remember that Indexes in the API are always 0 based (well,
        ' nearly always) - therefore Panel(2) = Panel(1) to the api
        SendMessageAny sb.hwnd, SB_GETRECT, Pan - 1, tRC
        ' and convert it to twips....
        With tRC
            .top = (.top * Screen.TwipsPerPixelY)
            .left = (.left * Screen.TwipsPerPixelX)
            .bottom = (.bottom * Screen.TwipsPerPixelY) - .top
            .right = (.right * Screen.TwipsPerPixelX) - .left
        End With
        ' Now Reparent the ProgressBar to the statusbar
        With pb
            SetParent .hwnd, sb.hwnd
            .Move tRC.left, tRC.top, tRC.right, tRC.bottom
            .Visible = True
            .Value = 0
        End With
        
    Else
        '
        ' Reparent the progress bar back to the form and hide it
        '
        SetParent pb.hwnd, sb.Parent
        pb.Visible = False
    End If
    
End Sub

Public Function AlreadyRunning() As Boolean
    'Check for previously loaded instances of the program'''''''''''
    If App.PrevInstance = True Then
        Beep
        MsgBox "Program cancelled." & vbCr & vbCr & _
            "There is a previous copy of " & App.Title & " already running.  " & vbCr & "Please" _
            & " check your currently running applications in " & vbCr & "the task manager and try again." _
              & vbCr & vbCr & "(Task Manager:  Ctrl+Alt+Del)", vbOKOnly + vbExclamation, App.Title & " already running!"
        End
        AlreadyRunning = True
    End If
End Function

Public Function GetMDIChildCount() As Integer
    Dim frm As Form
    Dim cnt As Integer
    
    For Each frm In Forms
        If (frm.MDIChild And frm.Visible) Then cnt = cnt + 1
    Next frm
    GetMDIChildCount = cnt
    'numFormsOpen = Forms.Count
  End Function
Function CleanCrLf(ByVal Text As String) As String
    
    Dim sText As String
    
    sText = Text
    sText = Replace(sText, vbCrLf, "")
    sText = Replace(sText, vbLf, "")
    sText = Replace(sText, vbCr, "")
    CleanCrLf = sText

End Function

Public Sub OnTop(TheForm As Form)
    '** Description:
    '** Put window on top
    SetWindowPos TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub NotOnTop(TheForm As Form)
    '** Description:
    '** Remove window from top
    SetWindowPos TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Function ReadINI(Section As String, Key As String, Optional sDefault As String)
    '** Description:
    '** Get settings from ini file
    Dim sRet As String
    ' Fill sRet with Null Chars
    sRet = String(255, Chr(0))
    ' Get data from INI file
    ReadINI = left(sRet, GetPrivateProfileString(Section, Key, sDefault, sRet, Len(sRet), App.Path & "\Webawy.ini"))
End Function

Public Sub WriteINI(Section As String, Key As String, Value As String)
    '** Description:
    '** Write settings to ini file
    ' Write to INI file
    WritePrivateProfileString Section, Key, Value, App.Path & "\Webawy.ini"
End Sub

'-----------------------------------------------------------------------------------------
Public Function IsInIDE() As Boolean
   Dim x As Long
   Debug.Assert Not TestIDE(x)
   IsInIDE = x = 1
End Function
Public Function TestIDE(x As Long) As Boolean
   x = 1
End Function
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------

