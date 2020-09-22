Attribute VB_Name = "mMain"
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

Public Const MyVersion As String = "0.60"
Public Const MyBuild As String = "042006"

Public AppPath As String
Public fMainForm As frmMain

Sub Main()

    '------------------------------------------------------
    'Set the application path
    AppPath = App.Path
    If right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
    '------------------------------------------------------
    '----------------------------------------------------------------
'    Dim fLogin As New frmLogin
'    fLogin.Show vbModal
'    If Not fLogin.OK Then
'        'Login Failed so exit app
'        End
'    End If
'    Unload fLogin

    'frmSplash.Show
    'frmSplash.Refresh
    '----------------------------------------------------------------
    Set fMainForm = New frmMain
    Load fMainForm
    'Unload frmSplash
    '----------------------------------------------------------------
    fMainForm.Show
    '----------------------------------------------------------------

End Sub

