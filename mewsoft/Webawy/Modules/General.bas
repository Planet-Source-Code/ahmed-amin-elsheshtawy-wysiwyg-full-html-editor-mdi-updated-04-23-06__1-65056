Attribute VB_Name = "General"
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
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

' Sub to sleep x seconds
Public Sub Sleep(lngSleep As Long)
   Dim lngSleepEnd As Long
   lngSleepEnd = GetTickCount + lngSleep * 1000
   While GetTickCount <= lngSleepEnd
      DoEvents
   Wend
End Sub

' Sub to freeze x seconds
Public Sub Freeze(lngFreeze As Long)
   Dim lngFreezeEnd As Long
   lngFreezeEnd = GetTickCount + lngFreeze * 1000
   While GetTickCount <= lngFreezeEnd
   Wend
End Sub
Public Function TrimNull(sString As String) As String
    TrimNull = left(sString, InStr(1, sString, vbNullChar) - 1)
End Function

'Sort string arrays
Sub SortArray(inpArray())
    Dim intRet
    Dim intCompare
    Dim intLoopTimes
    Dim strTemp
    
    For intLoopTimes = 1 To UBound(inpArray)
        For intCompare = LBound(inpArray) To UBound(inpArray) - 1
            intRet = StrComp(inpArray(intCompare), _
                     inpArray(intCompare + 1), vbTextCompare)
    
            If intRet = 1 Then 'String1 is greater than String2
                strTemp = inpArray(intCompare)
                inpArray(intCompare) = inpArray(intCompare + 1)
                inpArray(intCompare + 1) = strTemp
            End If
        Next
    Next

End Sub

' For put a windows in the middle of the screen
' FrmChild  = Windows to center
' FrmParent = MDI Windows (Optional)
Public Sub CenterForm(FrmChild As Form, Optional FrmParent As Variant)
    Dim iTop As Integer, iLeft As Integer
    
    If Not IsMissing(FrmParent) Then
        iTop = FrmParent.top + (FrmParent.ScaleHeight - FrmChild.Height) \ 2
        iLeft = FrmParent.left + (FrmParent.ScaleWidth - FrmChild.Width) \ 2
    Else
        iTop = (Screen.Height - FrmChild.Height) \ 2
        iLeft = (Screen.Width - FrmChild.Width) \ 2
    End If
    If iTop And iLeft Then
        FrmChild.Move iLeft, iTop
    End If
End Sub

Function NoNulo(Vrx As Variant) As String
    If IsNull(Vrx) Then
        NoNulo = ""
    Else
        NoNulo = Vrx
    End If
End Function

Public Function DirExists(ByVal sDir As String) As Boolean

    On Error GoTo ERR_Handler
    Dim strDir As String

    strDir = Dir(sDir, vbDirectory)

    If (strDir = "") Then
         'If it doesn't exist, create it
        CreateDirectoryStruct sDir
    End If
    
    DirExists = True
    Exit Function

ERR_Handler:
    DirExists = False
End Function

Public Function FileExists(ByVal sFile As String) As Boolean
  Dim lLength As Long

  If sFile <> vbNullString Then
    On Error Resume Next
    lLength = Len(Dir$(sFile))
    On Error GoTo err_routine
    FileExists = (Not Err And lLength > 0)
  Else
    FileExists = False
  End If

exit_routine:
  Exit Function

err_routine:
  FileExists = False
  Resume exit_routine

End Function

Public Sub CreateDirectoryStruct(ByVal CreateThisPath As String)

    On Error GoTo ERR_Handler
    'do initial check
    Dim RET As Boolean
    Dim Temp As String
    Dim ComputerName As String
    Dim IntoItCount As Integer
    Dim x As Integer
    Dim WakeString As String
    Dim MadeIt As Integer

    If Dir$(CreateThisPath, vbDirectory) <> "" Then Exit Sub
    'is this a network path?

    If left$(CreateThisPath, 2) = "\\" Then ' this is a UNC NetworkPath
        'must extract the machine name first, th
        '     en get to the first folder
        IntoItCount = 3
        ComputerName = Mid$(CreateThisPath, IntoItCount, InStr(IntoItCount, CreateThisPath, "\") - IntoItCount)
        IntoItCount = IntoItCount + Len(ComputerName) + 1
        IntoItCount = InStr(IntoItCount, CreateThisPath, "\") + 1
        'temp = Mid$(CreateThisPath, IntoItCount
        '     , x)
    Else ' this is a regular path
        IntoItCount = 4
    End If
    WakeString = left$(CreateThisPath, IntoItCount - 1)
    'start a loop through the CreateThisPath
    '     string

    Do
        x = InStr(IntoItCount, CreateThisPath, "\")

        If x <> 0 Then
            x = x - IntoItCount
            Temp = Mid$(CreateThisPath, IntoItCount, x)
        Else
            Temp = Mid$(CreateThisPath, IntoItCount)
        End If
        IntoItCount = IntoItCount + Len(Temp) + 1
        Temp = WakeString + Temp
        'Create a directory if it doesn't alread
        '     y exist
        RET = (Dir$(Temp, vbDirectory) <> "")


        If Not RET Then
            'ret& = CreateDirectory(temp, Security)
            MkDir Temp
        End If
        IntoItCount = IntoItCount 'track where we are in the String
        WakeString = left$(CreateThisPath, IntoItCount - 1)
    Loop While WakeString <> CreateThisPath

    Exit Sub

ERR_Handler:
    Err.Raise Err.Number
End Sub

Private Sub ClearDirectory(ByVal psDirName As String)
    'This function attempts to delete all fi
    '     les
    'and subdirectories of the given
    'directory name, and leaves the given
    'directory intact, but completely empty.
    '
    '
    'If the Kill command generates an error
    '     (i.e.
    'file is in use by another process -
    'permission denied error), then that fil
    '     e and
    'subdirectory will be skipped, and the
    'program will continue (On Error Resume
'     Next).
'
'EXAMPLE CALL:
' ClearDirectory "C:\Temp\"
Dim sSubDir


If Len(psDirName) > 0 Then


    If right(psDirName, 1) <> "\" Then
        psDirName = psDirName & "\"
    End If
    'Attempt to remove any files in director
    '     y
    'with one command (if error, we'll
    'attempt to delete the files one at a
    'time later in the loop):
    On Error Resume Next
    Kill psDirName & "*.*"


    DoEvents
        
        sSubDir = Dir(psDirName, vbDirectory)


        Do While Len(sSubDir) > 0
            'Ignore the current directory and the
            'encompassing directory:
            If sSubDir <> "." And _
            sSubDir <> ".." Then
            'Use bitwise comparison to make
            'sure MyName is a directory:
            If (GetAttr(psDirName & sSubDir) And _
            vbDirectory) = vbDirectory Then
            'Use recursion to clear files
            'from subdir:
            ClearDirectory psDirName & _
            sSubDir & "\"
            'Remove directory once files
            'have been cleared (deleted)
            'from it:
            RmDir psDirName & sSubDir


            DoEvents
                'ReInitialize Dir Command
                'after using recursion:
                sSubDir = Dir(psDirName, vbDirectory)
            Else
                'This file is remaining because
                'most likely, the Kill statement
                'before this loop errored out
                'when attempting to delete all
                'the files at once in this
                'directory. This attempt to
                'delete a single file by itself
                'may work because another
                '(locked) file within this same
                'directory may have prevented
                '(non-locked) files from being
                'deleted:
                Kill psDirName & sSubDir
                sSubDir = Dir
            End If
        Else
            sSubDir = Dir
        End If
    Loop
End If
End Sub


Public Sub ColorListviewRow(lv As ListView, RowNbr As Long, RowColor As OLE_COLOR)
'***************************************************************************
'Purpose: Color a ListView Row
'Inputs : lv - The ListView
'         RowNbr - The index of the row to be colored
'         RowColor - The color to color it
'Outputs: None
'***************************************************************************
    Dim itmX As ListItem
    Dim lvSI As ListSubItem
    Dim intIndex As Integer
    
    On Error GoTo ErrorRoutine
    
    Set itmX = lv.ListItems(RowNbr)
    
    itmX.ForeColor = RowColor
    For intIndex = 1 To lv.ColumnHeaders.Count - 2
        Set lvSI = itmX.ListSubItems(intIndex)
        lvSI.ForeColor = RowColor
    Next

    Set itmX = Nothing
    Set lvSI = Nothing
    Exit Sub

ErrorRoutine:

    MsgBox "List view set color error: " & Err.Description

End Sub

Public Function ReadFile1(Filename As String) As String
    
    On Error GoTo ERR_Handler
    
    Dim fs As Object
    Dim f As Object
    Dim Text As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(Filename, 1, 0)
    Text = f.ReadAll
    f.Close
    Set fs = Nothing
    Set f = Nothing
    ReadFile1 = Text
    Exit Function
    
ERR_Handler:
    Set fs = Nothing
    Set f = Nothing
    ReadFile1 = ""
End Function

Public Function ReadFile(Filename As String) As String
    
    On Error GoTo ERR_Handler
    
    Dim fs As FileSystemObject
    Dim ts As TextStream
    Dim Text As String
    
    Set fs = New FileSystemObject
    Set ts = fs.OpenTextFile(Filename)
    Text = ts.ReadAll
    Set ts = Nothing
    Set fs = Nothing
    ReadFile = Text
    
    Exit Function
    
ERR_Handler:
    Set ts = Nothing
    Set fs = Nothing
    ReadFile = ""
    
End Function
Public Function QualifyPath(sPath As String) As String

  'assures that a passed path ends in a slash
   If right$(sPath, 1) <> "\" Then
        QualifyPath = sPath & "\"
   Else
        QualifyPath = sPath
   End If
      
End Function

Public Sub ColorHtml(ByRef Text1 As Control)
    
    'C:\VBProjects\GroupD\HTML-Highlighter9
    Dim TagregEx, Match, Matches
    Dim tagPNregEx, Match2, Matches2
    Dim rtfstart As Long
    
    Set TagregEx = New RegExp
    
    TagregEx.Pattern = "<(.)[^> ]*( ){0,1}[^>]*>"
    
    TagregEx.IgnoreCase = False
    TagregEx.Global = True

    Set tagPNregEx = New RegExp
    
    tagPNregEx.Pattern = "(\w+ *=) *(\d+|""[^""]+"")"

    tagPNregEx.IgnoreCase = False
    tagPNregEx.Global = True
   
    rtfstart = Text1.SelStart
    
    If Text1.SelLength < 1 Then
        Exit Sub
    End If
    
    Set Matches = TagregEx.Execute(Text1.SelText)
    
    For Each Match In Matches
        If Match.Value <> "" Then
            Text1.SelStart = rtfstart + Match.FirstIndex
            Text1.SelLength = Match.length
            Text1.SelColor = vbRed
            
            If Match.SubMatches(0) = "!" Then
                Text1.SelColor = &H8000&
                GoTo nextmatch
            ElseIf Match.SubMatches(1) <> " " Then
                GoTo nextmatch
            End If
            
            Set Matches2 = tagPNregEx.Execute(Match.Value)
            
            For Each Match2 In Matches2
                If Match2.Value <> "" Then
                    Text1.SelStart = Match.FirstIndex + rtfstart + Match2.FirstIndex
                    Text1.SelLength = Match2.length
                    Text1.SelColor = &H8000&
                    Text1.SelLength = Len(Match2.SubMatches(0))
                    Text1.SelColor = vbBlue
                End If
            Next
            
        End If
nextmatch:
    Next
     'Label15.Caption = Matches.Count & " Tags"
End Sub

Public Function QualifyURL(ByVal strURL As String) As String
    
    Dim URL As String
    
    URL = Trim(LCase(strURL))
    
    If URL = "" Then
        QualifyURL = ""
        Exit Function
    End If
    
    If InStr(1, URL, "http://", vbTextCompare) = 1 Or _
        InStr(1, URL, "https://", vbTextCompare) = 1 Or _
        InStr(1, URL, "ftp://", vbTextCompare) = 1 Or _
        InStr(1, URL, "file://", vbTextCompare) = 1 Or _
        InStr(1, URL, "gopher://", vbTextCompare) = 1 Or _
        InStr(1, URL, "wais:", vbTextCompare) = 1 Or _
        InStr(1, URL, "telnet:", vbTextCompare) = 1 Or _
        InStr(1, URL, "mailto:", vbTextCompare) = 1 Or _
        InStr(1, URL, "news:", vbTextCompare) = 1 _
    Then
        QualifyURL = Trim(strURL)
    Else
        QualifyURL = "http://" & Trim(strURL)
    End If
    
End Function


