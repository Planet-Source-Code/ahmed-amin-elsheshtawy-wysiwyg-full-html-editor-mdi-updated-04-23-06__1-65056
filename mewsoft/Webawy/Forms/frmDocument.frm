VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "dhtmled.ocx"
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9735
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   9735
   Begin VB.TextBox Editawy1 
      Height          =   1935
      Left            =   4140
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2820
      Visible         =   0   'False
      Width           =   2955
   End
   Begin DHTMLEDLibCtl.DHTMLEdit DHTMLEdit1 
      Height          =   1575
      Left            =   1080
      TabIndex        =   0
      Top             =   2940
      Width           =   1695
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================
'           Copyright Information
'==========================================================
'Program Name: Mewsoft Visual Html Editor
'Program Author   : Elsheshtawy, A. A.
'Home Page        : http://www.mewsoft.com
'Copyrights © 2006 Mewsoft Corporation. All rights reserved.
'==========================================================
'==========================================================

Private Sub Form_Activate()
    
    On Error Resume Next
    If fMainForm.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ActiveDocument = Me.Tag
    fMainForm.Toolbar3.Buttons("Normal").Value = IIf(EditMode(CStr(ActiveDocument)) = NormalMode, tbrPressed, tbrUnpressed)
    fMainForm.Toolbar3.Buttons("HTML").Value = IIf(EditMode(CStr(ActiveDocument)) = HtmlMode, tbrPressed, tbrUnpressed)
    fMainForm.Toolbar3.Buttons("Preview").Value = IIf(EditMode(CStr(ActiveDocument)) = PreviewMode, tbrPressed, tbrUnpressed)
    
    fMainForm.tabFilesBar.Tabs.Item("F" & CStr(ActiveDocument)).Selected = True
    
End Sub

Private Sub Form_Deactivate()

    'fMainForm.DisableToolbar

End Sub

Private Sub Form_GotFocus()

   ActiveDocument = CLng(Me.Tag)

End Sub
Private Sub Form_Load()
    
    On Error GoTo ErrHandler
    
    DHTMLEditInitialized = False
    '----------------------------------------------------------------
    Me.WindowState = GetSetting(App.Title, "Settings", "WindowState", vbMaximized)
    DhtmlSnapToGrid = GetSetting(App.Title, "Settings", "DhtmlSnapToGrid", 0)
    DhtmlSnapToGridX = GetSetting(App.Title, "Settings", "DhtmlSnapToGridX", 5)
    DhtmlSnapToGridY = GetSetting(App.Title, "Settings", "DhtmlSnapToGridY", 5)
    DhtmlShowAll = GetSetting(App.Title, "Settings", "DhtmlShowAll", 0)
    DhtmlShowBorders = GetSetting(App.Title, "Settings", "DhtmlShowBorders", 1)
    '----------------------------------------------------------------
    ' hide the source code editor
    'Editawy1.Visible = False
    
    ' display the visual html editor
    DHTMLEdit1.Visible = True
    DHTMLEdit1.ShowBorders = DhtmlShowBorders
    DHTMLEdit1.ShowDetails = DhtmlShowAll
    DHTMLEdit1.SnapToGrid = DhtmlSnapToGrid
    DHTMLEdit1.SnapToGridX = DhtmlSnapToGridX
    DHTMLEdit1.SnapToGridY = DhtmlSnapToGridY
    '----------------------------------------------------------------
    'Editawy1.BackColor
    'Editawy1.ConfFile = "lexer.xml"
    'Editawy1.LineNumbers = True
    'Editawy1.language = "HTML"
    'Editawy1.MatchBraces = True
    'Editawy1.IndGuidesVisible = True
    'Editawy1.Column
    'Editawy1.TotalLines
    'Editawy1.SetFocus
    '----------------------------------------------------------------
    DHTMLEdit1.NewDocument
    DHTMLEdit1.DocumentHTML = GetNewDocHTML
    'DHTMLEdit1.CurrentDocumentPath

    'InitToolbarTable
    '
    'fMainForm.DisableToolbar
    
    'HTMLEdit1.ScrollBars = True
    'HTMLEdit1.ScrollbarAppearance = DEAPPEARANCE_FLAT
    DHTMLEdit1.SourceCodePreservation = True
    DHTMLEdit1.baseUrl = ""
    '----------------------------------------------------------------
    'DHTMLEdit1.BrowseMode = True    ' Clear the Undo/Redo buffer
    DHTMLEdit1.BrowseMode = False   ' Switch to edit mode
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    fMainForm.Toolbar2.Buttons("Save").Enabled = True
    fMainForm.Toolbar2.Buttons("SaveAs").Enabled = True
    
    fMainForm.Toolbar2.Buttons("ShowAll").Enabled = True
    fMainForm.Toolbar2.Buttons("ShowBorders").Enabled = True
    fMainForm.Toolbar3.Buttons("SnapToGrid").Enabled = True
    
    fMainForm.Toolbar3.Buttons("Normal").Enabled = True
    fMainForm.Toolbar3.Buttons("HTML").Enabled = True
    fMainForm.Toolbar3.Buttons("Preview").Enabled = True
    fMainForm.Toolbar3.Buttons("Normal").Value = tbrPressed
    
    fMainForm.Toolbar2.Buttons("Form").Enabled = True
    fMainForm.Toolbar3.Buttons("Textbox").Enabled = True
    fMainForm.Toolbar3.Buttons("Textarea").Enabled = True
    fMainForm.Toolbar3.Buttons("Checkbox").Enabled = True
    fMainForm.Toolbar3.Buttons("OptionButton").Enabled = True
    fMainForm.Toolbar3.Buttons("ListBox").Enabled = True
    fMainForm.Toolbar3.Buttons("DropDownBox").Enabled = True
    fMainForm.Toolbar3.Buttons("PushButton").Enabled = True
    fMainForm.Toolbar3.Buttons("HiddenData").Enabled = True
    fMainForm.Toolbar3.Buttons("Password").Enabled = True
    fMainForm.Toolbar3.Buttons("SubmitButton").Enabled = True
    fMainForm.Toolbar3.Buttons("ResetButton").Enabled = True
    fMainForm.Toolbar3.Buttons("ImageButton").Enabled = True
    'fMainForm.Toolbar3.Buttons("FileUpload").Enabled = True
    '----------------------------------------------------------------
    If DhtmlShowAll Then
           fMainForm.Toolbar2.Buttons("ShowAll").Value = tbrPressed
    Else
           fMainForm.Toolbar2.Buttons("ShowAll").Value = tbrUnpressed
    End If
    
    If DhtmlShowBorders Then
           fMainForm.Toolbar2.Buttons("ShowBorders").Value = tbrPressed
    Else
           fMainForm.Toolbar2.Buttons("ShowBorders").Value = tbrUnpressed
    End If
    
    If DhtmlSnapToGrid Then
           fMainForm.Toolbar3.Buttons("SnapToGrid").Value = tbrPressed
    Else
           fMainForm.Toolbar3.Buttons("SnapToGrid").Value = tbrUnpressed
    End If
    '----------------------------------------------------------------
    Do
        DoEvents
    Loop While DHTMLEditInitialized = False
    '----------------------------------------------------------------
    'Load document DOM to tree view
    'fMainForm.UpdateDocTree
    'fMainForm.cmdTreeRefresh_Click
    '----------------------------------------------------------------
    'DHTMLEdit1.DOM.ExecCommand "LiveResize", False, True
    'DHTMLEdit1.DOM.ExecCommand "MultipleSelection", False, True
    '----------------------------------------------------------------
Exit Sub
    
ErrHandler:
    MsgBox ("Error loading the dhtml editor. Error number: " & Abs(CLng(Err.Number)) & ". " & Err.Description)
    
End Sub

Private Sub Form_Resize()

    If fMainForm.WindowState = vbMinimized Then Exit Sub
    
    DHTMLEdit1.Width = Abs(Me.ScaleWidth - 0)
    DHTMLEdit1.Height = Abs(Me.ScaleHeight - 0)
    DHTMLEdit1.Move 0, 0
     
    Editawy1.Width = Abs(Me.ScaleWidth - 0)
    Editawy1.Height = Abs(Me.ScaleHeight - 0)
    Editawy1.Move 0, 0
    
    SaveSetting App.Title, "Settings", "WindowState", Me.WindowState
    
End Sub

Private Sub DHTMLEdit1_DocumentComplete()
    
    If Not DHTMLEditInitialized Then
        
        Dim fmt As DEGetBlockFmtNamesParam
        Dim i As Long
        Dim fontSize As Long
        Dim fmtName As Variant
        
        ' Create the block fmt names holder
        Set fmt = CreateObject("DEGetBlockFmtNamesParam.DEGetBlockFmtNamesParam.1")
        
        ' Get the localized strings for the DECMD_SETBLOCKFMT command
        DHTMLEdit1.execCommand DECMD_GETBLOCKFMTNAMES, OLECMDEXECOPT_DONTPROMPTUSER, fmt
        
        ' Put the strings into the Format menu
        i = 0
        fMainForm.StyleCombo.Clear
        For Each fmtName In fmt.Names
            'FormatSub(i).Caption = fmtName
            fMainForm.StyleCombo.AddItem fmtName
            i = i + 1
        Next
        
        'fMainForm.UpdateFontCombos
        
        fMainForm.FontSizeCombo.ListIndex = fontSize - 1
        
        fMainForm.cmdTreeRefresh_Click
    End If
    
    DHTMLEditInitialized = True
    
End Sub

Private Function GetNewDocHTML() As String
    
    Dim strText As String

    strText = "" & _
"<html>" & vbCrLf & _
"<head>" & vbCrLf & _
"<title> new document </title>" & vbCrLf & _
"<meta name=""generator"" content=""Mewsoft Webawy " & CStr(MyVersion) & """>" & vbCrLf & _
"<meta name=""ProgID"" content=""www.Mewsoft.com"">" & vbCrLf & _
"<meta name=""author"" content="""">" & vbCrLf & _
"<meta name=""keywords"" content="""">" & vbCrLf & _
"<meta name=""description"" content="""">" & vbCrLf & _
"<meta http-equiv=""Content-Type"" content=""text/html;"">" & vbCrLf & _
"</head>" & vbCrLf & _
"" & vbCrLf & _
"<body>" & vbCrLf & _
"" & vbCrLf & _
"</body>" & vbCrLf & _
"</html>"
    
    GetNewDocHTML = strText
    
End Function

Private Sub DHTMLEdit1_DisplayChanged()
    
    On Error Resume Next
    
    Dim State As DHTMLEDITCMDF
    Dim Cmd As DHTMLEDITCMDID
    Dim Button As String
    Dim cmds As Long
    
    ' DHTMLEdit indicates the UI should be updated
    ' First update the Toolbar
    '------------------------------------------------------
   ' Bold button
    fMainForm.UpdateToolbarButton1 "Bold", DECMD_BOLD
    
   ' Italic button
    fMainForm.UpdateToolbarButton1 "Italic", DECMD_ITALIC
    
   ' Underline button
    fMainForm.UpdateToolbarButton1 "Underline", DECMD_UNDERLINE
    
   ' Numbers button
    fMainForm.UpdateToolbarButton1 "Numbers", DECMD_ORDERLIST
    
   ' Bullets button
    fMainForm.UpdateToolbarButton1 "Bullets", DECMD_UNORDERLIST
    
   ' Outdent button
    fMainForm.UpdateToolbarButton1 "Outdent", DECMD_OUTDENT
    
   ' Indent button
    fMainForm.UpdateToolbarButton1 "Indent", DECMD_INDENT
    
   ' Left Justify button
    fMainForm.UpdateToolbarButton1 "LeftJustify", DECMD_JUSTIFYLEFT
    
   ' Center Justify button
    fMainForm.UpdateToolbarButton1 "Center", DECMD_JUSTIFYCENTER
   
   ' Right Justify button
    fMainForm.UpdateToolbarButton1 "RightJustify", DECMD_JUSTIFYRIGHT
    
    'JustifyFull
    fMainForm.UpdateToolbarButton11 "JustifyFull", "JustifyFull"
    
    'SuperScript
    fMainForm.UpdateToolbarButton11 "SuperScript", "SuperScript"
        
    'SubScript
    fMainForm.UpdateToolbarButton11 "SubScript", "SubScript"
                    
    'StrikeThrough
    fMainForm.UpdateToolbarButton11 "StrikeThrough", "StrikeThrough"
    
   ' Color button
    fMainForm.UpdateToolbarButton1 "Color", DECMD_SETFORECOLOR
    
   ' BackColor button
    fMainForm.UpdateToolbarButton1 "BackColor", DECMD_SETBACKCOLOR
    '------------------------------------------------------
'    If Len(DHTMLEdit1.CurrentDocumentPath) > 0 Then
'        'FileSave.Enabled = True
'        fMainForm.Toolbar2.Buttons("Save").Enabled = True
'    Else
'        'FileSave.Enabled = False
'        fMainForm.Toolbar2.Buttons("Save").Enabled = False
'    End If
    '------------------------------------------------------
    ' Update the fonts menu
    fMainForm.UpdateFontCombos
    '------------------------------------------------------
    ' Update the Format menu with the localized strings returned from
    ' the DECMD_GETBLOCKFMT command
    State = DHTMLEdit1.QueryStatus(DECMD_GETBLOCKFMT)
    If State >= DECMDF_ENABLED Then
        Dim blockFmt As String
        blockFmt = DHTMLEdit1.execCommand(DECMD_GETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER)
        fMainForm.StyleCombo.Text = blockFmt
    End If
    '------------------------------------------------------
   ' Undo button
    fMainForm.UpdateToolbarButton2 "Undo", DECMD_UNDO
    ' Redo button
    fMainForm.UpdateToolbarButton2 "Redo", DECMD_REDO
    ' Copy button
    fMainForm.UpdateToolbarButton2 "Copy", DECMD_COPY
    ' Cut button
    fMainForm.UpdateToolbarButton2 "Cut", DECMD_CUT
    ' Paste button
    fMainForm.UpdateToolbarButton2 "Paste", DECMD_PASTE
    ' Hyperlink button
    fMainForm.UpdateToolbarButton2 "Hyperlink", DECMD_HYPERLINK
    ' Image button
    fMainForm.UpdateToolbarButton2 "Image", DECMD_IMAGE
    ' Table button
    fMainForm.UpdateToolbarButton2 "Table", DECMD_INSERTTABLE
    ' Find text
    fMainForm.UpdateToolbarButton2 "Find", DECMD_FINDTEXT
    ' Properties
    fMainForm.UpdateToolbarButton2 "Properties", DECMD_PROPERTIES
    '------------------------------------------------------
    'Show Borders button
     State = DHTMLEdit1.ShowBorders
     If State Then
            fMainForm.Toolbar2.Buttons("ShowBorders").Value = tbrPressed
     Else
            fMainForm.Toolbar2.Buttons("ShowBorders").Value = tbrUnpressed
     End If
    
     'Show All button
     State = DHTMLEdit1.ShowDetails
     If State Then
            fMainForm.Toolbar2.Buttons("ShowAll").Value = tbrPressed
     Else
            fMainForm.Toolbar2.Buttons("ShowAll").Value = tbrUnpressed
     End If
    '------------------------------------------------------
    'Insert Rows
    fMainForm.UpdateToolbarButton3 "InsertRows", DECMD_INSERTROW
    
    'Insert Columns
    fMainForm.UpdateToolbarButton3 "InsertColumns", DECMD_INSERTCOL
    
    'Insert Cells
    fMainForm.UpdateToolbarButton3 "InsertCells", DECMD_INSERTCELL
    
    'Delete Cells
    fMainForm.UpdateToolbarButton3 "DeleteCells", DECMD_DELETECELLS
    
    'Delete Columns
    fMainForm.UpdateToolbarButton3 "DeleteColumns", DECMD_DELETECOLS
    
    'Delete Rows
    fMainForm.UpdateToolbarButton3 "DeleteRows", DECMD_DELETEROWS
    
    'Merge Cells
    fMainForm.UpdateToolbarButton3 "MergeCells", DECMD_MERGECELLS
    
    'Split Cells
    fMainForm.UpdateToolbarButton3 "SplitCells", DECMD_SPLITCELL
    
    '------------------------------------------------------
    ' Initialize 2D menu command table
    'PositionAbsolutely
    fMainForm.UpdateToolbarButton3 "PositionAbsolutely", DECMD_MAKE_ABSOLUTE

    'BringForward
    fMainForm.UpdateToolbarButton3 "BringForward", DECMD_BRING_FORWARD

    'SendBackward
    fMainForm.UpdateToolbarButton3 "SendBackward", DECMD_SEND_BACKWARD

    'BringToFront
    fMainForm.UpdateToolbarButton3 "BringToFront", DECMD_BRING_TO_FRONT

    'SendToBack
    fMainForm.UpdateToolbarButton3 "SendToBack", DECMD_SEND_TO_BACK
    
    'BringAboveText
    fMainForm.UpdateToolbarButton3 "BringAboveText", DECMD_BRING_ABOVE_TEXT
    
    'SendBelowText
    fMainForm.UpdateToolbarButton3 "SendBelowText", DECMD_SEND_BELOW_TEXT
    '------------------------------------------------------
    Exit Sub
    
ErrHandler:
    
End Sub

Private Sub DHTMLEdit1_ContextMenuAction(ByVal itemIndex As Long)

    ' Handle user selection on the custom context menu
    
    On Error Resume Next
    
   Select Case itemIndex
    Case 0
        DHTMLEdit1.execCommand DECMD_CUT, OLECMDEXECOPT_DODEFAULT
    Case 1
        DHTMLEdit1.execCommand DECMD_COPY, OLECMDEXECOPT_DODEFAULT
    Case 2
        DHTMLEdit1.execCommand DECMD_PASTE, OLECMDEXECOPT_DODEFAULT
    Case 4
        DHTMLEdit1.execCommand DECMD_SELECTALL, OLECMDEXECOPT_DODEFAULT
    Case 6
        DHTMLEdit1.execCommand DECMD_FONT, OLECMDEXECOPT_PROMPTUSER
    End Select
    
    If ctxtIs2DCapable Then
        Select Case itemIndex
        Case ctxtStdItemCount + 2
            DHTMLEdit1.execCommand DECMD_MAKE_ABSOLUTE, OLECMDEXECOPT_DODEFAULT
        End Select
    End If
    
    If ctxtIsTable Then
        Select Case itemIndex
        Case ctxtStdItemCount + ctxt2DItemCount + 2
            DHTMLEdit1.execCommand DECMD_INSERTROW, OLECMDEXECOPT_DODEFAULT
        Case ctxtStdItemCount + ctxt2DItemCount + 3
            DHTMLEdit1.execCommand DECMD_INSERTCOL, OLECMDEXECOPT_DODEFAULT
        Case ctxtStdItemCount + ctxt2DItemCount + 5
            DHTMLEdit1.execCommand DECMD_DELETEROWS, OLECMDEXECOPT_DODEFAULT
        Case ctxtStdItemCount + ctxt2DItemCount + 6
            DHTMLEdit1.execCommand DECMD_DELETECOLS, OLECMDEXECOPT_DODEFAULT
        End Select
    End If
    
End Sub

Private Sub DHTMLEdit1_ShowContextMenu(ByVal x As Long, ByVal y As Long)
    
    Dim cmdState As DHTMLEDITCMDF
    Dim strings() As String
    Dim states() As OLE_TRISTATE
    
   ' Create dynamic context menu that consists of
   ' a "standard" set of items and items that depend
   ' on the currently selected element.
   ' Look at the current selection and
   ' if its a table then add menu items for add/delete rows and cols
   ' if its 2DCapable then add items to toggle its absolute position attribute
   
    ctxtIs2DCapable = False
    ctxtIsAbsPos = False
    ctxtIsTable = False
        
    ' Determine if the selected element is 2D capable
    cmdState = DHTMLEdit1.QueryStatus(DECMD_MAKE_ABSOLUTE)
    If cmdState >= DECMDF_ENABLED Then
        ctxtIs2DCapable = True
    End If
    
    'Use DECMD_SEND_TO_BACK to determine if this element is abs positioned
    cmdState = DHTMLEdit1.QueryStatus(DECMD_SEND_TO_BACK)
    If cmdState >= DECMDF_ENABLED Then
        ctxtIsAbsPos = True
    End If
    
    'Use DECMD_INSERTROW to determine if this element is a table
    cmdState = DHTMLEdit1.QueryStatus(DECMD_INSERTROW)
    If cmdState >= DECMDF_ENABLED Then
        ctxtIsTable = True
    End If
    
    ctxtStdItemCount = 6
    
    If ctxtIs2DCapable Then
        ctxt2DItemCount = 2 '1 Item + 1 Separator
    Else
        ctxt2DItemCount = 0
    End If
    
    
    If ctxtIsTable Then
        ctxtTableItemCount = 6 '4 Items + 2 Separators
    Else
        ctxtTableItemCount = 0
    End If
    
    
    ReDim strings(0 To ctxtStdItemCount + ctxt2DItemCount + ctxtTableItemCount)
    ReDim states(0 To ctxtStdItemCount + ctxt2DItemCount + ctxtTableItemCount)
    
    strings(0) = "Cut"
    strings(1) = "Copy"
    strings(2) = "Paste"
    strings(3) = ""
    strings(4) = "Select All"
    strings(5) = ""
    strings(6) = "Font..."
        
    cmdState = DHTMLEdit1.QueryStatus(DECMD_CUT)
    If cmdState >= DECMDF_ENABLED Then
         states(0) = Unchecked
     Else
         states(0) = Gray
    End If
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_COPY)
    If cmdState >= DECMDF_ENABLED Then
         states(1) = Unchecked
     Else
         states(1) = Gray
    End If
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_PASTE)
    If cmdState >= DECMDF_ENABLED Then
         states(2) = Unchecked
     Else
         states(2) = Gray
    End If
        
    states(3) = Unchecked
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_SELECTALL)
    If cmdState >= DECMDF_ENABLED Then
         states(4) = Unchecked
     Else
         states(4) = Gray
    End If
    
    states(5) = Unchecked
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_FONT)
    If cmdState >= DECMDF_ENABLED Then
         states(6) = Unchecked
     Else
         states(6) = Gray
    End If
    
    If ctxtIs2DCapable Then
        strings(ctxtStdItemCount + 1) = ""
        states(ctxtStdItemCount + 1) = Unchecked
        If ctxtIsAbsPos Then
            strings(ctxtStdItemCount + 2) = "Make 1D"
        Else
            strings(ctxtStdItemCount + 2) = "Make 2D"
        End If
        states(ctxtStdItemCount + 2) = Unchecked
    End If
    
    If ctxtIsTable Then
        strings(ctxtStdItemCount + ctxt2DItemCount + 1) = ""
        states(ctxtStdItemCount + ctxt2DItemCount + 1) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 2) = "Insert Row"
        states(ctxtStdItemCount + ctxt2DItemCount + 2) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 3) = "Insert Column"
        states(ctxtStdItemCount + ctxt2DItemCount + 3) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 4) = ""
        states(ctxtStdItemCount + ctxt2DItemCount + 4) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 5) = "Delete Row"
        states(ctxtStdItemCount + ctxt2DItemCount + 5) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 6) = "Delete Column"
        states(ctxtStdItemCount + ctxt2DItemCount + 6) = Unchecked
        
    End If
    
    DHTMLEdit1.setContextMenu strings, states
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'fMainForm.DisableToolbar
    
    OpenDocuments.Remove Me.Tag
    EditMode.Remove Me.Tag
    OpenFilenames.Remove Me.Tag
    fMainForm.tabFilesBar.Tabs.Remove ("F" & Me.Tag)
    
    If OpenDocuments.Count < 1 Then
        fMainForm.DisableToolbar
    End If
    
    fMainForm.tvDomTree.Nodes.Clear
    fMainForm.lvProperties.ListItems.Clear
    fMainForm.lbProperty.Caption = ""
    fMainForm.txtProperty.Text = ""
    fMainForm.txtProperty.Tag = ""
    
    'fMainForm.tabFilesBar.Tabs.Remove fMainForm.tabFilesBar.SelectedItem.Key
    
End Sub

Private Sub Editawy1_UpdateUI(ByVal Line As Long, ByVal Column As Long, ByVal Position As Long, ByVal TotalLines As Long)

    'fMainForm.sbStatusBar.Panels(2).Text = "Line " & Editawy1.CurLine & " "
    'fMainForm.sbStatusBar.Panels(3).Text = "Column " & Editawy1.Column + 1 & " "
    fMainForm.sbStatusBar.Panels(1).Text = ""
    'fMainForm.sbStatusBar.Panels(4).Text = " Char " & Editawy1.CurPosition & " "

   'ActiveForm.Editawy1.CurPosition
    'ActiveForm.Editawy1.Column
    'ActiveForm.Editawy1.TotalLines

End Sub

Private Sub DHTMLEdit1_onmousemove()

    On Error GoTo ErrHandler
    
    Dim e As IHTMLEventObj
    Set e = DHTMLEdit1.DOM.parentWindow.event
    
    fMainForm.sbStatusBar.Panels(2).Text = "X " & e.clientX & " "
    fMainForm.sbStatusBar.Panels(3).Text = "Y " & e.clientY & " "
ErrHandler:
End Sub

Private Sub DHTMLEdit1_onclick()
    
    On Error GoTo ErrHandler
    
    Dim e As IHTMLEventObj
    Set e = DHTMLEdit1.DOM.parentWindow.event
    
    Const IDM_SETDIRTY = 2342
    Const IDM_BOLD = 52
    
    'http://msdn.microsoft.com/library/default.asp?url=/workshop/author/dhtml/reference/objects/fieldset.asp
    'Debug.Print DHTMLEdit1.IsDirty
    
    'execCommand, queryCommandEnabled, queryCommandIndeterm,
    'queryCommandState, queryCommandSupported, queryCommandValue.
    
    'DHTMLEdit1.DOM.ExecCommand ("Superscript")
    'DHTMLEdit1.DOM.ExecCommand ("DirRTL") 'DirLTR
    'DHTMLEdit1.DOM.ExecCommand ("InsertHorizontalRule")
    
    'Overwrites an inline frame on the text selection.
    'DHTMLEdit1.DOM.ExecCommand ("InsertIFrame")
    
    'Overwrites a line break on the text selection.
    'DHTMLEdit1.DOM.ExecCommand ("InsertParagraph") ' <p></p>
    
    'Overwrites a drop-down selection control on the text selection.
    'DHTMLEdit1.DOM.ExecCommand ("InsertSelectDropdown")
    
    'Overwrites a list box selection control on the text selection.
    'DHTMLEdit1.DOM.ExecCommand ("InsertSelectListbox")
    
    'DHTMLEdit1.DOM.ExecCommand ("JustifyFull") ' JustifyNone
    
    'Causes the MSHTML Editor to update an element's appearance
    'continuously during a resizing or moving operation, rather
    'than updating only at the completion of the move or resize.
    'DHTMLEdit1.DOM.ExecCommand ("LiveResize")
    
    'Allows for the selection of more than one site selectable
    'element at a time when the user holds down the SHIFT or CTRL keys.
    'DHTMLEdit1.DOM.ExecCommand ("MultipleSelection")
    
    'Removes the formatting tags from the current selection.
    'DHTMLEdit1.DOM.ExecCommand ("RemoveFormat")
    
    'DHTMLEdit1.DOM.ExecCommand ("RemoveFormat") 'RemoveParaFormat
    
    'DHTMLEdit1.DOM.ExecCommand ("Strikethrough")
    
    'DHTMLEdit1.DOM.ExecCommand ("Subscript")
    'DHTMLEdit1.DOM.ExecCommand ("Superscript")
    
    'DHTMLEdit1.DOM.ExecCommand ("BlockDirLTR")
    'DHTMLEdit1.DOM.ExecCommand ("BlockDirRTL")
    
    
    'Sets or retrieves the background color of the current selection.
    'DHTMLEdit1.DOM.ExecCommand "BackColor", True, "#00ff00"
    
    'DHTMLEdit1.DOM.ExecCommand "ForeColor", "false", "#FF0033"
    
    'DHTMLEdit1.DOM.ExecCommand ("UnBookmark")
    'Unlink
    
    
    'DHTMLEdit1.DOM.ExecCommand "CreateBookmark", False, "MyBookmark1"
    'Debug.Print DHTMLEdit1.DOM.queryCommandValue("CreateBookmark")
    
    
    'DHTMLEdit1.DOM.ExecCommand "FormatBlock"
    'Debug.Print DHTMLEdit1.DOM.queryCommandValue("FormatBlock")
    
    'Deletes the current selection.
    'DHTMLEdit1.DOM.ExecCommand ("Delete")
    
    'Debug.Print DHTMLEdit1.DOM.ExecCommand("CreateLink", False, "mewsoft.com")
    
    'Sets or retrieves the font for the current selection.
    'Debug.Print DHTMLEdit1.DOM.ExecCommand("FontName", False, "Times")
    

    'Sets or retrieves the font size for the current selection
    'Debug.Print DHTMLEdit1.DOM.queryCommandValue("fontSize")
    'DHTMLEdit1.DOM.ExecCommand "fontSize", False, 4
    
    'DHTMLEdit1.DOM.ExecCommand "InsertImage", False, "http://msdn.microsoft.com/library/toolbar/3.0/images/banners/msdn_masthead_ltr.gif"
    
    'DHTMLEdit1.DOM.ExecCommand "InsertInputFileUpload", False, "MyID"
    
    'DHTMLEdit1.DOM.ExecCommand "InsertInputHidden", False, "MyID"
    
    'DHTMLEdit1.DOM.ExecCommand ("Undo")
    
    'DHTMLEdit1.DOM.ExecCommand ("Unlink")
    
    'Overwrites a box on the text selection.
    'Draws a box around the text and other elements that the field
    'set contains. This element is useful for grouping elements in a
    'form and for distinctively marking text in a document.
    'The FIELDSET element has the same behavior as a window frame.

    'DHTMLEdit1.DOM.ExecCommand ("InsertFieldset")
    
    'InsertInputButton
    'Overwrites a button control on the text selection.
    '
    'InsertInputCheckbox
    'Overwrites a check box control on the text selection.
    '
    'InsertInputFileUpload
    'Overwrites a file upload control on the text selection.
    '
    'InsertInputHidden
    'Inserts a hidden control on the text selection.
    '
    'InsertInputImage
    'Overwrites an image control on the text selection.
    '
    'InsertInputPassword
    'Overwrites a password control on the text selection.
    '
    'InsertInputRadio
    'Overwrites a radio control on the text selection.
    '
    'InsertInputReset
    'Overwrites a reset control on the text selection.
    '
    'InsertInputSubmit
    'Overwrites a submit control on the text selection.
    '
    'InsertInputText
    'Overwrites a text control on the text selection.
    
    'Debug.Print DHTMLEdit1.IsDirty
    '----------------------------------------------------------------
    'sbStatusBar.Panels(1) = e.fromElement.tagName
    'Screen.TwipsPerPixelX
    
    fMainForm.sbStatusBar.Panels(2).Text = "X " & e.x & " "
    fMainForm.sbStatusBar.Panels(3).Text = "Y " & e.y & " "
    
    If Not e.fromElement Is Nothing Then
        'Debug.Print e.fromElement.tagName
        'fMainForm.sbStatusBar.Panels(1).Text = e.fromElement.tagName
    End If
        
    fMainForm.sbStatusBar.Panels(1).Text = GetElementUnderInsertionPointValue
    
    fMainForm.sbStatusBar.Panels(4).Text = "Char " & CStr(fMainForm.GetDocCartLocation)
    
    'Set el = DHTMLEdit1.DOM.elementFromPoint(e.x, e.y)
    'fMainForm.sbStatusBar.Panels(1).Text = el.tagName
    
    'ShowSelectedElement
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    'DHTMLEdit1.DOM.ExecCommand "GOTO", False, 1
    'DHTMLEdit1.DOM.ExecCommand "{DE4BA900-59CA-11CF-9592-444553540000}", 19
    
    Dim doc1 As MSHTML.HTMLDocument
    Set doc1 = DHTMLEdit1.DOM
    
    'doc1.ExecCommand
    
   'Debug.Print doc.queryCommandEnabled(IID_IDisplayServices)
   'Debug.Print frmBrowser.brwWebBrowser.Document.QueryInterface(IID_IDisplayServices, isp)
'   Debug.Print doc1.QueryInterface(IID_IDisplayServices, isp)
    'Set doc1 = DHTMLEdit1.DOM
    'Debug.Print doc1.ExecCommand("bold")
    'Call SendMessage(DHTMLEdit1.AbsoluteDropMode , EM_EMPTYUNDOBUFFER, 0, 0)
    
    '----------------------------------------------------------------
    Exit Sub

ErrHandler:

End Sub

Private Function GetElementUnderInsertionPointValue() As String

    On Error GoTo ErrHandler

    Dim rg As IHTMLTxtRange
    Dim ctlRg As IHTMLControlRange
    Dim Text As String
    
    Select Case DHTMLEdit1.DOM.selection.Type
       Case "None", "Text"
          ' This reduces the selection to just the insertion
          ' point. The parentElement method will then return the
          ' element directly under the mouse pointer.
          Set rg = DHTMLEdit1.DOM.selection.createRange
          rg.collapse
          Text = rg.parentElement.outerHTML
          'Text = rg.parentElement.innerHTML
          'Text = rg.parentElement.tagName
          Text = CleanCrLf(Text)
       
       Case "Control"
          ' A form or image is selected. The commonParentElement
          ' will return the site selected element.
          Set ctlRg = DHTMLEdit1.DOM.selection.createRange
          Text = ctlRg.commonParentElement.outerHTML
          'Text = Text & " - " & ctlRg.commonParentElement.tagName
          'Text = ctlRg.commonParentElement.innerHTML
          Text = CleanCrLf(Text)
    End Select
    
    GetElementUnderInsertionPointValue = Text
       
    Exit Function

ErrHandler:
   GetElementUnderInsertionPointValue = ""
End Function


Private Sub DHTMLEdit1_onmouseup()

    On Error Resume Next
    ShowSelectedElement

End Sub

Private Sub ViewProperties()
    
    On Error Resume Next
    
    Dim el As IHTMLElement
    Dim de As IHTMLDivElement

    Set el = GetElementUnderCaret()
    
    Debug.Print el.accessKey
    Debug.Print el.canHaveChildren
    Debug.Print el.clientHeight
    Debug.Print el.clientLeft
    Debug.Print el.clientTop
    Debug.Print el.clientWidth
    Debug.Print el.Dir
    Debug.Print el.scopeName
    Debug.Print el.TabIndex
    Debug.Print el.TagName
    
End Sub

Public Sub UpdateProperties(el As Object)
    
    On Error Resume Next
    
    Dim i As Integer
    Dim LvItem As ListItem
    Dim Counter As Long
    Dim Value As String
    
    On Error Resume Next
    
    fMainForm.lvProperties.ListItems.Clear
    
    Counter = 0
    For i = 0 To el.Attributes.length - 1
        If left$(el.Attributes(i).nodeName, 2) <> "on" Then
        
            'Debug.Print el.Attributes(i).nodeName & "=" & el.Attributes(i).nodeValue
            
            Value = el.Attributes(i).nodeValue
            If Value = "" Then Value = " "
            
            Set LvItem = fMainForm.lvProperties.ListItems.Add
            Counter = Counter + 1
            With LvItem
                .Text = el.Attributes(i).nodeName
                .Key = el.Attributes(i).nodeName
                .ListSubItems.Add , , Value
                .ListSubItems.Add , , CStr(Counter)
            End With
        End If
    Next
    
ehUpdateProperties:
    Exit Sub
    
End Sub

' Update the UI to display the tag of the selected element
Private Sub ShowSelectedElement()

    On Error Resume Next
    
    Dim el As IHTMLElement
    Dim divEl As HTMLDivElement
    Dim lKey As Long

    Set el = GetElementUnderCaret
    
    ' keep track of currently selected element for updates
    Set ActiveIHTMLElement = el

    If el Is Nothing Then
        Exit Sub
    End If

    '----------------------------------------------------------------
    ' Disable the command button for synching DIV events if
    ' this element is not a DIV
    If el.TagName = "DIV" Then

        On Error Resume Next
        Set divEl = el
        If Not divEl Is Nothing Then
'            SyncDivCommand.Enabled = True
            UpdateProperties divEl
        Else
'            SyncDivCommand.Enabled = False
        End If
    Else
'        SyncDivCommand.Enabled = False
        UpdateProperties el
    End If
    '----------------------------------------------------------------
    lKey = el.sourceIndex + 1 ' sourceIndex starts from zero, treeview from 1
    'debug.Print "lKey: " & lKey
    'MainForm.tvDomTree.SelectedItem.Index = lKey
    fMainForm.tvDomTree.Nodes.Item(lKey).Selected = True
    '----------------------------------------------------------------
    
End Sub

'Returns the element directly under the insertion point
Private Function GetElementUnderCaret() As IHTMLElement
    
    On Error Resume Next
    
    Dim rg As IHTMLTxtRange
    Dim ctlRg As IHTMLControlRange

    ' Branch on the type of selection and
    ' get the element under the caret or the site selected object
    ' and return it

    Select Case DHTMLEdit1.DOM.selection.Type

    Case "None", "Text"

        Set rg = DHTMLEdit1.DOM.selection.createRange

        ' Collapse the range so that the scope of the
        ' range of the selection is the the caret. That way the
        ' parentElement method will return the element directly
        ' under the caret. If you don't want to change the state of the
        ' selection, then duplicate the range and collapse it

        If Not rg Is Nothing Then
            rg.collapse
            Set GetElementUnderCaret = rg.parentElement
        End If

    Case "Control"
        ' An element is site selected
        Set ctlRg = DHTMLEdit1.DOM.selection.createRange
 
        ' There can only be one site selected element at a time so the
        ' commonParentElement will return the site selected element
        Set GetElementUnderCaret = ctlRg.commonParentElement
        
    End Select

End Function

Private Sub MoveToEnd()
    Dim range As IHTMLTxtRange
    Set range = DHTMLEdit1.DOM.body.createTextRange()
    range.collapse False
    range.Select
End Sub

Public Sub CheckSpelling()
    
    On Error Resume Next
    
'    Dim objWord As New Word.Application
'    Dim objDocument As New Document
'    objWord.WindowState = wdWindowStateMinimize
'    Set objDocument = objWord.Documents.Add()
'    With objDocument
'        .Select
'        .Range.Text = DHTMLEdit1.DocumentHTML
'        .CheckSpelling , True, True
'        .CheckGrammar
'        DHTMLEdit1.DocumentHTML = Replace(.Range.Text, vbCr & vbCr, vbCrLf & vbCrLf)
'        .Saved = True
'        .Close
'    End With
'    objWord.Quit
'    Set objDocument = Nothing
'    Set objWord = Nothing
End Sub

Private Sub DHTMLEdit1_DragDrop(source As Control, x As Single, y As Single)
'(source as tbutton).left := x;
'(source as tbutton).top := y;
Debug.Print "DHTMLEdit1_DragDrop"
End Sub

Private Sub DHTMLEdit1_DragOver(source As Control, x As Single, y As Single, State As Integer)
'DHTMLEdit.OleObject.DOM.body.bgColor := '800080';
'http://msdn.microsoft.com/library/default.asp?url=/workshop/author/dhtml/reference/objects/obj_document.asp

Debug.Print "DHTMLEdit1_DragOver"

End Sub

