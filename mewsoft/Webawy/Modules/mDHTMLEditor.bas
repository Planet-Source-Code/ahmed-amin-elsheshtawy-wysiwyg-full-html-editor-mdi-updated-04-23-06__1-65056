Attribute VB_Name = "mDHTMLEditor"
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

Public DHTMLEditInitialized As Boolean
Public buttonCmds(1 To 12) As DHTMLEDITCMDID
Public buttonNames(1 To 12) As String

' Tables for menus commands
Public editMenuCmds(0 To 8) As DHTMLEDITCMDID
Public insertMenuCmds(0 To 1) As DHTMLEDITCMDID
Public tableMenuCmds(0 To 11) As DHTMLEDITCMDID
Public twoDMenuCmds(0 To 8) As DHTMLEDITCMDID
' Document path name
Public docPath
' State variables for dynamic context menu
Public ctxtIs2DCapable As Boolean
Public ctxtIsAbsPos As Boolean
Public ctxtIsTable As Boolean
Public ctxtStdItemCount As Long
Public ctxt2DItemCount As Long
Public ctxtTableItemCount As Long

Private Enum General
    DE_E_INVALIDARG = &H5
    DE_E_ACCESS_DENIED = &H46
    DE_E_PATH_NOT_FOUND = &H80070003
    DE_E_FILE_NOT_FOUND = &H80070002
    DE_E_UNEXPECTED = &H8000FFFF
    DE_E_DISK_FULL = &H80070027
    DE_E_NOTSUPPORTED = &H80040100
    DE_E_FILTER_FRAMESET = &H80100001
    DE_E_FILTER_SERVERSCRIPT = &H80100002
    DE_E_FILTER_MULTIPLETAGS = &H80100004
    DE_E_FILTER_SCRIPTLISTING = &H80100008
    DE_E_FILTER_SCRIPTLABEL = &H80100010
    DE_E_FILTER_SCRIPTTEXTAREA = &H80100020
    DE_E_FILTER_SCRIPTSELECT = &H80100040
    DE_E_URL_SYNTAX = &H800401E4
    DE_E_INVALID_URL = &H800C0002
    DE_E_NO_SESSION = &H800C0003
    DE_E_CANNOT_CONNECT = &H800C0004
    DE_E_RESOURCE_NOT_FOUND = &H800C0005
    DE_E_OBJECT_NOT_FOUND = &H800C0006
    DE_E_DATA_NOT_AVAILABLE = &H800C0007
    DE_E_DOWNLOAD_FAILURE = &H800C0008
    DE_E_AUTHENTICATION_REQUIRED = &H800C0009
    DE_E_NO_VALID_MEDIA = &H800C000A
    DE_E_CONNECTION_TIMEOUT = &H800C000B
    DE_E_INVALID_REQUEST = &H800C000C
    DE_E_UNKNOWN_PROTOCOL = &H800C000D
    DE_E_SECURITY_PROBLEM = &H800C000E
    DE_E_CANNOT_LOAD_DATA = &H800C000F
    DE_E_CANNOT_INSTANTIATE_OBJECT = &H800C0010
    DE_E_REDIRECT_FAILED = &H800C0014
    DE_E_REDIRECT_TO_DIR = &H800C0015
    DE_E_CANNOT_LOCK_REQUEST = &H8
End Enum


Public Const NormalMode As String = 1
Public Const HtmlMode As String = 2
Public Const PreviewMode As String = 3

Public FormIndex As Long
Public TextboxIndex As Long
Public TextareaIndex As Long
Public CheckboxIndex As Long
Public OptionButtonIndex As Long
Public ListBoxIndex As Long
Public DropDownBoxIndex As Long
Public PushButtonIndex As Long
Public HiddenDataIndex As Long
Public PasswordIndex As Long
Public SubmitButtonIndex As Long
Public ResetButtonIndex As Long
Public ImageButtonIndex As Long
Public FileUploadIndex As Long
'----------------------------------------------------------
Public fontNames() ' list of fonts is the fontComboBox
Public OpenDocuments As New Collection ' ' Open documents forms
Public OpenFilenames As New Collection ' Open documents files names
Public EditMode As New Collection ' Open documents current edit mode
Public ActiveDocument As Integer        ' Holds various info to differentiate documents
Public lDocumentCount As Long           ' Count of documents
Public NewDocumentCount As Long           ' Count of new documents
Public ActiveIHTMLElement As IHTMLElement ' Selected document element
Public TreeNodeID As Double           'Static counter for tree nodes unique ID
Public DhtmlSnapToGrid As Boolean          'Toogle Snap To Grid
Public DhtmlSnapToGridX As Long
Public DhtmlSnapToGridY As Long

Public DhtmlShowBorders As Boolean
Public DhtmlShowAll As Boolean
'----------------------------------------------------------
'----------------------------------------------------------

