VERSION 5.00
Begin VB.Form frmOpenURL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open URL"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frmOpenURL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5820
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpenURL 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Top             =   1020
      Width           =   6615
   End
   Begin VB.Label Label2 
      Caption         =   "URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Type the internet address of the document to open."
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4395
   End
End
Attribute VB_Name = "frmOpenURL"
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
Option Explicit

Private Sub cboAddress_Click()
    
     cmdOpenURL_Click
     
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)

    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cmdOpenURL_Click
    End If

End Sub

Private Sub CmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdOpenURL_Click()
    
    On Error GoTo ErrHandler
    Dim URL As String
    
    URL = QualifyURL(cboAddress.Text)
    
    Screen.MousePointer = vbHourglass
    If URL <> "" Then
        fMainForm.LoadDownloadedDoc
        OpenDocuments(CStr(ActiveDocument)).DHTMLEdit1.LoadURL URL
    End If
    Screen.MousePointer = vbDefault
    
    Unload Me
    fMainForm.tabFilesBar.Tabs.Item("F" & CStr(lDocumentCount)).Selected = True
    Exit Sub
    
ErrHandler:
    MsgBox "Error opening the web page you specified. Please check the address again."
    cboAddress.SetFocus
    
    'Unload Me
End Sub
