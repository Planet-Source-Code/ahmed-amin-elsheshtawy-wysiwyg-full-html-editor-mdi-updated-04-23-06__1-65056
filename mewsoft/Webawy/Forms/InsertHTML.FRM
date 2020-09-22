VERSION 5.00
Begin VB.Form InsertHTMLDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert HTML"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6540
      TabIndex        =   3
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6540
      TabIndex        =   2
      Top             =   660
      Width           =   1215
   End
   Begin VB.TextBox HTMLText 
      Height          =   3225
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "HTML source:"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7035
   End
End
Attribute VB_Name = "InsertHTMLDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1998 Microsoft Corporation.
' All rights reserved.
Private Sub CmdCancel_Click()
    Unload Me
    OpenDocuments(CStr(ActiveDocument)).DHTMLEdit1.SetFocus
End Sub

Private Sub cmdOK_Click()

    Dim doc As Object
    Dim sel As Object
    Dim tr As Object
    
    ' get the DHTML Document object
    Set doc = OpenDocuments(CStr(ActiveDocument)).DHTMLEdit1.DOM
    ' get the IE4 selection object
    Set sel = doc.selection
    ' create a TextRange from the current selection
    Set tr = sel.createRange
    
    ' paste our html into the range
    tr.pasteHTML (HTMLText.Text)
    Unload Me
End Sub

