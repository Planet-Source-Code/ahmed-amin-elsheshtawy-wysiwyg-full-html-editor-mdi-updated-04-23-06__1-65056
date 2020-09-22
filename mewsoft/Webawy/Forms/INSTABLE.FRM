VERSION 5.00
Begin VB.Form InsertTableDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Table"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "INSTABLE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TableCaption 
      Height          =   285
      Left            =   1980
      TabIndex        =   11
      Top             =   2400
      Width           =   5295
   End
   Begin VB.TextBox CellAttrs 
      Height          =   285
      Left            =   1980
      TabIndex        =   9
      Top             =   1860
      Width           =   5295
   End
   Begin VB.TextBox TableAttrs 
      Height          =   285
      Left            =   1980
      TabIndex        =   7
      Top             =   1440
      Width           =   5295
   End
   Begin VB.TextBox Cols 
      Height          =   285
      Left            =   1980
      TabIndex        =   5
      Top             =   900
      Width           =   1035
   End
   Begin VB.TextBox Rows 
      Height          =   285
      Left            =   1980
      TabIndex        =   3
      Top             =   360
      Width           =   1035
   End
   Begin VB.CommandButton CancelCmd 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3900
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton OkCmd 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label CaptionLabel 
      Caption         =   "Caption:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label CellTagLabel 
      Caption         =   "Cell Tag Attributes:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1860
      Width           =   1815
   End
   Begin VB.Label TableTagLabel 
      Caption         =   "Table Tag Attributes:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label ColLabel 
      Caption         =   "Number of columns:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   1455
   End
   Begin VB.Label RowLabel 
      Caption         =   "Number of rows:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   1455
   End
End
Attribute VB_Name = "InsertTableDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1998 Microsoft Corporation.
' All rights reserved.
Private tableParam As DEInsertTableParam

Private Sub CancelCmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' create the table parameter object
    Set tableParam = CreateObject("DEInsertTableParam.DEInsertTableParam.1")
    
'    Rows = tableParam.NumRows
'    Cols = tableParam.NumCols
'    TableAttrs = tableParam.TableAttrs
    Rows = 4
    Cols = 4
    TableAttrs = "border=""1"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#808080"" width=""80%"""
    
    CellAttrs = tableParam.CellAttrs
    TableCaption = tableParam.Caption

End Sub

Private Sub OkCmd_Click()
    
    On Error Resume Next
    
    If Rows = "" Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    ElseIf IsNumeric(Rows) = False Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    ElseIf Rows <= 0 Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    End If
       
    If Cols = "" Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    ElseIf IsNumeric(Cols) = False Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    ElseIf Cols <= 0 Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    End If
    
    tableParam.NumRows = Rows
    tableParam.NumCols = Cols
    
    If Len(TableAttrs.Text) Then
        tableParam.TableAttrs = TableAttrs.Text
    End If
    If Len(CellAttrs.Text) Then
        tableParam.CellAttrs = CellAttrs.Text
    End If
    If Len(TableCaption.Text) Then
        tableParam.Caption = TableCaption.Text
    End If
    
    'Document(ActiveDocument).DHTMLEdit1.ExecCommand DECMD_INSERTTABLE, OLECMDEXECOPT_DONTPROMPTUSER, tableParam
    OpenDocuments(CStr(ActiveDocument)).DHTMLEdit1.ExecCommand DECMD_INSERTTABLE, OLECMDEXECOPT_DONTPROMPTUSER, tableParam
    
    Unload Me
End Sub

