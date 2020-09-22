VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin TabDlg.SSTab SSTab1 
      Height          =   5715
      Left            =   180
      TabIndex        =   9
      Top             =   240
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   10081
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "general Settings"
      TabPicture(0)   =   "frmOptions.frx":628A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "options"
      TabPicture(1)   =   "frmOptions.frx":62A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture2 
         Height          =   5175
         Left            =   60
         ScaleHeight     =   5115
         ScaleWidth      =   8775
         TabIndex        =   12
         Top             =   420
         Width           =   8835
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   4875
            Left            =   120
            TabIndex        =   13
            Top             =   180
            Width           =   8595
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   5235
         Left            =   -74880
         ScaleHeight     =   5175
         ScaleWidth      =   8595
         TabIndex        =   10
         Top             =   360
         Width           =   8655
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            Height          =   4995
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   8415
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2970
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   6075
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Tag             =   "Cancel"
      Top             =   6075
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Tag             =   "&Apply"
      Top             =   6075
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   8
         Tag             =   "Sample 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   7
         Tag             =   "Sample 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   5
         Tag             =   "Sample 2"
         Top             =   305
         Width           =   2033
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    'ToDo: Add 'cmdApply_Click' code.
    MsgBox "Apply Code goes here to set options w/o closing dialog!"
End Sub


Private Sub CmdCancel_Click()
    Unload Me
End Sub


