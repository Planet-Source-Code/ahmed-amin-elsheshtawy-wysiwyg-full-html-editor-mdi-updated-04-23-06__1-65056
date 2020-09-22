VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Webawy"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9690
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   5580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":628A
            Key             =   "Bold"
            Object.Tag             =   "Bold"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6464
            Key             =   "Underline"
            Object.Tag             =   "Underline"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":663E
            Key             =   "Italic"
            Object.Tag             =   "Italic"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6818
            Key             =   "LeftJustify"
            Object.Tag             =   "LeftJustify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69F2
            Key             =   "RightJustify"
            Object.Tag             =   "RightJustify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6BCC
            Key             =   "CenterJustify"
            Object.Tag             =   "CenterJustify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DA6
            Key             =   "FullJustify"
            Object.Tag             =   "FullJustify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F80
            Key             =   "Bullets"
            Object.Tag             =   "Bullets"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":715A
            Key             =   "Numbers"
            Object.Tag             =   "Numbers"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7334
            Key             =   "Indent"
            Object.Tag             =   "Indent"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":750E
            Key             =   "Outdent"
            Object.Tag             =   "Outdent"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":76E8
            Key             =   "LTR"
            Object.Tag             =   "LTR"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":78C2
            Key             =   "SubScript"
            Object.Tag             =   "SubScript"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A1C
            Key             =   "SuperScript"
            Object.Tag             =   "SuperScript"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B76
            Key             =   "StrikeThrough"
            Object.Tag             =   "StrikeThrough"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D50
            Key             =   "RTL"
            Object.Tag             =   "RTL"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F2A
            Key             =   "ForeColor1"
            Object.Tag             =   "ForeColor1"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8084
            Key             =   "ForeColor"
            Object.Tag             =   "ForeColor"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":81DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8778
            Key             =   "BackColor"
            Object.Tag             =   "BackColor"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   180
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D12
            Key             =   "UpDown"
            Object.Tag             =   "UpDown"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   180
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8EEC
            Key             =   "WebFile"
            Object.Tag             =   "WebFile"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F186
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbFilesBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   27
      Top             =   1080
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UpDown"
            Description     =   "UpDown"
            Object.ToolTipText     =   "Position This Toolbar Up or Down"
            ImageKey        =   "UpDown"
            Object.Width           =   1e7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileBar"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.TabStrip tabFilesBar 
         Height          =   375
         Left            =   3180
         TabIndex        =   28
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Style           =   2
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         TabMinWidth     =   88
         ImageList       =   "ImageList4"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picRight 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4980
      Left            =   5925
      ScaleHeight     =   4980
      ScaleMode       =   0  'User
      ScaleWidth      =   372.957
      TabIndex        =   7
      Top             =   1440
      Width           =   3765
      Begin VB.PictureBox picRightPanTitle 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   3195
         TabIndex        =   8
         Top             =   120
         Width           =   3255
         Begin VB.CommandButton cmdRightPanRL 
            Caption         =   "<>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   10
            ToolTipText     =   "Move this panel to the other side"
            Top             =   60
            Width           =   375
         End
         Begin VB.CommandButton cmdCloseRightPan 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   9
            ToolTipText     =   "Close this panel"
            Top             =   60
            Width           =   255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tools"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   60
            Width           =   1035
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3945
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   6959
         _Version        =   393216
         Style           =   1
         Tab             =   1
         TabHeight       =   520
         TabMaxWidth     =   1781
         WordWrap        =   0   'False
         MouseIcon       =   "frmMain.frx":F2E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   " "
         TabPicture(0)   =   "frmMain.frx":F2FC
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "FrameFiles"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   " "
         TabPicture(1)   =   "frmMain.frx":F456
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "FrameProperties"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " "
         TabPicture(2)   =   "frmMain.frx":F5B0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frameTree"
         Tab(2).ControlCount=   1
         Begin VB.Frame FrameProperties 
            Caption         =   "Tag Properties"
            Height          =   2895
            Left            =   120
            TabIndex        =   23
            Top             =   420
            Width           =   3135
            Begin MSComctlLib.ListView lvProperties 
               Height          =   1215
               Left            =   180
               TabIndex        =   26
               Top             =   1500
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   2143
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin VB.TextBox txtProperty 
               Height          =   795
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   25
               Top             =   540
               Width           =   2895
            End
            Begin VB.Label lbProperty 
               Caption         =   "lbProperty"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   1635
            End
         End
         Begin VB.Frame FrameFiles 
            Caption         =   "Files List"
            Height          =   4215
            Left            =   -74880
            TabIndex        =   18
            Top             =   360
            Width           =   3195
            Begin VB.FileListBox File1 
               Height          =   1260
               Left            =   240
               TabIndex        =   22
               Top             =   2460
               Width           =   2655
            End
            Begin VB.ComboBox cboFilesFilter 
               Height          =   315
               Left            =   180
               TabIndex        =   21
               Text            =   "Combo1"
               Top             =   2040
               Width           =   2655
            End
            Begin VB.DirListBox Dir1 
               Height          =   1215
               Left            =   180
               TabIndex        =   20
               Top             =   660
               Width           =   2775
            End
            Begin VB.DriveListBox Drive1 
               Height          =   315
               Left            =   180
               TabIndex        =   19
               Top             =   240
               Width           =   2715
            End
         End
         Begin VB.Frame frameTree 
            Caption         =   "HTML Tree"
            Height          =   2175
            Left            =   -74940
            TabIndex        =   13
            Top             =   420
            Width           =   3135
            Begin MSComctlLib.TreeView tvDomTree 
               Height          =   1155
               Left            =   -300
               TabIndex        =   17
               Top             =   720
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   2037
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   88
               LabelEdit       =   1
               Style           =   7
               Appearance      =   1
            End
            Begin VB.CheckBox chkAttributes 
               Caption         =   "Attributes"
               Height          =   255
               Left            =   1920
               TabIndex        =   16
               Top             =   300
               Width           =   1035
            End
            Begin VB.CheckBox chkExpandAll 
               Caption         =   "+/-"
               Height          =   255
               Left            =   1260
               TabIndex        =   15
               Top             =   300
               Width           =   615
            End
            Begin VB.CommandButton cmdTreeRefresh 
               Caption         =   "Refresh Tree"
               Height          =   255
               Left            =   60
               TabIndex        =   14
               Top             =   300
               Width           =   1095
            End
         End
      End
   End
   Begin VB.Timer TimerMemoryStatus 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   5760
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   900
      Top             =   5700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6420
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7162
            MinWidth        =   71
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2461
            MinWidth        =   2470
            Object.ToolTipText     =   "Character Postion"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Object.ToolTipText     =   "Available Physical Memory"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   635
            MinWidth        =   88
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   88
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   820
            MinWidth        =   88
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   240
      Top             =   4620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   63
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FB4A
            Key             =   "New2"
            Object.Tag             =   "New2"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FD24
            Key             =   "Open2"
            Object.Tag             =   "Open2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FEFE
            Key             =   "Save2"
            Object.Tag             =   "Save2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":100D8
            Key             =   "SaveAs1"
            Object.Tag             =   "SaveAs1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1062A
            Key             =   "Print"
            Object.Tag             =   "Print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10804
            Key             =   "Preview"
            Object.Tag             =   "Preview"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":109DE
            Key             =   "Spell"
            Object.Tag             =   "Spell"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10BB8
            Key             =   "Cut1"
            Object.Tag             =   "Cut1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10D92
            Key             =   "Copy1"
            Object.Tag             =   "Copy1"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10F6C
            Key             =   "Paste1"
            Object.Tag             =   "Paste1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11146
            Key             =   "Undo1"
            Object.Tag             =   "Undo1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":112A0
            Key             =   "Redo1"
            Object.Tag             =   "Redo1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":113FA
            Key             =   "Table1"
            Object.Tag             =   "Table1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":115D4
            Key             =   "Image2"
            Object.Tag             =   "Image2"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":117AE
            Key             =   "Link"
            Object.Tag             =   "Link"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11988
            Key             =   "ShowAll"
            Object.Tag             =   "ShowAll"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11B62
            Key             =   "DeleteCells"
            Object.Tag             =   "DeleteCells"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11D3C
            Key             =   "InsertColumns"
            Object.Tag             =   "InsertColumns"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11F16
            Key             =   "InsertRows"
            Object.Tag             =   "InsertRows"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":120F0
            Key             =   "DeleteColumns"
            Object.Tag             =   "DeleteColumns"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":122CA
            Key             =   "ShowBorders"
            Object.Tag             =   "ShowBorders"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":124A4
            Key             =   "HideBorders"
            Object.Tag             =   "HideBorders"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1267E
            Key             =   "ColsEven"
            Object.Tag             =   "ColsEven"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12858
            Key             =   "RowsEven"
            Object.Tag             =   "RowsEven1"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12A32
            Key             =   "Download"
            Object.Tag             =   "Download"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12B8C
            Key             =   "MergeCells"
            Object.Tag             =   "MergeCells"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12D66
            Key             =   "SplitCells"
            Object.Tag             =   "SplitCells"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12F40
            Key             =   "Video"
            Object.Tag             =   "Video"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1311A
            Key             =   "PageSetup"
            Object.Tag             =   "PageSetup"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":132F4
            Key             =   "PrintPreview"
            Object.Tag             =   "PrintPreview"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":134CE
            Key             =   "Properties"
            Object.Tag             =   "Properties"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":136A8
            Key             =   "Publish"
            Object.Tag             =   "Publish"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13882
            Key             =   "WebTransfer"
            Object.Tag             =   "WebTransfer"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":139DC
            Key             =   "Find"
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13B36
            Key             =   "AlignBottom"
            Object.Tag             =   "AlignBottom"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13D10
            Key             =   "AlignTop."
            Object.Tag             =   "AlignTop."
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13EEA
            Key             =   "BringForward"
            Object.Tag             =   "BringForward"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":140C4
            Key             =   "BringToFront"
            Object.Tag             =   "BringToFront"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1429E
            Key             =   "SendBackward"
            Object.Tag             =   "SendBackward"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14478
            Key             =   "SendToBack"
            Object.Tag             =   "SendToBack"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14652
            Key             =   "TextDirection"
            Object.Tag             =   "TextDirection"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1482C
            Key             =   "AutoFit"
            Object.Tag             =   "AutoFit"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14A06
            Key             =   "Comment"
            Object.Tag             =   "Comment"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14BE0
            Key             =   "Website"
            Object.Tag             =   "Website"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15032
            Key             =   "TableAutoFormat"
            Object.Tag             =   "TableAutoFormat"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1520C
            Key             =   "Form"
            Object.Tag             =   "Form"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153E6
            Key             =   "Open"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15AB8
            Key             =   "New"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1618A
            Key             =   "Save1"
            Object.Tag             =   "Save1"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16364
            Key             =   "SnapToGrid"
            Object.Tag             =   "SnapToGrid"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1653E
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16C10
            Key             =   "SaveAs"
            Object.Tag             =   "SaveAs"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":172E2
            Key             =   "Copy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":179B4
            Key             =   "Cut"
            Object.Tag             =   "Cut"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17B0E
            Key             =   "Paste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":181E0
            Key             =   "Undo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1833A
            Key             =   "Redo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18494
            Key             =   "Replace"
            Object.Tag             =   "Replace"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18B66
            Key             =   "Find1"
            Object.Tag             =   "Find1"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18CC0
            Key             =   "Image"
            Object.Tag             =   "Image"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19392
            Key             =   "Table"
            Object.Tag             =   "Table"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":194EC
            Key             =   "FindNext"
            Object.Tag             =   "FindNext"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19646
            Key             =   "GoToLine"
            Object.Tag             =   "GoToLine"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   120
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":197A0
            Key             =   "Normal"
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1997A
            Key             =   "HTML"
            Object.Tag             =   "HTML"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19D2E
            Key             =   "Preview"
            Object.Tag             =   "Preview"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19E88
            Key             =   "Refresh"
            Object.Tag             =   "Refresh"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A062
            Key             =   "Back"
            Object.Tag             =   "Back"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A23C
            Key             =   "Forword"
            Object.Tag             =   "Forword"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A416
            Key             =   "Stop"
            Object.Tag             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A5F0
            Key             =   "InsertCells"
            Object.Tag             =   "InsertCells"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A7CA
            Key             =   "InsertColumns"
            Object.Tag             =   "InsertColumns"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A9A4
            Key             =   "InsertRows2"
            Object.Tag             =   "InsertRows2"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AB7E
            Key             =   "MergeCells"
            Object.Tag             =   "MergeCells"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AD58
            Key             =   "SplitCells"
            Object.Tag             =   "SplitCells"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AF32
            Key             =   "DeleteCells"
            Object.Tag             =   "DeleteCells"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B10C
            Key             =   "DeleteColumns"
            Object.Tag             =   "DeleteColumns"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B2E6
            Key             =   "DeleteRows"
            Object.Tag             =   "DeleteRows"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B440
            Key             =   "InsertRows"
            Object.Tag             =   "InsertRows"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B59A
            Key             =   "PositionAbsolutely"
            Object.Tag             =   "PositionAbsolutely"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B774
            Key             =   "BringForward"
            Object.Tag             =   "BringForward"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B94E
            Key             =   "SendBackward"
            Object.Tag             =   "SendBackward"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BB28
            Key             =   "BringToFront"
            Object.Tag             =   "BringToFront"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BD02
            Key             =   "SendBackward1-delete"
            Object.Tag             =   "SendBackward1"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BEDC
            Key             =   "SendToBack"
            Object.Tag             =   "SendToBack"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C0B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C5C8
            Key             =   "Textbox"
            Object.Tag             =   "Textbox"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C7D5
            Key             =   "Textarea"
            Object.Tag             =   "Textarea"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C9DA
            Key             =   "Checkbox"
            Object.Tag             =   "Checkbox"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CBDF
            Key             =   "OptionButton"
            Object.Tag             =   "OptionButton"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CDD9
            Key             =   "DropDown"
            Object.Tag             =   "DropDown"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CFE2
            Key             =   "PushButton"
            Object.Tag             =   "PushButton"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D1E0
            Key             =   "HiddenData"
            Object.Tag             =   "HiddenData"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D3BA
            Key             =   "Password"
            Object.Tag             =   "Password"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D594
            Key             =   "SubmitButton"
            Object.Tag             =   "SubmitButton"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D76E
            Key             =   "ResetButton"
            Object.Tag             =   "ResetButton"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D948
            Key             =   "ImageButton"
            Object.Tag             =   "ImageButton"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DB22
            Key             =   "Form"
            Object.Tag             =   "Form"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DCFC
            Key             =   "BringAboveText"
            Object.Tag             =   "BringAboveText"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DED6
            Key             =   "SendBelowText"
            Object.Tag             =   "SendBelowText"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E0B0
            Key             =   "SnapToGrid"
            Object.Tag             =   "SnapToGrid"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E28A
            Key             =   "ListBox"
            Object.Tag             =   "ListBox"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar3 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   39
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Normal"
            Description     =   "Normal"
            Object.ToolTipText     =   "Normal"
            Object.Tag             =   "Normal"
            ImageKey        =   "Normal"
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HTML"
            Description     =   "HTML"
            Object.ToolTipText     =   "HTML"
            Object.Tag             =   "HTML"
            ImageKey        =   "HTML"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Preview"
            Description     =   "Preview"
            Object.ToolTipText     =   "Preview"
            Object.Tag             =   "Preview"
            ImageKey        =   "Preview"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            Object.Tag             =   "Refresh"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            Object.Tag             =   "Stop"
            ImageKey        =   "Stop"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertRows"
            Object.ToolTipText     =   "Insert Rows"
            Object.Tag             =   "Insert Rows"
            ImageKey        =   "InsertRows"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertColumns"
            Object.ToolTipText     =   "Insert Columns"
            Object.Tag             =   "Insert Columns"
            ImageKey        =   "InsertColumns"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertCells"
            Object.ToolTipText     =   "Insert Cells"
            Object.Tag             =   "Insert Cells"
            ImageKey        =   "InsertCells"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteCells"
            Object.ToolTipText     =   "Delete Cells"
            Object.Tag             =   "Delete Cells"
            ImageKey        =   "DeleteCells"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteRows"
            Object.ToolTipText     =   "Delete Rows"
            Object.Tag             =   "DeleteRows"
            ImageKey        =   "DeleteRows"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteColumns"
            Object.ToolTipText     =   "Delete Columns"
            Object.Tag             =   "Delete Columns"
            ImageKey        =   "DeleteColumns"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MergeCells"
            Object.ToolTipText     =   "Merge Cells"
            Object.Tag             =   "Merge Cells"
            ImageKey        =   "MergeCells"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SplitCells"
            Object.ToolTipText     =   "Split Cells"
            Object.Tag             =   "Split Cells"
            ImageKey        =   "SplitCells"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PositionAbsolutely"
            Object.ToolTipText     =   "Position Absolutely"
            Object.Tag             =   "Position Absolutely"
            ImageKey        =   "PositionAbsolutely"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   12
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PositionAbsolutely"
                  Object.Tag             =   "Position Absolutley"
                  Text            =   "Position Absolutley"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BringForward"
                  Object.Tag             =   "Bring Forward"
                  Text            =   "Bring Forward"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SendBackward"
                  Object.Tag             =   "Send Backward"
                  Text            =   "Send Backward"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BringToFront"
                  Object.Tag             =   "Bring To Front"
                  Text            =   "Bring To Front"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SendToBack"
                  Object.Tag             =   "Send To Back"
                  Text            =   "Send To Back"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BringAboveText"
                  Object.Tag             =   "Bring Above Text"
                  Text            =   "Bring Above Text"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SendBelowText"
                  Object.Tag             =   "Send Below Text"
                  Text            =   "Send Below Text"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LockElement"
                  Object.Tag             =   "Lock Element"
                  Text            =   "Lock Element"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BringForward"
            Object.ToolTipText     =   "Bring Forward"
            Object.Tag             =   "Bring Forward"
            ImageKey        =   "BringForward"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SendBackward"
            Object.ToolTipText     =   "Send Backward"
            Object.Tag             =   "Send Backward"
            ImageKey        =   "SendBackward"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BringToFront"
            Object.ToolTipText     =   "Bring To Front"
            Object.Tag             =   "Bring To Front"
            ImageKey        =   "BringToFront"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SendToBack"
            Object.ToolTipText     =   "Send To Back"
            Object.Tag             =   "Send To Back"
            ImageKey        =   "SendToBack"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BringAboveText"
            Object.ToolTipText     =   "Bring Above Text"
            Object.Tag             =   "Bring Above Text"
            ImageKey        =   "BringAboveText"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SendBelowText"
            Object.ToolTipText     =   "Send Below Text"
            Object.Tag             =   "Send Below Text"
            ImageKey        =   "SendBelowText"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SnapToGrid"
            Object.ToolTipText     =   "Snap to Grid"
            Object.Tag             =   "Snap to Grid"
            ImageKey        =   "SnapToGrid"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Textbox"
            Object.ToolTipText     =   "Textbox"
            Object.Tag             =   "Textbox"
            ImageKey        =   "Textbox"
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Textarea"
            Object.ToolTipText     =   "Textarea"
            Object.Tag             =   "Textarea"
            ImageKey        =   "Textarea"
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Checkbox"
            Object.ToolTipText     =   "Checkbox"
            Object.Tag             =   "Checkbox"
            ImageKey        =   "Checkbox"
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OptionButton"
            Object.ToolTipText     =   "Option Button"
            Object.Tag             =   "Option Button"
            ImageKey        =   "OptionButton"
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ListBox"
            Object.ToolTipText     =   "ListBox"
            Object.Tag             =   "ListBox"
            ImageKey        =   "ListBox"
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DropDownBox"
            Object.ToolTipText     =   "Drop Down Box"
            Object.Tag             =   "Drop Down Box"
            ImageKey        =   "DropDown"
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PushButton"
            Object.ToolTipText     =   "Push Button"
            Object.Tag             =   "Push Button"
            ImageKey        =   "PushButton"
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HiddenData"
            Object.ToolTipText     =   "Hidden Data"
            Object.Tag             =   "Hidden Data"
            ImageKey        =   "HiddenData"
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Password"
            Object.ToolTipText     =   "Password"
            Object.Tag             =   "Password"
            ImageKey        =   "Password"
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SubmitButton"
            Object.ToolTipText     =   "Submit Button"
            Object.Tag             =   "Submit Button"
            ImageKey        =   "SubmitButton"
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ResetButton"
            Object.ToolTipText     =   "Reset Button"
            Object.Tag             =   "Reset Button"
            ImageKey        =   "ResetButton"
         EndProperty
         BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ImageButton"
            Object.ToolTipText     =   "Image Button"
            Object.Tag             =   "Image Button"
            ImageKey        =   "ImageButton"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Style"
            Description     =   "Style"
            Object.ToolTipText     =   "Style"
            Object.Tag             =   "Style"
            Style           =   4
            Object.Width           =   5000
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Description     =   "Bold"
            Object.ToolTipText     =   "Bold"
            Object.Tag             =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Description     =   "Italic"
            Object.ToolTipText     =   "Italic"
            Object.Tag             =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Description     =   "Underline"
            Object.ToolTipText     =   "Underline"
            Object.Tag             =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LeftJustify"
            Description     =   "Left Justify"
            Object.ToolTipText     =   "Left Justify"
            ImageKey        =   "LeftJustify"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Description     =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "CenterJustify"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RightJustify"
            Description     =   "Right Justify"
            Object.ToolTipText     =   "Right Justify"
            ImageKey        =   "RightJustify"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "JustifyFull"
            Object.ToolTipText     =   "Justify Full"
            Object.Tag             =   "Justify Full"
            ImageKey        =   "FullJustify"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SuperScript"
            Object.ToolTipText     =   "Superscript"
            Object.Tag             =   "Superscript"
            ImageKey        =   "SuperScript"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SubScript"
            Object.ToolTipText     =   "Subscript"
            Object.Tag             =   "Subscript"
            ImageKey        =   "SubScript"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "StrikeThrough"
            Object.ToolTipText     =   "Strike through"
            Object.Tag             =   "Strike through"
            ImageKey        =   "StrikeThrough"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Numbers"
            Description     =   "Numbers"
            Object.ToolTipText     =   "Numbers"
            Object.Tag             =   "Numbers"
            ImageKey        =   "Numbers"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bullets"
            Description     =   "Bullets"
            Object.ToolTipText     =   "Bullets"
            ImageKey        =   "Bullets"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Outdent"
            Description     =   "Outdent"
            ImageKey        =   "Outdent"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Indent"
            Description     =   "Indent"
            Object.ToolTipText     =   "Indent"
            ImageKey        =   "Indent"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Color"
            Description     =   "Color"
            Object.ToolTipText     =   "Font Color"
            ImageKey        =   "ForeColor"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BackColor"
            Description     =   "BackColor"
            Object.ToolTipText     =   "Background Color"
            ImageKey        =   "BackColor"
         EndProperty
      EndProperty
      Begin VB.ComboBox FontCombo 
         Height          =   315
         Left            =   1740
         TabIndex        =   5
         Text            =   "FontCombo"
         ToolTipText     =   "Font"
         Top             =   0
         Width           =   2355
      End
      Begin VB.ComboBox FontSizeCombo 
         Height          =   315
         Left            =   4140
         TabIndex        =   4
         Text            =   "Combo1"
         ToolTipText     =   "Font size"
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox StyleCombo 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   0
         Width           =   1650
      End
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   27
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "New"
            Object.ToolTipText     =   "New"
            Object.Tag             =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveAs"
            Object.ToolTipText     =   "Save As"
            Object.Tag             =   "Save As"
            ImageKey        =   "SaveAs"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Download"
            Object.ToolTipText     =   "Open URL"
            Object.Tag             =   "Download"
            ImageKey        =   "Download"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print"
            Object.ToolTipText     =   "Print"
            Object.Tag             =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View"
            Description     =   "View"
            Object.ToolTipText     =   "View in Browser"
            Object.Tag             =   "View"
            ImageKey        =   "Preview"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SpellCheck"
            Description     =   "Spell Check"
            Object.ToolTipText     =   "Spell Check"
            Object.Tag             =   "Spell"
            ImageKey        =   "Spell"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            Object.Tag             =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut"
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste"
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Description     =   "Undo"
            Object.ToolTipText     =   "Undo"
            Object.Tag             =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Description     =   "Redo"
            Object.ToolTipText     =   "Redo"
            Object.Tag             =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Table"
            Description     =   "Table"
            Object.ToolTipText     =   "Insert Table"
            Object.Tag             =   "Table"
            ImageKey        =   "Table"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Image"
            Description     =   "Image"
            Object.ToolTipText     =   "Insert Image"
            Object.Tag             =   "Image"
            ImageKey        =   "Image"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Hyperlink"
            Description     =   "Hyperlink"
            Object.ToolTipText     =   "Insert Hyperlink"
            Object.Tag             =   "Hyperlink"
            ImageKey        =   "Link"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Form"
            Object.ToolTipText     =   "Insert Form"
            Object.Tag             =   "Form"
            ImageKey        =   "Form"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   16
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Form"
                  Object.Tag             =   "Form"
                  Text            =   "Form"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Textbox"
                  Object.Tag             =   "Textbox"
                  Text            =   "Textbox"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Textarea"
                  Object.Tag             =   "Textarea"
                  Text            =   "Text Area"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Checkbox"
                  Object.Tag             =   "Checkbox"
                  Text            =   "Checkbox"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OptionButton"
                  Object.Tag             =   "Option Button"
                  Text            =   "Option Button"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DropDownBox"
                  Object.Tag             =   "Drop-Down Box"
                  Text            =   "Drop-Down Box"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PushButton"
                  Object.Tag             =   "Push Button"
                  Text            =   "Push Button"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AdvancedButton"
                  Object.Tag             =   "Advanced Button"
                  Text            =   "Advanced Button"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ImageButton"
                  Object.Tag             =   "Picture"
                  Text            =   "Picture"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "FileUpload"
                  Object.Tag             =   "File Upload"
                  Text            =   "File Upload"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "HiddenData"
                  Object.Tag             =   "Hidden Data"
                  Text            =   "Hidden Data"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Password"
                  Object.Tag             =   "Password"
                  Text            =   "Password"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SubmitButton"
                  Object.Tag             =   "Submit Button"
                  Text            =   "Submit Button"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ResetButton"
                  Object.Tag             =   "Reset Button"
                  Text            =   "Reset Button"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowAll"
            Description     =   "ShowAll"
            Object.ToolTipText     =   "Show All"
            Object.Tag             =   "ShowAll"
            ImageKey        =   "ShowAll"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowBorders"
            Description     =   "ShowBorders"
            Object.ToolTipText     =   "Show Borders"
            Object.Tag             =   "ShowBorders"
            ImageKey        =   "ShowBorders"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            Object.Tag             =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileCloseAll 
         Caption         =   "Clos&e All"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Sen&d..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsWindow 
         Caption         =   "Tools Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowBorders 
         Caption         =   "Show Borders"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "Document Details"
      End
      Begin VB.Menu mnuView2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "&Web Browser"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuBreak 
         Caption         =   "Break"
      End
      Begin VB.Menu mnuHorizontalLine 
         Caption         =   "Horizontal Line"
      End
      Begin VB.Menu mnuInlineFrame 
         Caption         =   "Inline Frame"
      End
      Begin VB.Menu mnuComment 
         Caption         =   "Comment"
      End
      Begin VB.Menu mnuInsert0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPicture 
         Caption         =   "Picture..."
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   "Bookmark"
      End
      Begin VB.Menu mnuHyperlink 
         Caption         =   "Hyperlink"
      End
      Begin VB.Menu mnuInsert1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertHTML 
         Caption         =   "HTML Source"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuFont 
         Caption         =   "&Font"
      End
      Begin VB.Menu mnuFont0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuItalic 
         Caption         =   "Italic"
      End
      Begin VB.Menu mnuUnderline 
         Caption         =   "Underline"
      End
      Begin VB.Menu mnuFont1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "Left"
      End
      Begin VB.Menu mnuCenter 
         Caption         =   "Center"
      End
      Begin VB.Menu mnuRight 
         Caption         =   "Right"
      End
      Begin VB.Menu mnuJustify 
         Caption         =   "Justify"
      End
      Begin VB.Menu mnuFont2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Color"
      End
      Begin VB.Menu mnuFont3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNumberedList 
         Caption         =   "Numbered List"
      End
      Begin VB.Menu mnuBullets 
         Caption         =   "Bullets"
      End
      Begin VB.Menu mnuFont4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndent 
         Caption         =   "Indent"
      End
      Begin VB.Menu mnuOutdent 
         Caption         =   "Outdent"
      End
   End
   Begin VB.Menu mnuPosition 
      Caption         =   "&Postion"
      Begin VB.Menu mnuPositionAbsolutly 
         Caption         =   "Position Absolutly"
      End
      Begin VB.Menu mnuPositionAbsolutly0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBringForward 
         Caption         =   "Bring Forward"
      End
      Begin VB.Menu mnuSendBackward 
         Caption         =   "Send Backward"
      End
      Begin VB.Menu mnuPositionAbsolutly1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBringToFront 
         Caption         =   "Bring To Front"
      End
      Begin VB.Menu mnuSendToBack 
         Caption         =   "Send To Back"
      End
      Begin VB.Menu mnuPositionAbsolutly2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBringAboveText 
         Caption         =   "Bring Above Text"
      End
      Begin VB.Menu mnuSendBelowText 
         Caption         =   "Send Below Text"
      End
      Begin VB.Menu mnuPositionAbsolutly3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLockElement 
         Caption         =   "Lock Element"
      End
      Begin VB.Menu mnuPositionAbsolutly4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSnapToGrid 
         Caption         =   "Snap to Grid"
      End
   End
   Begin VB.Menu mnuTable 
      Caption         =   "&Table"
      Begin VB.Menu mnuInsertTable 
         Caption         =   "&Insert Table"
      End
      Begin VB.Menu mnuInsertTable0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertRows 
         Caption         =   "Insert &Rows"
      End
      Begin VB.Menu mnuInsertColumns 
         Caption         =   "Insert Columns"
      End
      Begin VB.Menu mnuInsertCells 
         Caption         =   "Insert Cells"
      End
      Begin VB.Menu mnuInsertTable1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteRows 
         Caption         =   "Delete Rows"
      End
      Begin VB.Menu mnuDeleteColumns 
         Caption         =   "Delete Columns"
      End
      Begin VB.Menu mnuDeleteCell 
         Caption         =   "Delete Cells"
      End
      Begin VB.Menu mnuInsertTable2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMergeCells 
         Caption         =   "&Merge Cells"
      End
      Begin VB.Menu mnuSplitCells 
         Caption         =   "&Split Cell"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help Online"
      End
      Begin VB.Menu mnuOnlineSupport 
         Caption         =   "Support"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private m_cSplitLeft As New cMDISplit
Private m_cSplitBottom As New cMDISplit
Private m_cSplitTop As New cMDISplit
Private m_cSplitRight As New cMDISplit

Private Sub MDIForm_Load()
    
    On Error Resume Next
    '---------------------------------------
    Me.left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    '---------------------------------------
    mnuToolsWindow.Checked = GetSetting(App.Title, "Settings", "RightPan", True)
    
    SSTab1.Tab = GetSetting(App.Title, "Settings", "RightPanTab", 0)
    
    chkExpandAll.Value = GetSetting(App.Title, "Settings", "chkExpandAll", 1)
    chkAttributes.Value = GetSetting(App.Title, "Settings", "chkAttributes", 0)
    
    tbFilesBar.Align = GetSetting(App.Title, "Settings", "FilesBarAlign", 2)
    '---------------------------------------
    picRight.ScaleMode = vbTwips
    picRightPanTitle.ScaleMode = vbTwips
    Dim a As Long
    a = GetSetting(App.Title, "Settings", "RightPanWidth", 3700)
    picRight.Width = a
    
    picRight.Align = GetSetting(App.Title, "Settings", "RightPanAlign", 4)
    picRight.Visible = GetSetting(App.Title, "Settings", "RightPan", True)
    
    picRight_Resize
    '---------------------------------------
    ' Fill the font list with the system fonts
    ReDim fontNames(0 To Screen.FontCount - 1)
    Dim x As Integer
    
    For x = 0 To Screen.FontCount - 1
        fontNames(x) = Screen.Fonts(x)
    Next x
    
    SortArray fontNames()
    
    For x = 0 To Screen.FontCount - 1
        FontCombo.AddItem fontNames(x)
    Next x

    fMainForm.FontCombo.ListIndex = 0
    '---------------------------------------
    ' Font size menu
    For x = 1 To 7
        fMainForm.FontSizeCombo.AddItem CStr(x)
    Next x
    
    fMainForm.FontSizeCombo.ListIndex = 0
    '---------------------------------------
    '---------------------------------------
    'Document counter
    lDocumentCount = 0
    'New documents counter
    NewDocumentCount = 0
    
    'Open a new document
    LoadNewDoc
    cmdTreeRefresh_Click
    
    'Enable memory status timer
    TimerMemoryStatus.Enabled = True
    '---------------------------------------
    cboFilesFilter.AddItem "*.htm;*.html;*.shtml;*.shtm;*.stm;*.asp;*.aspx;*.css"
    cboFilesFilter.AddItem "*.txt"
    cboFilesFilter.AddItem "*.*"
    cboFilesFilter.ListIndex = 0
    '---------------------------------------
    ' Properties Listview
    Dim ColHeader As ColumnHeader
    
    Set ColHeader = lvProperties.ColumnHeaders.Add()
    ColHeader.Text = "Property"
    'ColHeader.Width = 1500
    
    Set ColHeader = lvProperties.ColumnHeaders.Add()
    ColHeader.Text = "Value"
    'ColHeader.Width = 1000
    Set ColHeader = lvProperties.ColumnHeaders.Add()
    ColHeader.Text = "#"
    'ColHeader.Width = 1000
    '---------------------------------------
    m_cSplitRight.Attach picRight
    'm_cSplitLeft.MaxSize = 128
    'm_cSplitRight.FullDrag = True
    'm_cSplitLeft.SplitSize
    '---------------------------------------
    tbFilesBar.ButtonWidth = 100
    picRightPanTitleGradient
    tabFilesBar.Height = tbFilesBar.Height
    '---------------------------------------
    'SSTab1.Tabs.Item(1).ToolTipText = "Files"
    
End Sub

Private Sub mnuBold_Click()

    ExecAndUpdateToolbarButton1 "Bold", DECMD_BOLD
    
'
'        Case "Italic"
'            ExecAndUpdateToolbarButton1 "Italic", DECMD_ITALIC
'
'        Case "Underline"
'            ExecAndUpdateToolbarButton1 "Underline", DECMD_UNDERLINE
'
'        Case "Numbers"
'            ExecAndUpdateToolbarButton1 "Numbers", DECMD_ORDERLIST
'
'        Case "Bullets"
'            ExecAndUpdateToolbarButton1 "Bullets", DECMD_UNORDERLIST
'
'        Case "Outdent"
'            ExecAndUpdateToolbarButton1 "Outdent", DECMD_OUTDENT
'
'        Case "Indent"

End Sub

Private Sub mnuBullets_Click()
    ExecAndUpdateToolbarButton1 "Bullets", DECMD_UNORDERLIST
End Sub

Private Sub mnuCenter_Click()
    ExecAndUpdateToolbarButton1 "Center", DECMD_JUSTIFYCENTER
    Toolbar1.Refresh
End Sub

Private Sub mnuIndent_Click()
    ExecAndUpdateToolbarButton1 "Indent", DECMD_INDENT
End Sub

Private Sub mnuItalic_Click()
    ExecAndUpdateToolbarButton1 "Italic", DECMD_ITALIC
End Sub

Private Sub mnuJustify_Click()
    ExecAndUpdateToolbarButton11 "JustifyFull", "JustifyFull"
End Sub

Private Sub mnuLeft_Click()
    ExecAndUpdateToolbarButton1 "LeftJustify", DECMD_JUSTIFYLEFT
    Toolbar1.Refresh
End Sub

Private Sub mnuNumberedList_Click()
    ExecAndUpdateToolbarButton1 "Numbers", DECMD_ORDERLIST
End Sub

Private Sub mnuOutdent_Click()
    ExecAndUpdateToolbarButton1 "Outdent", DECMD_OUTDENT
End Sub

Private Sub mnuRight_Click()
    ExecAndUpdateToolbarButton1 "RightJustify", DECMD_JUSTIFYRIGHT
    Toolbar1.Refresh
End Sub

Private Sub mnuUnderline_Click()
    ExecAndUpdateToolbarButton1 "Underline", DECMD_UNDERLINE
End Sub

Private Sub tbFilesBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    On Error Resume Next
    Select Case Button.Key
        Case "UpDown"
            Dim x As Integer
            x = tbFilesBar.Align
            x = x + 1
            If x > 2 Then x = 1
            tbFilesBar.Align = x
            SaveSetting App.Title, "Settings", "FilesBarAlign", tbFilesBar.Align
    End Select
End Sub

Public Sub LoadNewDoc()

    On Error Resume Next
    Dim frmD As frmDocument
    Dim sFile As String
    
    lDocumentCount = lDocumentCount + 1
    NewDocumentCount = NewDocumentCount + 1
    
    Set frmD = New frmDocument
    
    frmD.Tag = lDocumentCount
    ActiveDocument = lDocumentCount
    
    sFile = "New_Page_" & CStr(NewDocumentCount) & ".html"
    
    EditMode.Add Item:=NormalMode, Key:=CStr(lDocumentCount)
    OpenFilenames.Add Item:=sFile, Key:=CStr(lDocumentCount)
    OpenDocuments.Add Item:=frmD, Key:=CStr(lDocumentCount)
    
    frmD.Caption = sFile
    frmD.Show
    
    tabFilesBar.Tabs.Add , "F" & CStr(lDocumentCount), sFile
    tabFilesBar.Tabs.Item("F" & CStr(lDocumentCount)).Selected = True
    tabFilesBar.Tabs.Item("F" & CStr(lDocumentCount)).Image = 1
    
    
End Sub

Public Sub LoadOpenDoc(sFile As String)

    On Error Resume Next
    Dim frmD As frmDocument
    
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    
    frmD.Tag = lDocumentCount
    ActiveDocument = lDocumentCount
    
    EditMode.Add Item:=NormalMode, Key:=CStr(lDocumentCount)
    OpenFilenames.Add Item:=sFile, Key:=CStr(lDocumentCount)
    OpenDocuments.Add Item:=frmD, Key:=CStr(lDocumentCount)
    
    frmD.Caption = sFile
    frmD.Show
    
    tabFilesBar.Tabs.Add , "F" & CStr(lDocumentCount), sFile
    tabFilesBar.Tabs.Item("F" & CStr(lDocumentCount)).Selected = True
    tabFilesBar.Tabs.Item("F" & CStr(lDocumentCount)).Image = 1
    
End Sub

Public Sub LoadDownloadedDoc()

    On Error Resume Next
    Dim frmD As frmDocument
    Dim sFile As String
    
    lDocumentCount = lDocumentCount + 1
    NewDocumentCount = NewDocumentCount + 1
    
    Set frmD = New frmDocument
    
    frmD.Tag = lDocumentCount
    ActiveDocument = lDocumentCount
    
    sFile = "New_Page_" & CStr(NewDocumentCount) & ".html"
    
    EditMode.Add Item:=NormalMode, Key:=CStr(lDocumentCount)
    OpenFilenames.Add Item:=sFile, Key:=CStr(lDocumentCount)
    OpenDocuments.Add Item:=frmD, Key:=CStr(lDocumentCount)
    
    frmD.Caption = sFile
    'frmD.Show
    
    tabFilesBar.Tabs.Add , "F" & CStr(lDocumentCount), sFile
    'tabFilesBar.Tabs.Item("F" & CStr(lDocumentCount)).Selected = True
    tabFilesBar.Tabs.Item("F" & CStr(lDocumentCount)).Image = 1
    
    '----------------------------------------------------------------
    ActiveForm.DHTMLEdit1.BrowseMode = True    ' Clear the Undo/Redo buffer
    ActiveForm.DHTMLEdit1.BrowseMode = False   ' Switch to edit mode
    '----------------------------------------------------------------
    
End Sub

Private Sub MDIForm_Resize()
    
    On Error Resume Next
    '----------------------------------------------------------------
    ' Resize the open files Tab strip in the toolbar
    tabFilesBar.Move tbFilesBar.Buttons("FileBar").left, 0, _
            Me.Width - tbFilesBar.Buttons("UpDown").Width - 200
    tabFilesBar.Height = tbFilesBar.Height
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.left
        SaveSetting App.Title, "Settings", "MainTop", Me.top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        SaveSetting App.Title, "Settings", "RightPanWidth", picRight.Width
        SaveSetting App.Title, "Settings", "chkExpandAll", chkExpandAll.Value
        SaveSetting App.Title, "Settings", "chkAttributes", chkAttributes.Value
    End If
    
    Dim frm As Object
    For Each frm In Forms
        Unload frm
    Next frm
    
    Dim Rtn As Long
    Rtn = ShellExecute(Me.hwnd, "Open", "http://www.mewsoft.com", "", "", 5)
    
End Sub

Private Sub cboFilesFilter_Change()
    'File1.Pattern = cboFilesFilter.Text
    On Error Resume Next
    File1.Filename = cboFilesFilter.Text
    'File1.Refresh
    Dir1_Change
    
End Sub

Private Sub cboFilesFilter_Click()

    On Error Resume Next
    File1.Pattern = cboFilesFilter.Text
    'File1.Filename = cboFilesFilter.Text
    'File1.Refresh
    'Dir1_Change

End Sub

Private Sub Dir1_Change()

    On Error Resume Next
    File1.Path = Dir1.Path
    
End Sub

Private Sub Drive1_Change()
    
    On Error Resume Next
    Dir1.Path = Drive1.Drive
    
End Sub

Private Sub File1_DblClick()

    Dim File As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    File = File1.Path & IIf(right(File1.Path, 1) = "\", "", "\") & File1.Filename
    
    If FileExists(File) Then
           LoadDocumentFile File
    End If
    
    fMainForm.cmdTreeRefresh_Click
    
    Screen.MousePointer = vbNormal
    
    ConfigDTHMLEditBehaviour
    
End Sub

Private Sub cmdRightPanRL_Click()
    
    On Error Resume Next
    m_cSplitRight.Detach
    If picRight.Align < 4 Then
        picRight.Align = picRight.Align + 1
    Else
        picRight.Align = 3
    End If
    
    picRight.Refresh
    m_cSplitRight.Attach picRight
        
    SaveSetting App.Title, "Settings", "RightPanAlign", picRight.Align
    
End Sub

Private Sub picRight_Resize()
    
    On Error Resume Next
    
    picRight.ScaleMode = vbTwips
    picRightPanTitle.ScaleMode = vbTwips
    
    picRightPanTitle.Move 0, 0, picRight.Width, picRightPanTitle.Height
    
    cmdCloseRightPan.Move picRight.Width - cmdCloseRightPan.Width - 75, 15
    cmdRightPanRL.Move picRight.Width - cmdRightPanRL.Width - cmdCloseRightPan.Width - 100, 15
     
    SSTab1.Move 50, _
        picRightPanTitle.top + picRightPanTitle.Height + 50, _
        picRight.Width - 100, _
        picRight.Height - (picRightPanTitle.top + picRightPanTitle.Height + 50) - 100
    
    '----------------------------------------------------------------
    ' Files Tab
    If SSTab1.Tab = 0 Then
        FrameFiles.Move SSTab1.left + 0, _
            SSTab1.top, _
            SSTab1.Width - 75, _
             SSTab1.Height - 440
        
        Drive1.Move FrameFiles.left + 5, _
            FrameFiles.top + 1, _
            FrameFiles.Width - 100
        
        Dir1.Move FrameFiles.left + 5, _
            Drive1.top + Drive1.Height + 20, _
            FrameFiles.Width - 100, _
            2500
        
        cboFilesFilter.Move FrameFiles.left + 5, _
            Dir1.top + Dir1.Height + 20, _
            FrameFiles.Width - 100
        
        File1.Move FrameFiles.left + 5, _
            cboFilesFilter.top + cboFilesFilter.Height + 20, _
            Abs(FrameFiles.Width - 100), _
            Abs(FrameFiles.Height - (cboFilesFilter.top + cboFilesFilter.Height + 20) - 30)
    End If
    '----------------------------------------------------------------
    'Properties Tab
    If SSTab1.Tab = 1 Then
        FrameProperties.Move SSTab1.left + 0, _
            SSTab1.top, _
            SSTab1.Width - 75, _
             SSTab1.Height - 440
        
        lbProperty.Move FrameProperties.left + 5
        txtProperty.Move FrameProperties.left + 5, _
                lbProperty.top + lbProperty.Height + 10, _
                FrameProperties.Width - 150, _
                1000
        lvProperties.Move FrameProperties.left + 5, _
                txtProperty.top + txtProperty.Height + 10, _
                FrameProperties.Width - 150, _
                FrameProperties.Height - (txtProperty.top + txtProperty.Height + 10) - 10
                
    End If
    '----------------------------------------------------------------
    'Tree Tab
    If SSTab1.Tab = 2 Then
        frameTree.Move SSTab1.left + 0, _
            SSTab1.top, _
            SSTab1.Width - 75, _
             SSTab1.Height - 440
        
        tvDomTree.Move frameTree.left + 5, _
            cmdTreeRefresh.top + cmdTreeRefresh.Height + 10, _
            frameTree.Width - 50, _
            frameTree.Height - (cmdTreeRefresh.top + cmdTreeRefresh.Height + 10) - 50
    End If
    '----------------------------------------------------------------
    picRightPanTitleGradient
    
End Sub

Private Sub picRightWin_Resize()
    
    'SSTab1.Move 0, 20, picRightWin.ScaleWidth, picRightWin.ScaleHeight
        
End Sub

Sub picRightPanTitleGradient()
    
    'Applies the gradient effect to the "Title" control
    'For veryical
    'Line (0, intY)-(500, intY) [, RGB(0, 0, CInt((intY / 500) * 255))]
    Dim intY As Integer
    On Error Resume Next
    
    picRightPanTitle.Scale (0, 0)-(1000, 1000)
    
    For intY = 0 To 1000
        picRightPanTitle.Line (intY, 0)-(intY, 1000), RGB(30, 40, CInt((intY / 1200) * 205))
    Next
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    
    On Error Resume Next
    
    If SSTab1.Tab = 0 Then
            FrameProperties.Visible = False
            frameTree.Visible = False
            FrameFiles.Visible = True
    ElseIf SSTab1.Tab = 1 Then
            FrameFiles.Visible = False
            frameTree.Visible = False
            FrameProperties.Visible = True
    ElseIf SSTab1.Tab = 2 Then
            FrameFiles.Visible = False
            FrameProperties.Visible = False
            frameTree.Visible = True
    End If
    
    picRight_Resize
    SaveSetting App.Title, "Settings", "RightPanTab", SSTab1.Tab
    
End Sub

Private Sub lvProperties_Click()

   'Debug.Print lvProperties.SelectedItem.Key & "= " & lvProperties.SelectedItem.Text
    On Error Resume Next
   
   If ActiveForm Is Nothing Then Exit Sub
   If lvProperties.ListItems.count < 1 Then Exit Sub
   
   lbProperty.Caption = lvProperties.SelectedItem.Text
   txtProperty.Text = lvProperties.SelectedItem.SubItems(1)
   
   txtProperty.Tag = lvProperties.SelectedItem.Text

End Sub

Private Sub tabFilesBar_Click()
    
    On Error Resume Next
    
    Dim TabKey As String
    TabKey = tabFilesBar.SelectedItem.Key
    TabKey = Replace(TabKey, "F", "")
    ActiveDocument = TabKey
    OpenDocuments(CStr(ActiveDocument)).SetFocus

End Sub


Private Sub txtProperty_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    
    Dim LvItem As ListItem
    
    If ActiveForm Is Nothing Then Exit Sub
    If lvProperties.ListItems.count < 1 Then Exit Sub
    
    'On Error Resume Next
    'Debug.Print "my index is : " & txtProperty.Tag & " = " & txtProperty.Text
    'Debug.Print "Nodename: " & ActiveIHTMLElement.Attributes(txtProperty.Tag).nodeName
    
    'ActiveIHTMLElement.Attributes(txtProperty.Tag).nodeValue = txtProperty.Text
    ActiveIHTMLElement.Attributes(ActiveIHTMLElement.Attributes(txtProperty.Tag).nodeName).nodeValue = txtProperty.Text
    
    Set LvItem = lvProperties.ListItems.Item((txtProperty.Tag))
    LvItem.SubItems(1) = ActiveIHTMLElement.Attributes(txtProperty.Tag).nodeValue


End Sub

Private Sub txtProperty_LostFocus()

    'Debug.Print txtProperty.Tag & " = " & txtProperty.Text
    'ActiveIHTMLElement.Attributes(txtProperty.Tag).nodeValue = txtProperty.Text
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    If lvProperties.ListItems.count < 1 Then Exit Sub
    txtProperty.Text = lvProperties.SelectedItem.SubItems(1)
 
End Sub

'====================================================================
'====================================================================
Public Function GetElementAttributes(el As Object)
    
    Dim i As Integer
    Dim Text As String
    
    On Error Resume Next
     
    Text = ""
    For i = 0 To el.Attributes.length - 1
        If left$(el.Attributes(i).nodeName, 2) <> "on" Then
            If el.Attributes(i).nodeValue <> "" Then
                Text = Text & el.Attributes(i).nodeName & "=" & el.Attributes(i).nodeValue & ", "
            End If
        End If
    Next
    Text = Mid(Text, 1, Len(Text) - 2)
    
    GetElementAttributes = Text
    Exit Function
    
ehUpdateProperties:
    GetElementAttributes = ""
    
End Function

Private Sub chkAttributes_Click()

    If ActiveForm Is Nothing Then Exit Sub
    cmdTreeRefresh_Click

End Sub

Private Sub chkExpandAll_Click()
    
    On Error GoTo ErrHandler
    Dim x As Integer
    
    If ActiveForm Is Nothing Then Exit Sub
    
    cmdTreeRefresh_Click
    
    Screen.MousePointer = vbHourglass
    
    For x = 1 To tvDomTree.Nodes.count
        tvDomTree.Nodes.Item(x).Expanded = chkExpandAll.Value
    Next x
    tvDomTree.Nodes.Item(1).EnsureVisible
    
    Screen.MousePointer = vbDefault
ErrHandler:

End Sub

Public Sub cmdTreeRefresh_Click()
    
    If ActiveForm Is Nothing Then Exit Sub
    Screen.MousePointer = vbHourglass
    UpdateDocTree
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub UpdateDocTree()
     
    On Error GoTo ErrHandler
    
    Dim doc As HTMLDocument
    Dim el As IHTMLElement
    Dim x As Integer
    
    If ActiveForm Is Nothing Then Exit Sub
    While ActiveForm.DHTMLEdit1.Busy = True
    DoEvents
    Wend
    Set doc = ActiveForm.DHTMLEdit1.DOM
    
    Set el = doc.documentElement
    tvDomTree.Nodes.Clear
    
    TreeNodeID = 1
    
    'AddElement el, TreeNodeID
    AddElement el
    
    'Programatically show the first expansion. Note there is no 0 node.
    tvDomTree.Nodes.Item(1).Expanded = True
    
    For x = 1 To tvDomTree.Nodes.count
        tvDomTree.Nodes.Item(x).Expanded = chkExpandAll.Value
    Next x
    
    tvDomTree.Nodes.Item(1).EnsureVisible
    
ErrHandler:

End Sub

Public Sub AddElement(ByVal el As IHTMLElement, Optional lRelative As Variant)
    
    Dim el2 As IHTMLElement2
    Dim Attr As String
    Dim lKey As Double
    
    lKey = el.sourceIndex
    'Debug.Print "sourceIndex: " & el.sourceIndex
    
    Attr = ""
    If chkAttributes.Value = vbChecked Then
        Attr = " (" & GetElementAttributes(el) & ")"
    End If
        
    If IsMissing(lRelative) Then
        'Debug.Print "key:" & lKey & " lRelative: "
        tvDomTree.Nodes.Add , , "E" & CStr(lKey), el.TagName & Attr
        
    Else
        'Debug.Print "key:" & lKey & " lRelative: " & lRelative
        tvDomTree.Nodes.Add "E" & CStr(lRelative), tvwChild, "E" & CStr(lKey), el.TagName & Attr
    End If
    If Not el.children Is Nothing Then AddChildren el, lKey
    
End Sub

Public Sub AddChildren(ByVal el As IHTMLElement, ByVal lParentKey As Double)
    
    Dim i As Integer
    
    For i = 0 To el.children.length - 1
        TreeNodeID = TreeNodeID + 1
        'AddElement el.children.Item(i), TreeNodeID, lParentKey
        AddElement el.children.Item(i), lParentKey
    Next
    
End Sub

'====================================================================
'====================================================================
Private Sub cmdCloseRightPan_Click()
    
    picRight.Visible = False
    mnuToolsWindow.Checked = False
    SaveSetting App.Title, "Settings", "RightPan", picRight.Visible
    
End Sub

Private Sub mnuToolsWindow_Click()

    mnuToolsWindow.Checked = Not mnuToolsWindow.Checked
    
    If mnuToolsWindow.Checked Then
        picRight.Visible = True
    Else
        picRight.Visible = False
    End If

    SaveSetting App.Title, "Settings", "RightPan", picRight.Visible
    
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowMinimizeAll_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then Exit Sub
    Dim i As Integer
    ' Minimize all documents
    For i = 1 To Forms.count - 1
        Forms(i).WindowState = vbMinimized
    Next i
End Sub
Private Sub mnuWindowRestoreAll_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then Exit Sub
    Dim i As Integer
    ' Restore all documents
    For i = 1 To Forms.count - 1
        Forms(i).WindowState = vbNormal
    Next i
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Private Sub mnuToolsOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewWebBrowser_Click()
    Dim frmB As New frmBrowser
    frmB.StartingAddress = "http://www.mewsoft.com"
    frmB.Show
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    'mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    'tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtfText.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF
    ActiveForm.rtfText.SelText = vbNullString

End Sub

Private Sub mnuEditUndo_Click()
    'ToDo: Add 'mnuEditUndo_Click' code.
    MsgBox "Add 'mnuEditUndo_Click' code."
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    'ToDo: Add 'mnuFileSend_Click' code.
    MsgBox "Add 'mnuFileSend_Click' code."
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    
'    With dlgCommonDialog
'        .DialogTitle = "Print"
'        .CancelError = True
'        .FLAGS = cdlPDReturnDC + cdlPDNoPageNums
'        If ActiveForm.rtfText.SelLength = 0 Then
'            .FLAGS = .FLAGS + cdlPDAllPages
'        Else
'            .FLAGS = .FLAGS + cdlPDSelection
'        End If
'        .ShowPrinter
'        If Err <> MSComDlg.cdlCancel Then
'            ActiveForm.rtfText.SelPrint .hdc
'        End If
'    End With

End Sub

Private Sub mnuFilePrintPreview_Click()
    'ToDo: Add 'mnuFilePrintPreview_Click' code.
    MsgBox "Add 'mnuFilePrintPreview_Click' code."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
'    With dlgCommonDialog
'        .DialogTitle = "Page Setup"
'        .CancelError = True
'        .ShowPrinter
'    End With

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub

Private Sub mnuFileSaveAll_Click()
    'ToDo: Add 'mnuFileSaveAll_Click' code.
    MsgBox "Add 'mnuFileSaveAll_Click' code."
End Sub

Private Sub mnuFileSaveAs_Click()
    
    Dim sFile As String
    
    If ActiveForm Is Nothing Then Exit Sub
    

'    With dlgCommonDialog
'        .DialogTitle = "Save As"
'        .CancelError = False
'        'ToDo: set the flags and attributes of the common dialog control
'        .Filter = "All Files (*.*)|*.*"
'        .ShowSave
'        If Len(.Filename) = 0 Then
'            Exit Sub
'        End If
'        sFile = .Filename
'    End With
'    ActiveForm.Caption = sFile
'    ActiveForm.rtfText.SaveFile sFile

End Sub

Private Sub mnuFileSave_Click()
    Dim sFile As String
    
'    If left$(ActiveForm.Caption, 8) = "Document" Then
'        With dlgCommonDialog
'            .DialogTitle = "Save"
'            .CancelError = False
'            'ToDo: set the flags and attributes of the common dialog control
'            .Filter = "All Files (*.*)|*.*"
'            .ShowSave
'            If Len(.Filename) = 0 Then
'                Exit Sub
'            End If
'            sFile = .Filename
'        End With
'        ActiveForm.rtfText.SaveFile sFile
'    Else
'        sFile = ActiveForm.Caption
'        ActiveForm.rtfText.SaveFile sFile
'    End If

End Sub

Private Sub mnuFileClose_Click()
    'ToDo: Add 'mnuFileClose_Click' code.
    MsgBox "Add 'mnuFileClose_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


'    If ActiveForm Is Nothing Then LoadNewDoc
'
'    With dlgCommonDialog
'        .DialogTitle = "Open"
'        .CancelError = False
'        'ToDo: set the flags and attributes of the common dialog control
'        .Filter = "All Files (*.*)|*.*"
'        .ShowOpen
'        If Len(.Filename) = 0 Then
'            Exit Sub
'        End If
'        sFile = .Filename
'    End With
'    ActiveForm.rtfText.LoadFile sFile
'    ActiveForm.Caption = sFile

End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub

Public Sub UpdateToolbarButton1(Button As String, Cmd As DHTMLEDITCMDID)
          
    On Error GoTo ErrHandler
    
    If ActiveForm Is Nothing Then Exit Sub
    
    Dim State As DHTMLEDITCMDF
    
    State = ActiveForm.DHTMLEdit1.QueryStatus(Cmd)
    
    If (State >= DECMDF_ENABLED) Then
        Toolbar1.Buttons(Button).Enabled = True
    Else
        Toolbar1.Buttons(Button).Enabled = False
    End If
        
    If (State = DECMDF_LATCHED) Then
        Toolbar1.Buttons(Button).Value = tbrPressed
    Else
        Toolbar1.Buttons(Button).Value = tbrUnpressed
    End If
    Exit Sub
ErrHandler:
    'MsgBox "Error: " & Err.Number & ", " & Err.Description
End Sub

Public Sub UpdateToolbarButton2(Button As String, Cmd As DHTMLEDITCMDID)
    
    On Error Resume Next
    
    If ActiveForm Is Nothing Then Exit Sub
    
    Dim State As DHTMLEDITCMDF
    
    State = ActiveForm.DHTMLEdit1.QueryStatus(Cmd)
    
    If (State >= DECMDF_ENABLED) Then
        Toolbar2.Buttons(Button).Enabled = True
    Else
        Toolbar2.Buttons(Button).Enabled = False
    End If
        
    If (State = DECMDF_LATCHED) Then
        'Toolbar2.Buttons(Button).Value = tbrPressed
    Else
        'Toolbar2.Buttons(Button).Value = tbrUnpressed
    End If
    
End Sub

Public Sub UpdateToolbarButton3(Button As String, Cmd As DHTMLEDITCMDID)
    
    On Error Resume Next
    
    If ActiveForm Is Nothing Then Exit Sub
    
    Dim State As DHTMLEDITCMDF
    
    State = ActiveForm.DHTMLEdit1.QueryStatus(Cmd)
    
    If (State >= DECMDF_ENABLED) Then
        Toolbar3.Buttons(Button).Enabled = True
    Else
        Toolbar3.Buttons(Button).Enabled = False
    End If
        
    If (State = DECMDF_LATCHED) Then
        'Toolbar2.Buttons(Button).Value = tbrPressed
    Else
        'Toolbar2.Buttons(Button).Value = tbrUnpressed
    End If
    
End Sub

Private Sub ExecAndUpdateToolbarButton1(Button As String, Command As DHTMLEDITCMDID)
    
    On Error Resume Next
    
    Dim State As DHTMLEDITCMDF

    If Not Command = 0 Then
        ActiveForm.DHTMLEdit1.execCommand Command, OLECMDEXECOPT_DONTPROMPTUSER
        
        State = ActiveForm.DHTMLEdit1.QueryStatus(Command)
        
        If (State >= DECMDF_ENABLED) Then
            Toolbar1.Buttons(Button).Value = tbrPressed
        Else
            Toolbar1.Buttons(Button).Value = tbrUnpressed
        End If
    End If

End Sub

Public Sub ExecAndUpdateToolbarButton11(Button As String, Command As String)
     
    On Error Resume Next
    
    Dim State As Boolean
    
    If Command = "" Or Button = "" Then Exit Sub
    
    'execCommand, queryCommandEnabled, queryCommandIndeterm,
    'queryCommandState, queryCommandSupported, queryCommandValue.
    State = ActiveForm.DHTMLEdit1.DOM.execCommand(Command)
    
    State = ActiveForm.DHTMLEdit1.DOM.queryCommandEnabled(Command)
    If (State) Then
        Toolbar1.Buttons(Button).Enabled = True
    Else
        Toolbar1.Buttons(Button).Enabled = False
    End If
    
    State = ActiveForm.DHTMLEdit1.DOM.queryCommandValue(Command)
    If (State) Then
        Toolbar1.Buttons(Button).Value = tbrPressed
    Else
        Toolbar1.Buttons(Button).Value = tbrUnpressed
    End If
    
    'fMainForm.UpdateToolbarButton11 "JustifyFull", "JustifyFull"
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

End Sub

Public Sub UpdateToolbarButton11(Button As String, Command As String)
     
    On Error Resume Next
    
    Dim State As Boolean
    
    If Command = "" Or Button = "" Then Exit Sub
    
    State = ActiveForm.DHTMLEdit1.DOM.queryCommandEnabled(Command)
    
    If (State) Then
        Toolbar1.Buttons(Button).Enabled = True
    Else
        Toolbar1.Buttons(Button).Enabled = False
    End If
    
    State = ActiveForm.DHTMLEdit1.DOM.queryCommandValue(Command)
    
    If (State) Then
        Toolbar1.Buttons(Button).Value = tbrPressed
    Else
        Toolbar1.Buttons(Button).Value = tbrUnpressed
    End If
    
    
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    On Error Resume Next
    
    Dim State As DHTMLEDITCMDF
    
    ' Handle toolbar commands
    Select Case Button.Key
        Case "Bold"
            ExecAndUpdateToolbarButton1 "Bold", DECMD_BOLD
            
        Case "Italic"
            ExecAndUpdateToolbarButton1 "Italic", DECMD_ITALIC
            
        Case "Underline"
            ExecAndUpdateToolbarButton1 "Underline", DECMD_UNDERLINE
            
        Case "Numbers"
            ExecAndUpdateToolbarButton1 "Numbers", DECMD_ORDERLIST
            
        Case "Bullets"
            ExecAndUpdateToolbarButton1 "Bullets", DECMD_UNORDERLIST
            
        Case "Outdent"
            ExecAndUpdateToolbarButton1 "Outdent", DECMD_OUTDENT
            
        Case "Indent"
            ExecAndUpdateToolbarButton1 "Indent", DECMD_INDENT
            
        Case "LeftJustify"
            ExecAndUpdateToolbarButton1 "LeftJustify", DECMD_JUSTIFYLEFT
            'Toolbar1.Buttons("Center").Value = tbrUnpressed
            'Toolbar1.Buttons("RightJustify").Value = tbrUnpressed
            Toolbar1.Refresh
            
        Case "Center"
            ExecAndUpdateToolbarButton1 "Center", DECMD_JUSTIFYCENTER
            'Toolbar1.Buttons("LeftJustify").Value = tbrUnpressed
            'Toolbar1.Buttons("RightJustify").Value = tbrUnpressed
            Toolbar1.Refresh
            
        Case "RightJustify"
            ExecAndUpdateToolbarButton1 "RightJustify", DECMD_JUSTIFYRIGHT
            'Toolbar1.Buttons("Center").Value = tbrUnpressed
            'Toolbar1.Buttons("LeftJustify").Value = tbrUnpressed
            Toolbar1.Refresh
            
        Case "JustifyFull"
            ExecAndUpdateToolbarButton11 "JustifyFull", "JustifyFull"
            'ActiveForm.DHTMLEdit1.DOM.body.createTextRange.collapse False
                    
        Case "SuperScript"
            ExecAndUpdateToolbarButton11 "SuperScript", "Superscript"
                    
        Case "SubScript"
            ExecAndUpdateToolbarButton11 "SubScript", "Subscript"
            'ActiveForm.DHTMLEdit1.DOM.body.createTextRange.collapse False
                    
        Case "StrikeThrough"
            ExecAndUpdateToolbarButton11 "StrikeThrough", "StrikeThrough"
            'ActiveForm.DHTMLEdit1.DOM.body.createTextRange.collapse False
        
        ' Fore color
        Case "Color"
            Dim ForeColor As String
            On Error GoTo cleanup
            CommonDialog1.color = 0
            CommonDialog1.CancelError = True
            CommonDialog1.ShowColor
            ForeColor = ""
            ForeColor = FormatRGBString(CommonDialog1.color)
            ActiveForm.DHTMLEdit1.execCommand DECMD_SETFORECOLOR, OLECMDEXECOPT_DONTPROMPTUSER, ForeColor
            
        ' Back color
        Case "BackColor"
            Dim BackColor As String
            On Error GoTo cleanup
            CommonDialog1.color = 0
            CommonDialog1.CancelError = True
            CommonDialog1.ShowColor
            BackColor = ""
            BackColor = FormatRGBString(CommonDialog1.color)
            ActiveForm.DHTMLEdit1.execCommand DECMD_SETBACKCOLOR, OLECMDEXECOPT_DONTPROMPTUSER, BackColor
        End Select
        
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
        
cleanup:
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    On Error Resume Next
    
    Dim State As Boolean
    ' Handle toolbar commands
    Select Case Button.Key
        Case "New"
            FileNew_Click
            fMainForm.cmdTreeRefresh_Click
            ConfigDTHMLEditBehaviour
            
        Case "Open"
            FileOpen_Click
            fMainForm.cmdTreeRefresh_Click
            ConfigDTHMLEditBehaviour
            
        Case "Save"
            FileSave_Click
        
        Case "SaveAs"
            FileSaveAs_Click
            
        Case "Download"
            If OpenDocuments.count < 1 Then
                'fMainForm.LoadNewDoc
            End If
            CenterForm frmOpenURL, fMainForm
            frmOpenURL.Show vbModal, fMainForm
            ConfigDTHMLEditBehaviour
            
        Case "SpellCheck"
            ActiveForm.CheckSpelling
            
        Case "Undo"
            ActiveForm.DHTMLEdit1.execCommand DECMD_UNDO, OLECMDEXECOPT_DODEFAULT
        
        Case "Redo"
            ActiveForm.DHTMLEdit1.execCommand DECMD_REDO, OLECMDEXECOPT_DODEFAULT
        
        Case "Cut"
            ActiveForm.DHTMLEdit1.execCommand DECMD_CUT, OLECMDEXECOPT_DODEFAULT
        
        Case "Copy"
            ActiveForm.DHTMLEdit1.execCommand DECMD_COPY, OLECMDEXECOPT_DODEFAULT
        
        Case "Paste"
            ActiveForm.DHTMLEdit1.execCommand DECMD_PASTE, OLECMDEXECOPT_DODEFAULT
        
        Case "Find"
            ActiveForm.DHTMLEdit1.execCommand DECMD_FINDTEXT, OLECMDEXECOPT_DODEFAULT
        
        ' Insert Table
        Case "Table"
            InsertTable
        
        Case "Image"
            ActiveForm.DHTMLEdit1.execCommand DECMD_IMAGE, OLECMDEXECOPT_DODEFAULT
        
        Case "Hyperlink"
            ActiveForm.DHTMLEdit1.execCommand DECMD_HYPERLINK, OLECMDEXECOPT_DODEFAULT
    
        Case "ShowAll"
            SwitchShowAll
            
        Case "ShowBorders"
            SwitchShowBorders
            
        Case "Print"
            ActiveForm.DHTMLEdit1.PrintDocument
            
        Case "Form"
            InsertForm
        
        Case "Properties"
            ActiveForm.DHTMLEdit1.execCommand DECMD_PROPERTIES, OLECMDEXECOPT_DODEFAULT
        
            
    End Select
    
    'object.SourceCodePreservation [ = enablePreservation ]
    'object.FilterSourceCode sourceCodeIn
    'object.UseDivOnCarriageReturn [ = div ]
    

End Sub

Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

    On Error Resume Next
    
    Select Case ButtonMenu.Key
        Case "Form"
            InsertForm
            
        Case "Textbox"
            InsertTextbox
    
        Case "Textarea"
            InsertTextarea
       
        Case "Checkbox"
            InsertCheckbox
            
        Case "OptionButton"
            InsertOptionButton
        
        Case "ListBox"
            InsertListBox
            
        Case "DropDownBox"
            InsertDropDownBox

        Case "PushButton"
            InsertPushButton
            
        Case "HiddenData"
            InsertHiddenData
            
        Case "Password"
            InsertPassword
            
        Case "SubmitButton"
            InsertSubmitButton
            
        Case "ResetButton"
            InsertResetButton
       
        Case "ImageButton"
            InsertImageButton
            
        Case "FileUpload"
            InsertFileUpload
            
    End Select
    
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next
    
    Dim State As DHTMLEDITCMDF
        
    ' Handle toolbar commands
    Select Case Button.Key
        Case "Normal"
            EditMode.Remove CStr(ActiveDocument)
            EditMode.Add Key:=CStr(ActiveDocument), Item:=NormalMode
            SwitchToNormal
            
        Case "HTML"
            EditMode.Remove CStr(ActiveDocument)
            EditMode.Add Key:=CStr(ActiveDocument), Item:=HtmlMode
            SwitchToHTML
            
        Case "Preview"
            EditMode.Remove CStr(ActiveDocument)
            EditMode.Add Key:=CStr(ActiveDocument), Item:=PreviewMode
            SwitchToPreview
    
        Case "Refresh"
            If EditMode(CStr(ActiveDocument)) <> HtmlMode Then
                ActiveForm.DHTMLEdit1.Refresh
            End If
            
        Case "Stop"
    
        Case "InsertRows"
            CommandExec DECMD_INSERTROW
    
        Case "InsertColumns"
            CommandExec DECMD_INSERTCOL
        
        Case "InsertCells"
            CommandExec DECMD_INSERTCELL
        
        Case "DeleteCells"
            CommandExec DECMD_DELETECELLS
        
        Case "DeleteColumns"
            CommandExec DECMD_DELETECOLS
        
        Case "DeleteRows"
            CommandExec DECMD_DELETEROWS
        
        Case "MergeCells"
            CommandExec DECMD_MERGECELLS
        
        Case "SplitCells"
            CommandExec DECMD_SPLITCELL
        '------------------------------------------------------------
        'PositionAbsolutely
        Case "PositionAbsolutely"
            CommandExec DECMD_MAKE_ABSOLUTE
        
        'BringForward
        Case "BringForward"
            CommandExec DECMD_BRING_FORWARD
        
        'SendBackward
        Case "SendBackward"
            CommandExec DECMD_SEND_BACKWARD
        
        'BringToFront
        Case "BringToFront"
            CommandExec DECMD_BRING_TO_FRONT
        
        'SendToBack
        Case "SendToBack"
            CommandExec DECMD_SEND_TO_BACK
        
        'Bring Above Text
        Case "BringAboveText"
            CommandExec DECMD_BRING_ABOVE_TEXT
        
        'Send Below Text
        Case "SendBelowText"
            CommandExec DECMD_SEND_BELOW_TEXT
        
        Case "SnapToGrid"
            DhtmlSnapToGrid = Not DhtmlSnapToGrid
            ActiveForm.DHTMLEdit1.SnapToGrid = DhtmlSnapToGrid
            If DhtmlSnapToGrid Then
                   Toolbar3.Buttons("SnapToGrid").Value = tbrPressed
            Else
                   Toolbar3.Buttons("SnapToGrid").Value = tbrUnpressed
            End If
             SaveSetting App.Title, "Settings", "DhtmlSnapToGrid", DhtmlSnapToGrid
        
        'Textbox
        Case "Textbox"
            InsertTextbox
    
        Case "Textarea"
            InsertTextarea
       
        Case "Checkbox"
            InsertCheckbox
            
        Case "OptionButton"
            InsertOptionButton
        
        Case "ListBox"
            InsertListBox
            
        Case "DropDownBox"
            InsertDropDownBox

        Case "PushButton"
            InsertPushButton
            
        Case "HiddenData"
            InsertHiddenData
            
        Case "Password"
            InsertPassword
            
        Case "SubmitButton"
            InsertSubmitButton
            
        Case "ResetButton"
            InsertResetButton
       
        Case "ImageButton"
            InsertImageButton
        Case "FileUpload"
            InsertFileUpload
            
    End Select

    ' Initialize 2D menu command table
'    twoDMenuCmds(0) = DECMD_MAKE_ABSOLUTE
'    twoDMenuCmds(1) = DECMD_BRING_TO_FRONT
'    twoDMenuCmds(2) = DECMD_SEND_TO_BACK
'    twoDMenuCmds(3) = DECMD_BRING_FORWARD
'    twoDMenuCmds(4) = DECMD_SEND_BACKWARD

'    twoDMenuCmds(5) = DECMD_BRING_ABOVE_TEXT
'    twoDMenuCmds(6) = DECMD_SEND_BELOW_TEXT
'    twoDMenuCmds(7) = 0
'    twoDMenuCmds(8) = DECMD_LOCK_ELEMENT

End Sub

Private Sub Toolbar3_ButtonDropDown(ByVal Button As MSComctlLib.Button)

    On Error Resume Next
    
    Select Case Button.Key
        
        'PositionAbsolutely
        Case "PositionAbsolutely"
            UpdatePositionAbsolutelyMenu
        
    End Select
    
End Sub

Private Sub Toolbar3_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

    On Error Resume Next
    
    Select Case ButtonMenu.Key
        
        'PositionAbsolutely
        Case "PositionAbsolutely"
            CommandExec DECMD_MAKE_ABSOLUTE
        
        'BringForward
        Case "BringForward"
            CommandExec DECMD_BRING_FORWARD
        
        'SendBackward
        Case "SendBackward"
            CommandExec DECMD_SEND_BACKWARD
        
        'BringToFront
        Case "BringToFront"
            CommandExec DECMD_BRING_TO_FRONT
        
        'SendToBack
        Case "SendToBack"
            CommandExec DECMD_SEND_TO_BACK
        
        'Bring Above Text
        Case "BringAboveText"
            CommandExec DECMD_BRING_ABOVE_TEXT
        
        'Send Below Text
        Case "SendBelowText"
            CommandExec DECMD_SEND_BELOW_TEXT
            
        'Lock Element
        Case "LockElement"
            CommandExec DECMD_LOCK_ELEMENT
    
    End Select
    
End Sub

Private Sub CommandExec(Command As DHTMLEDITCMDID)

    On Error Resume Next
    
    Dim State As DHTMLEDITCMDF
    
    State = ActiveForm.DHTMLEdit1.QueryStatus(Command)
    
    If State >= DECMDF_ENABLED Then
        ActiveForm.DHTMLEdit1.execCommand Command, OLECMDEXECOPT_DODEFAULT
    End If

End Sub

Private Function CommandStatus(Command As DHTMLEDITCMDID)

    On Error Resume Next
    
    Dim State As DHTMLEDITCMDF
    
    State = ActiveForm.DHTMLEdit1.QueryStatus(Command)
    
    If State >= DECMDF_ENABLED Then
        CommandStatus = True
    Else
        CommandStatus = False
    End If

End Function

Private Sub UpdatePositionAbsolutelyMenu()
    
'tbToolBar.Buttons.Item("Edit").ButtonMenus.Item("Cut").Enabled = True

    'PositionAbsolutely
    Toolbar3.Buttons.Item("PositionAbsolutely").ButtonMenus.Item("PositionAbsolutely").Enabled = CommandStatus(DECMD_MAKE_ABSOLUTE)
    
    'BringForward
    Toolbar3.Buttons.Item("PositionAbsolutely").ButtonMenus.Item("BringForward").Enabled = CommandStatus(DECMD_BRING_FORWARD)
    
    'SendBackward
    Toolbar3.Buttons.Item("PositionAbsolutely").ButtonMenus.Item("SendBackward").Enabled = CommandStatus(DECMD_SEND_BACKWARD)
    
    'BringToFront
    Toolbar3.Buttons.Item("PositionAbsolutely").ButtonMenus.Item("BringToFront").Enabled = CommandStatus(DECMD_BRING_TO_FRONT)
    
    'SendToBack
    Toolbar3.Buttons.Item("PositionAbsolutely").ButtonMenus.Item("SendToBack").Enabled = CommandStatus(DECMD_SEND_TO_BACK)
    
    'Bring Above Text
    Toolbar3.Buttons.Item("PositionAbsolutely").ButtonMenus.Item("BringAboveText").Enabled = CommandStatus(DECMD_BRING_ABOVE_TEXT)
    
    'Send Below Text
    Toolbar3.Buttons.Item("PositionAbsolutely").ButtonMenus.Item("SendBelowText").Enabled = CommandStatus(DECMD_SEND_BELOW_TEXT)
        
    'Lock Element
    Toolbar3.Buttons.Item("PositionAbsolutely").ButtonMenus.Item("LockElement").Enabled = CommandStatus(DECMD_LOCK_ELEMENT)
    
    
End Sub

Private Sub SwitchToNormal()

    ' trun on the design mode
    ActiveForm.DHTMLEdit1.BrowseMode = False
    
    'get the text from the source editor to the visual editor
    ActiveForm.DHTMLEdit1.DocumentHTML = ActiveForm.Editawy1.Text
    While ActiveForm.DHTMLEdit1.Busy
        DoEvents
    Wend
    
    
    'hide the source code editor
    ActiveForm.Editawy1.Visible = False
    
    ' display the visual html editor
    ActiveForm.DHTMLEdit1.Visible = True
    ActiveForm.DHTMLEdit1.SetFocus
    
    Toolbar3.Buttons("Normal").Value = tbrPressed
    Toolbar3.Buttons("HTML").Value = tbrUnpressed
    Toolbar3.Buttons("Preview").Value = tbrUnpressed
    '----------------------------------------------------------------
    ' Moving the cursor
    Dim rg As IHTMLTxtRange
    Dim ctlRg As IHTMLControlRange
    'Set rg = ActiveForm.DHTMLEdit1.DOM.selection.createRange
    Set rg = ActiveForm.DHTMLEdit1.DOM.body.createTextRange
    rg.collapse True
    'rg.findText
    rg.moveStart "character", ActiveForm.Editawy1.CurPosition
    rg.moveEnd "sentence", 1
    rg.Select
    'IOleCommandTarget.Exce
    
'    fMainForm.sbStatusBar.Panels(2).Text = "Line " & Editawy1.CurLine & " "
'    fMainForm.sbStatusBar.Panels(3).Text = "Column " & Editawy1.Column + 1 & " "
'    fMainForm.sbStatusBar.Panels(4).Text = " Char " & Editawy1.CurPosition & " "
    '----------------------------------------------------------------

End Sub

Private Sub SwitchToHTML()

    On Error GoTo ErrHandler
    Dim cart As Long
    
    Dim e As IHTMLEventObj
    Set e = ActiveForm.DHTMLEdit1.DOM.parentWindow.event
    
    cart = GetDocCartLocation
     
    ' Copy the source code from the DHTML to the source editor
    '.FilterSourceCode
    ActiveForm.Editawy1.Text = ActiveForm.DHTMLEdit1.DocumentHTML
    While ActiveForm.DHTMLEdit1.Busy
        DoEvents
    Wend
    
    ' hide the visual html editor
    ActiveForm.DHTMLEdit1.Visible = False
    
    ' display the source code editor
    ActiveForm.Editawy1.Visible = True
    'ActiveForm.Editawy1.SetFocus
    'ActiveForm.Editawy1.KillFocus
    'ActiveForm.Editawy1.CurLine = 5
    
    ActiveForm.Editawy1.CurPosition = cart
    
    'Debug.Print "Cart: " & Cart
        
    
'    Dim rg As IHTMLTxtRange
'    Dim ctlRg As IHTMLControlRange
'    Dim Text As String
    'Set rg = ActiveForm.DHTMLEdit1.DOM.selection.createRange
    'rg.collapse
    'rg.moveToPoint x, y
    'Text = rg.parentElement.outerHTML
    
    'Text = rg.parentElement.innerHTML
    'Text = rg.parentElement.tagName
    
    Toolbar3.Buttons("Normal").Value = tbrUnpressed
    Toolbar3.Buttons("HTML").Value = tbrPressed
    Toolbar3.Buttons("Preview").Value = tbrUnpressed
    
ErrHandler:

End Sub

Private Sub SwitchToPreview()
    
    On Error Resume Next
    ' display the source code editor
    If EditMode(CStr(ActiveDocument)) = HtmlMode Then
        ActiveForm.DHTMLEdit1.DocumentHTML = ActiveForm.Editawy1.Text
        While ActiveForm.DHTMLEdit1.Busy
            DoEvents
        Wend
    End If
     
    ActiveForm.Editawy1.Text = ActiveForm.DHTMLEdit1.DocumentHTML
    
    ActiveForm.Editawy1.Visible = False
    ActiveForm.DHTMLEdit1.BrowseMode = True
    ActiveForm.DHTMLEdit1.Visible = True
    ActiveForm.DHTMLEdit1.SetFocus

End Sub

Public Sub ConfigDTHMLEditBehaviour()

    On Error Resume Next
    'Causes the MSHTML Editor to update an element's appearance
    'continuously during a resizing or moving operation, rather
    'than updating only at the completion of the move or resize.
    ActiveForm.DHTMLEdit1.DOM.execCommand "LiveResize", False, True
    '----------------------------------------------------------------
    'Allows for the selection of more than one site selectable
    'element at a time when the user holds down the SHIFT or CTRL keys.
    ActiveForm.DHTMLEdit1.DOM.execCommand "MultipleSelection", False, True
    '----------------------------------------------------------------

End Sub
Public Sub SwitchShowAll()
    
    On Error Resume Next
    DhtmlShowAll = Not DhtmlShowAll
    ActiveForm.DHTMLEdit1.ShowDetails = DhtmlShowAll
    If DhtmlShowAll Then
        Toolbar2.Buttons("ShowAll").Value = tbrPressed
    Else
        Toolbar2.Buttons("ShowAll").Value = tbrUnpressed
    End If
    
    SaveSetting App.Title, "Settings", "DhtmlShowAll", DhtmlShowAll
    
End Sub
            
Public Sub SwitchShowBorders()

    On Error Resume Next
    DhtmlShowBorders = Not DhtmlShowBorders
    ActiveForm.DHTMLEdit1.ShowBorders = DhtmlShowBorders
    If DhtmlShowBorders Then
        Toolbar2.Buttons("ShowBorders").Value = tbrPressed
    Else
        Toolbar2.Buttons("ShowBorders").Value = tbrUnpressed
    End If
    SaveSetting App.Title, "Settings", "DhtmlShowBorders", DhtmlShowBorders

End Sub

Private Sub TimerMemoryStatus_Timer()

    On Error Resume Next
    'Dim MS As MEMORYSTATUS
    MemoryInfo.dwLength = Len(MemoryInfo)
    GlobalMemoryStatus MemoryInfo
    sbStatusBar.Panels(5).Text = Format$(MemoryInfo.dwAvailPhys / 1048576, "###,###") & " MB "

End Sub

Private Sub AboutVBEdit_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub D2DSub_Click(index As Integer)

    Dim Cmd As DHTMLEDITCMDID
    Dim State As DHTMLEDITCMDF
    
    Cmd = twoDMenuCmds(index)
           
    If Not Cmd = 0 Then
        ActiveForm.DHTMLEdit1.execCommand Cmd, OLECMDEXECOPT_DODEFAULT
    End If

    State = ActiveForm.DHTMLEdit1.QueryStatus(DECMD_MAKE_ABSOLUTE)
    
    If State = DECMDF_LATCHED Then
        'D2DSub(0).Caption = "Set Position Attribute To 1D"
        'D2DSub(0).Enabled = True
    ElseIf State = DECMDF_ENABLED Then
        'D2DSub(0).Caption = "Set Position Attribute To Absolute"
        'D2DSub(0).Enabled = True
    Else
        'D2DSub(0).Caption = "Set Position Attribute To Absolute"
        'D2DSub(0).Enabled = False
    End If

    
End Sub

Private Sub EditSub_Click(index As Integer)
    Dim Cmd As DHTMLEDITCMDID
    Dim State As Boolean
    
    
    If index = 10 Then
        State = ActiveForm.DHTMLEdit1.SnapToGrid
        State = Not State
        ActiveForm.DHTMLEdit1.SnapToGrid = State
        'EditSub(Index).Checked = state
    Else
        Cmd = editMenuCmds(index)
           
        If Not Cmd = 0 Then
            ActiveForm.DHTMLEdit1.execCommand Cmd, OLECMDEXECOPT_DODEFAULT
        End If
        
    End If
        
End Sub

Private Sub FileExit_Click()
    Unload Me
End Sub

Private Sub FileNew_Click()

    If Not SaveChanges = vbCancel Then
        LoadNewDoc
        SetFormCaption
        While ActiveForm.DHTMLEdit1.Busy
            DoEvents
        Wend
        
    End If
    
End Sub

Private Sub FileOpen_Click()

'    On Error Resume Next
    On Error GoTo ErrHandler
    
    With CommonDialog1
       .DialogTitle = "Open File"
       .CancelError = True
       .Filter = "Web Pages (*.htm;*.html;*.shtml;*.shtm;*.stm;*.asp;*.aspx;*.css)|*.htm;*.html;*.shtml;*.shtm;*.stm;*.asp;*.aspx;*.css|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
       .ShowOpen
        
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        
        If FileExists(.Filename) Then
               LoadDocumentFile .Filename
        End If
    End With
        
    Exit Sub
    
ErrHandler:
    
End Sub

Public Sub LoadDocumentFile(File As String)

    On Error Resume Next
    
    Dim fs As FileSystemObject
    Dim ts As TextStream
    Dim Text As String
    
    Set fs = New FileSystemObject
    Set ts = fs.OpenTextFile(File)
    Text = ts.ReadAll
    
    DisableToolbar
    
    LoadOpenDoc File
    
    ActiveForm.DHTMLEdit1.DocumentHTML = Text
    'ActiveForm.DHTMLEdit1.BaseURL = .Filename
    
    While ActiveForm.DHTMLEdit1.Busy
        DoEvents
    Wend
    
    'ActiveForm.DHTMLEdit1.DOM.selection.createTextRange.collapse
    'OpenFilenames(CStr(ActiveDocument)).SetFocus
    'OpenFilenames(CStr(ActiveDocument)).DHTMLEdit1.DOM.selection.createTextRange.collapse
    
    Set ts = Nothing
    Set fs = Nothing
    '----------------------------------------------------------------
    ActiveForm.DHTMLEdit1.BrowseMode = True    ' Clear the Undo/Redo buffer
    ActiveForm.DHTMLEdit1.BrowseMode = False   ' Switch to edit mode
    '----------------------------------------------------------------

End Sub
Private Sub File_Click()
    If Len(ActiveForm.DHTMLEdit1.CurrentDocumentPath) > 0 Then
        'FileSave.Enabled = True
    Else
        'FileSave.Enabled = False
    End If
    
End Sub

Private Sub FileSave_Click()
    SaveDocument False
End Sub

Private Sub FileSaveAs_Click()
    SaveDocument True
End Sub

Private Sub StyleCombo_Click()

    On Error Resume Next
    
    Dim State As DHTMLEDITCMDF
    Dim Format As String
    
    State = ActiveForm.DHTMLEdit1.QueryStatus(DECMD_SETBLOCKFMT)
    
    If State >= DECMDF_ENABLED Then
        ActiveForm.DHTMLEdit1.execCommand DECMD_SETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER, StyleCombo.Text
    End If
    
End Sub

Private Sub FontCombo_Click()
    
    On Error Resume Next
    
    Dim fn As String
    Dim State As DHTMLEDITCMDF
    
    fn = fontNames(FontCombo.ListIndex)
    
    If (DHTMLEditInitialized) Then
        State = ActiveForm.DHTMLEdit1.QueryStatus(DECMD_SETFONTNAME)
        If State >= DECMDF_ENABLED Then
            ActiveForm.DHTMLEdit1.execCommand DECMD_SETFONTNAME, OLECMDEXECOPT_DONTPROMPTUSER, fn
        End If
    End If
    
End Sub

Private Sub FontSizeCombo_Click()
    
    On Error Resume Next
    
    Dim fs As Long
    
    fs = FontSizeCombo.ListIndex
    fs = fs + 1
    
    If (DHTMLEditInitialized) Then
        ActiveForm.DHTMLEdit1.execCommand DECMD_SETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER, fs
    End If
    
End Sub

Private Sub FormatSub_Click(index As Integer)
    
    On Error Resume Next
    
    Dim State As DHTMLEDITCMDF
    Dim Format As String
    
    State = ActiveForm.DHTMLEdit1.QueryStatus(DECMD_SETBLOCKFMT)
    
    If State >= DECMDF_ENABLED Then
        'ActiveForm.DHTMLEdit1.ExecCommand DECMD_SETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER, FormatSub(Index).Caption
    End If
    
End Sub

Private Sub Insert_Click()
    
    On Error Resume Next
    
    Dim cmdIndex As Long
    
    For cmdIndex = LBound(insertMenuCmds) To UBound(insertMenuCmds)
        'UpdateMenu InsertSub(cmdIndex), insertMenuCmds(cmdIndex)
    Next cmdIndex
    
    If ActiveForm.DHTMLEdit1.DOM.selection.Type = "Control" Then ' a control, table, ActiveX control is selected
        'InsertButton.Enabled = False
        'InsertHTML.Enabled = False
    Else
        'InsertButton.Enabled = True
        'InsertHTML.Enabled = True
    End If

End Sub

Private Sub InsertButton_Click()
    
    Dim doc As Object
    Dim selection As Object
    Dim tr As Object
    ' This routine inserts a button at the current selection
    
    ' Get the DHTML Document object
    Set doc = ActiveForm.DHTMLEdit1.DOM
    ' Get the DHTML Selection object
    Set selection = doc.selection
    ' Create a TextRange on the current selection
    Set tr = selection.createRange
    
    tr.pasteHTML ("<BUTTON TITLE=Button>Button!</BUTTON>")
    
End Sub

Private Sub mnuInsertHTML_Click()
    InsertHTMLDlg.Show vbModal, Me
End Sub

Private Sub InsertSub_Click(index As Integer)
    Dim Cmd As DHTMLEDITCMDID
    
    Cmd = insertMenuCmds(index)
           
    If Not Cmd = 0 Then
        ActiveForm.DHTMLEdit1.execCommand Cmd, OLECMDEXECOPT_DODEFAULT
    End If


End Sub

Private Sub Table_Click()
    Dim cmdIndex As Long
    
    For cmdIndex = LBound(tableMenuCmds) To UBound(tableMenuCmds)
        'UpdateMenu TableSub(cmdIndex), tableMenuCmds(cmdIndex)
    Next cmdIndex

End Sub

Private Sub D2D_Click()
    Dim cmdIndex As Long
    Dim State As DHTMLEDITCMDF
    
    For cmdIndex = LBound(twoDMenuCmds) To UBound(twoDMenuCmds)
        'UpdateMenu D2DSub(cmdIndex), twoDMenuCmds(cmdIndex)
    Next cmdIndex

    State = ActiveForm.DHTMLEdit1.QueryStatus(DECMD_LOCK_ELEMENT)
    If State = DECMDF_LATCHED Then
        'D2DSub(8).Checked = True
    Else
        'D2DSub(8).Checked = False
    End If

End Sub

Private Sub TableSub_Click(index As Integer)
    Dim Cmd As DHTMLEDITCMDID
    
    If index = 0 Then
        InsertTableDlg.Show vbModal, Me
    Else
        Cmd = tableMenuCmds(index)
               
        If Not Cmd = 0 Then
            ActiveForm.DHTMLEdit1.execCommand Cmd, OLECMDEXECOPT_DODEFAULT
        End If
    End If
    
End Sub

Private Sub RefreshToolbarButton(Button As String, Command As DHTMLEDITCMDID)
    Dim State As DHTMLEDITCMDF

    If Not Command = 0 Then
        State = ActiveForm.DHTMLEdit1.QueryStatus(Command)
        
        If (State >= DECMDF_ENABLED) Then
            Toolbar1.Buttons(Button).Value = tbrPressed
        Else
            Toolbar1.Buttons(Button).Value = tbrUnpressed
        End If
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If

End Sub


Private Sub Edit_Click()

    Dim cmdIndex As Long
    
    For cmdIndex = LBound(editMenuCmds) To UBound(editMenuCmds)
        'UpdateMenu EditSub(cmdIndex), editMenuCmds(cmdIndex)
    Next cmdIndex
        
End Sub


Private Sub ViewSub_Click(index As Integer)
    Dim State As Boolean
    
    ' Toggle different properties on DHTMLEdit.
    ' Check the menu items if the properties are set
    ' to true
    Select Case index
        Case 0
            State = ActiveForm.DHTMLEdit1.ShowBorders
            State = Not State
            ActiveForm.DHTMLEdit1.ShowBorders = State
            'ViewSub(Index).Checked = state
        Case 1
            State = ActiveForm.DHTMLEdit1.ShowDetails
            State = Not State
            ActiveForm.DHTMLEdit1.ShowDetails = State
            'ViewSub(Index).Checked = state
    End Select
        
End Sub

Private Sub Format_Click()
    Dim State As DHTMLEDITCMDF
    Dim Format As String
    Dim menuItem As Variant
    
    State = ActiveForm.DHTMLEdit1.QueryStatus(DECMD_GETBLOCKFMT)
    
    If State >= DECMDF_ENABLED Then
        Format = ActiveForm.DHTMLEdit1.execCommand(DECMD_GETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER)
        
        'For Each menuItem In FormatSub
            
            ' enable menu item
        '    menuItem.Enabled = True

            ' Check the menu that reflects the
            ' current formatting
        '    If menuItem.Caption = Format Then
        '        menuItem.Checked = True
        '    Else
        '        menuItem.Checked = False
        '    End If
            
        'Next
    ElseIf State = DECMDF_DISABLED Then
        ' disable format menu menuItems
        'For Each menuItem In FormatSub
        '    menuItem.Enabled = False
        '    menuItem.Checked = False
        'Next
    End If
End Sub

Private Sub SetFormCaption()
    If Len(ActiveForm.DHTMLEdit1.CurrentDocumentPath) > 0 Then
        'MainForm.Caption = "VBEdit - " & ActiveForm.DHTMLEdit1.CurrentDocumentPath
    Else
        'MainForm.Caption = "VBEdit"
    End If
End Sub

Private Function FormatRGBString(val As Long) As String
    Dim color As String
    Dim pad As Long
    Dim R As String
    Dim g As String
    Dim b As String
    
    ' This function formats a long consisting of rgb values
    ' taken from the CommonDialog color dialog
    ' to a string in the form of "#RRGGBB" where RRGGBB are
    ' hex values
    
    ' convert to hex
    color = Hex(val)
    'determine how many zeros to pad in front of converted value
    pad = 6 - Len(color)
    
    If pad Then
        color = String(pad, "0") & color
    End If
        
    'Extract the rgb components
    R = right(color, 2)
    g = Mid(color, 3, 2)
    b = left(color, 2)
    
    ' Swab r and b position, color dialog returns
    ' bgr instead of rgb
    color = "#" & R & g & b
    
    FormatRGBString = color
End Function

Public Sub DisableToolbar()
    
    StyleCombo.Text = ""
    StyleCombo.Enabled = False
    
    FontCombo.Text = ""
    FontCombo.Enabled = False
    
    FontSizeCombo.Text = ""
    FontSizeCombo.Enabled = False
    
    Dim b As Object
    For Each b In Toolbar1.Buttons
        b.Enabled = False
    Next

    For Each b In Toolbar2.Buttons
        b.Enabled = False
    Next

    For Each b In Toolbar3.Buttons
        b.Enabled = False
    Next
    
    Toolbar2.Buttons("New").Enabled = True
    Toolbar2.Buttons("Open").Enabled = True
    Toolbar2.Buttons("Save").Enabled = True
    Toolbar2.Buttons("SaveAs").Enabled = True
    Toolbar2.Buttons("Download").Enabled = True

    DoEvents 'give toolbar a chance to update itself
    
End Sub

Public Sub UpdateFontCombos()
    
    On Local Error Resume Next
    
    Dim State As DHTMLEDITCMDF
    
    ' Update the font name combo box on the toolbar
    State = ActiveForm.DHTMLEdit1.QueryStatus(DECMD_GETFONTNAME)
    
    If State = DECMDF_ENABLED Or State = DECMDF_LATCHED Then
        Dim fontName As String
        fontName = ActiveForm.DHTMLEdit1.execCommand(DECMD_GETFONTNAME, OLECMDEXECOPT_DONTPROMPTUSER)
        FontCombo.Text = fontName
        FontCombo.Enabled = True
    Else
        FontCombo.Text = ""
        If State = DECMDF_NINCHED Then
            FontCombo.Enabled = True
        Else
            FontCombo.Enabled = False
        End If
        
    End If
        
    ' Update the font size combo box on the toolbar
    State = ActiveForm.DHTMLEdit1.QueryStatus(DECMD_GETFONTSIZE)
    If State = DECMDF_ENABLED Or State = DECMDF_LATCHED Then
        Dim fontSize As Long
        fontSize = ActiveForm.DHTMLEdit1.execCommand(DECMD_GETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER)
        If fontSize >= 1 Then
            FontSizeCombo.Text = fontSize
        Else
            FontSizeCombo.Text = ""
        End If
        FontSizeCombo.Enabled = True
    Else
        FontSizeCombo.Text = ""
        If State = DECMDF_NINCHED Then
            FontSizeCombo.Enabled = True
        Else
            FontSizeCombo.Enabled = False
        End If
    End If
    
    '------------------------------------------------------
    ' Update the Format menu with the localized strings returned from
    ' the DECMD_GETBLOCKFMT command
    State = ActiveForm.DHTMLEdit1.QueryStatus(DECMD_GETBLOCKFMT)
    If State = DECMDF_ENABLED Or State = DECMDF_LATCHED Then
        Dim blockFmt As String
        blockFmt = ActiveForm.DHTMLEdit1.execCommand(DECMD_GETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER)
        StyleCombo.Text = blockFmt
        StyleCombo.Enabled = True
    Else
        StyleCombo.Text = ""
        If State = DECMDF_NINCHED Then
            StyleCombo.Enabled = True
        Else
            StyleCombo.Enabled = False
        End If
    End If
    '------------------------------------------------------
    
End Sub

Private Function SaveChanges() As Long
    
    Exit Function
    
    Dim retVal As Long
    If ActiveForm.DHTMLEdit1.IsDirty Then
            
        retVal = MsgBox("The current document has changed." & vbCrLf & vbCrLf & "Do you want to save changes?", vbExclamation Or vbYesNoCancel)
    
        Select Case retVal
            Case vbCancel
                SaveChanges = vbCancel
            Case vbYes
                Dim saveSuccess As Boolean
                saveSuccess = False
                If Len(ActiveForm.DHTMLEdit1.CurrentDocumentPath) > 0 Then
                    saveSuccess = SaveDocument(False)
                Else
                    saveSuccess = SaveDocument(True)
                End If
                
                If saveSuccess = True Then
                    SaveChanges = vbOK
                Else
                    SaveChanges = vbCancel
                End If
            
            Case vbNo
                SaveChanges = vbNo
        End Select
    End If
    
End Function

Public Function SaveDocument(promptUser As Boolean) As Boolean
    
    'If ActiveForm Is Nothing Then Exit Function
    On Error Resume Next
    
    Dim Text As String
    Dim sFile As String
    
    If EditMode(CStr(ActiveDocument)) = HtmlMode Then
        ActiveForm.DHTMLEdit1.DocumentHTML = ActiveForm.Editawy1.Text
        While ActiveForm.DHTMLEdit1.Busy
            DoEvents
        Wend
    End If
    
    Text = ActiveForm.DHTMLEdit1.DocumentHTML
    '----------------------------------------------------------------
    DisableToolbar
    '----------------------------------------------------------------
    
    sFile = OpenFilenames(CStr(ActiveDocument))
    If promptUser Then
        With CommonDialog1
            .Filename = sFile
            .DefaultExt = ".html"
            .DialogTitle = "Save As"
            .CancelError = False
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "Web Pages (*.htm;*.html;*.shtml;*.shtm;*.stm;*.asp;*.aspx;*.css)|*.htm;*.html;*.shtml;*.shtm;*.stm;*.asp;*.aspx;*.css|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            .ShowSave
            If Len(.Filename) = 0 Then
                SaveDocument = False
                Exit Function
            End If
            sFile = .Filename
        End With
    Else
        sFile = OpenFilenames(CStr(ActiveDocument))
    End If
    
    ActiveForm.Caption = sFile
    '----------------------------------------------------------------
    On Error GoTo ERR_Handler
    
    Dim fs As FileSystemObject
    Dim ts As TextStream
    
    Set fs = New FileSystemObject
    Set ts = fs.CreateTextFile(sFile)
    ts.Write Text
    ts.Close
    
    Set ts = Nothing
    Set fs = Nothing
    
    On Error Resume Next
    ' Force a DisplayChanged event to update toolbar
    ' in case user canceled file save dialog
    ActiveForm.DHTMLEdit1.DOM.selection.createTextRange.collapse
    SetFormCaption
    
    SaveDocument = True
    
    Exit Function
    
ERR_Handler:
    SaveDocument = False
    
    Set ts = Nothing
    Set fs = Nothing
    
'    If Err.Number < 0 Then
'        Dim errMsg As String
'        Select Case Err.Number
'            Case DE_E_INVALIDARG
'                errMsg = "Invalid argument"
'            Case DE_E_PATH_NOT_FOUND
'                errMsg = "Path not found"
'            Case DE_E_DISK_FULL
'                errMsg = "Disk is full"
'            Case DE_E_ACCESS_DENIED
'                errMsg = "Access denied"
'            Case DE_E_UNEXPECTED
'                errMsg = "Unexpected error"
'            Case Else
'                errMsg = "Unknown error"
'        End Select
'        SaveDocument = False
'        MsgBox "Error occurred while saving document: " & errMsg & ".", vbCritical
'    End If
        
    On Error Resume Next
    ' Force a DisplayChanged event to update toolbar
    ' in case user canceled file save dialog
    ActiveForm.DHTMLEdit1.DOM.selection.createTextRange.collapse
    SetFormCaption
    
End Function

'====================================================================
' object.FilterSourceCode sourceCodeIn
'====================================================================

Sub InsertHTMLCode(strHTML As String)
    
    On Error Resume Next
    
    Dim doc As Object
    Dim sel As Object
    Dim tr As Object
    
    ' get the DHTML Document object
    Set doc = ActiveForm.DHTMLEdit1.DOM
    
    ' get the IE4 selection object
    Set sel = doc.selection
    
    ' create a TextRange from the current selection
    Set tr = sel.createRange
    
    ' paste our html into the range
    tr.pasteHTML (strHTML)
    
    Set tr = Nothing
    Set sel = Nothing
    Set doc = Nothing
    
End Sub

Sub InsertForm()
    
    Dim strHTML As String
    
    FormIndex = FormIndex + 1
    
    strHTML = "<form name=""from" & CStr(FormIndex) & """" & " method=""POST"" action=""http://"">" & vbCrLf & _
            "   <br><br>" & vbCrLf & _
            "   <input type=""submit"" value=""Submit"" name=""B1"">&nbsp;&nbsp;" & vbCrLf & _
            "   <input type=""reset"" value=""Reset"" name=""B2"">" & vbCrLf & _
            "</form>"
    
    InsertHTMLCode strHTML
    
End Sub
Sub InsertTextbox()
    
    TextboxIndex = TextboxIndex + 1
    InsertHTMLCode "<input type=""text"" name=""T" & CStr(TextboxIndex) & """ value="""" size=""20"">"

End Sub
Sub InsertTextarea()
    
    TextareaIndex = TextareaIndex + 1
    InsertHTMLCode "<textarea rows=""2"" name=""S" & CStr(TextareaIndex) & """ cols=""20""></textarea>"

End Sub
Sub InsertCheckbox()
    
    CheckboxIndex = CheckboxIndex + 1
    InsertHTMLCode "<input type=""checkbox"" name=""C" & CStr(CheckboxIndex) & """ value=""ON"">"

End Sub
Sub InsertOptionButton()
    
    OptionButtonIndex = OptionButtonIndex + 1
    InsertHTMLCode "<input type=""radio"" value=""V" & CStr(OptionButtonIndex) & """ name=""R" & CStr(OptionButtonIndex) & """ checked>"

End Sub
            
Sub InsertListBox()
    
    ListBoxIndex = ListBoxIndex + 1
    InsertHTMLCode "<select size=""3"" name=""L" & CStr(ListBoxIndex) & """><option value=""Option 1 Key"">Option 1 Value</option><option value=""Option 2 Key"">Option 2 Value</option><option value=""Option 3 Key"">Option 3 Value</option></select>"

End Sub
            
Sub InsertDropDownBox()
    
    DropDownBoxIndex = DropDownBoxIndex + 1
    InsertHTMLCode "<select size=""1""  name=""D" & CStr(DropDownBoxIndex) & """><option value=""Option 1 Key"">Option 1 Value</option></select>"

End Sub


Sub InsertPushButton()
    
    PushButtonIndex = PushButtonIndex + 1
    InsertHTMLCode "<input type=""button"" value=""Button"" name=""Button" & CStr(PushButtonIndex) & """>"

End Sub
            
Sub InsertHiddenData()
    
    HiddenDataIndex = HiddenDataIndex + 1
    InsertHTMLCode "<input type=""hidden"" name=""H" & CStr(HiddenDataIndex) & """ value="""">"

End Sub
            
Sub InsertPassword()
    
    PasswordIndex = PasswordIndex + 1
    InsertHTMLCode "<input type=""password"" name=""P" & CStr(PasswordIndex) & """ value="""">"

End Sub
            
Sub InsertSubmitButton()
    
    SubmitButtonIndex = SubmitButtonIndex + 1
    InsertHTMLCode "<input type=""submit"" value=""Submit"" name=""SB" & CStr(SubmitButtonIndex) & """>"

End Sub
            
Sub InsertResetButton()
    
    ResetButtonIndex = ResetButtonIndex + 1
    InsertHTMLCode "<input type=""reset"" value=""Reset"" name=""RB" & CStr(ResetButtonIndex) & """>"

End Sub

Sub InsertImageButton()
    
    ImageButtonIndex = ImageButtonIndex + 1
    InsertHTMLCode "<input type=""image"" border=""0"" src="""" name=""I" & CStr(ImageButtonIndex) & """ width=""20"" height=""20"">"

End Sub

Sub InsertFileUpload()

    FileUploadIndex = FileUploadIndex + 1
    InsertHTMLCode "<input type=""file"" name=""F" & CStr(FileUploadIndex) & """>"

End Sub

'====================================================================
' Insert Table
Public Sub InsertTable()

    If ActiveForm Is Nothing Then MsgBox "No open documents !": Exit Sub
    CenterForm InsertTableDlg, fMainForm
    InsertTableDlg.Show vbModal, fMainForm

End Sub

Private Sub UpdateMenu(menu As Control, Command As DHTMLEDITCMDID)

    Dim State As DHTMLEDITCMDF

    If Not Command = 0 Then
        State = ActiveForm.DHTMLEdit1.QueryStatus(Command)
        
        If (State >= DECMDF_ENABLED) Then
            menu.Enabled = True
        Else
            menu.Enabled = False
        End If
    End If

End Sub
'====================================================================
'====================================================================
'Table Menu

Private Sub mnuTable_Click()
    UpdateMenu mnuInsertTable, DECMD_INSERTTABLE
    
    UpdateMenu mnuInsertCells, DECMD_INSERTCELL
    UpdateMenu mnuInsertColumns, DECMD_INSERTCOL
    UpdateMenu mnuInsertRows, DECMD_INSERTROW
    
    UpdateMenu mnuDeleteCell, DECMD_DELETECELLS
    UpdateMenu mnuDeleteColumns, DECMD_DELETECOLS
    UpdateMenu mnuDeleteRows, DECMD_DELETEROWS
    
    UpdateMenu mnuMergeCells, DECMD_MERGECELLS
    UpdateMenu mnuSplitCells, DECMD_SPLITCELL
End Sub
Private Sub mnuInsertTable_Click()
    InsertTable
End Sub

Private Sub mnuInsertCells_Click()
    CommandExec DECMD_INSERTCELL
End Sub

Private Sub mnuInsertColumns_Click()
    CommandExec DECMD_INSERTCOL
End Sub

Private Sub mnuInsertRows_Click()
    CommandExec DECMD_INSERTROW
End Sub

Private Sub mnuDeleteCell_Click()
    CommandExec DECMD_DELETECELLS
End Sub

Private Sub mnuDeleteColumns_Click()
    CommandExec DECMD_DELETECOLS
End Sub

Private Sub mnuDeleteRows_Click()
    CommandExec DECMD_DELETEROWS
End Sub

Private Sub mnuMergeCells_Click()
    CommandExec DECMD_MERGECELLS
End Sub

Private Sub mnuSplitCells_Click()
    CommandExec DECMD_SPLITCELL
End Sub

'====================================================================
'====================================================================
' Position Absoultly Menu
Private Sub mnuPosition_Click()
   
    'ActiveForm.DHTMLEdit1.SnapToGrid = DhtmlSnapToGrid
    
    If ActiveForm.DHTMLEdit1.SnapToGrid Then
        mnuSnapToGrid.Checked = True
    Else
        mnuSnapToGrid.Checked = False
    End If
    
    UpdateMenu mnuPositionAbsolutly, DECMD_MAKE_ABSOLUTE
    UpdateMenu mnuBringForward, DECMD_BRING_FORWARD
    
    UpdateMenu mnuSendBackward, DECMD_SEND_BACKWARD
    
    UpdateMenu mnuBringToFront, DECMD_BRING_TO_FRONT
    
    UpdateMenu mnuSendToBack, DECMD_SEND_TO_BACK
   
    UpdateMenu mnuBringAboveText, DECMD_BRING_ABOVE_TEXT
    
    UpdateMenu mnuSendBelowText, DECMD_SEND_BELOW_TEXT
   
    UpdateMenu mnuLockElement, DECMD_LOCK_ELEMENT
    
End Sub

Private Sub mnuPositionAbsolutly_Click()
    CommandExec DECMD_MAKE_ABSOLUTE
End Sub

Private Sub mnuBringForward_Click()
    CommandExec DECMD_BRING_FORWARD
End Sub

Private Sub mnuSendBackward_Click()
    CommandExec DECMD_SEND_BACKWARD
End Sub

Private Sub mnuBringToFront_Click()
    CommandExec DECMD_BRING_TO_FRONT
End Sub

Private Sub mnuSendToBack_Click()
    CommandExec DECMD_SEND_TO_BACK
End Sub

Private Sub mnuBringAboveText_Click()
    CommandExec DECMD_BRING_ABOVE_TEXT
End Sub

Private Sub mnuSendBelowText_Click()
    CommandExec DECMD_SEND_BELOW_TEXT
End Sub

Private Sub mnuLockElement_Click()
    CommandExec DECMD_LOCK_ELEMENT
End Sub

Private Sub mnuSnapToGrid_Click()

    DhtmlSnapToGrid = Not DhtmlSnapToGrid
    
    ActiveForm.DHTMLEdit1.SnapToGrid = DhtmlSnapToGrid
    
    If DhtmlSnapToGrid Then
           fMainForm.Toolbar3.Buttons("SnapToGrid").Value = tbrPressed
    Else
           fMainForm.Toolbar3.Buttons("SnapToGrid").Value = tbrUnpressed
    End If
    
    SaveSetting App.Title, "Settings", "DhtmlSnapToGrid", DhtmlSnapToGrid
    
End Sub

Public Function GetDocCartLocation() As Long

    On Error GoTo ErrHandler

    Dim rg As IHTMLTxtRange
    Dim ctlRg As IHTMLControlRange
    Dim Text As String
    Dim cart As Long
    
    Select Case ActiveForm.DHTMLEdit1.DOM.selection.Type
       Case "None", "Text"
          ' This reduces the selection to just the insertion
          ' point. The parentElement method will then return the
          ' element directly under the mouse pointer.
          Set rg = ActiveForm.DHTMLEdit1.DOM.selection.createRange
          rg.collapse
          cart = rg.Move("character", -999999999#)
       
       Case "Control"
          ' A form or image is selected. The commonParentElement
          ' will return the site selected element.
          Set ctlRg = ActiveForm.DHTMLEdit1.DOM.selection.createRange
          ctlRg.collapse
          cart = ctlRg.Move("character", -999999999#)
          
    End Select
    
    GetDocCartLocation = Abs(cart)
       
    Exit Function

ErrHandler:
   GetDocCartLocation = 0
   
   
End Function

'====================================================================
'====================================================================
Private Sub tvDomTree_DblClick()
    
    Exit Sub
    
    On Error Resume Next
    'HtmlMode, PreviewMode
    'If EditMode(CStr(ActiveDocument)) <> NormalMode Then Exit Sub
    If ActiveForm Is Nothing Then Exit Sub
    '----------------------------------------------------------------
    Dim doc As HTMLDocument
    Dim el As IHTMLElement
    Dim ItemNum As Long
    Dim NodeKey As String
    Dim rg As IHTMLTxtRange
    Dim ctlRg As IHTMLControlRange
    Dim TagName As String
    '----------------------------------------------------------------
    NodeKey = tvDomTree.Nodes.Item(tvDomTree.SelectedItem.index).Key
    ItemNum = CLng(Replace(NodeKey, "E", ""))
   
    Set doc = ActiveForm.DHTMLEdit1.DOM
    Set el = doc.All.Item(ItemNum, 0)
    
    TagName = LCase(el.TagName)
    
    'el.focus
    'el.scrollIntoView
    Debug.Print "TagName: " & TagName
    If TagName = "table" Or TagName = "img" Then
        Set ctlRg = ActiveForm.DHTMLEdit1.DOM.body.createControlRange
        ctlRg.Add el
        'Debug.Print ctlRg.length
        ctlRg.collapse True
        'ctlRg.focus
        ctlRg.Select
    Else
'        Set rg = ActiveForm.DHTMLEdit1.DOM.selection.createRange
'        rg.collapse
'        rg.moveToElementText el
'        rg.Select
    End If
    
End Sub

Private Sub tvDomTree_NodeClick(ByVal Node As MSComctlLib.Node)
    
    'DHTMLedit1.DOM
    'http://msdn.microsoft.com/library/default.asp?url=/workshop/author/dhtml/reference/objects/obj_document.asp
    
    'On Error GoTo ErrHandler
    Dim dp As IDisplayPointer
    Dim isp As IDisplayServices
    Dim HTMLCaret As IHTMLCaret

    Dim doc1 As IHTMLDocument2
    
   'Debug.Print doc.queryCommandEnabled(IID_IDisplayServices)
   'Debug.Print frmBrowser.brwWebBrowser.Document.QueryInterface(IID_IDisplayServices, isp)
'   Debug.Print doc1.QueryInterface(IID_IDisplayServices, isp)

    'isp.GetCaret dp
    'isp.GetCaret HTMLCaret
    'HTMLCaret.MoveDisplayPointerToCaret dp
    'HTMLCaret.MoveCaretToPointer dp.QueryBreaks

    On Error Resume Next
    
    'HtmlMode, PreviewMode
    'If EditMode(CStr(ActiveDocument)) <> NormalMode Then Exit Sub
    
    If ActiveForm Is Nothing Then Exit Sub
    '----------------------------------------------------------------
    Dim doc As HTMLDocument
    Dim el As IHTMLElement
    Dim ItemNum As Long
    Dim rg As IHTMLTxtRange
    Dim ctlRg As IHTMLControlRange
    Dim TagName As String
    
    '----------------------------------------------------------------
    ItemNum = CLng(Replace(Node.Key, "E", ""))
   
    Set doc = ActiveForm.DHTMLEdit1.DOM
    Set el = doc.All.Item(ItemNum, 0)
    
    TagName = LCase(el.TagName)
    
    'Clears the current selection.
    'ActiveForm.DHTMLEdit1.DOM.ExecCommand "Unselect"
    
   
    el.focus
    el.scrollIntoView
    
    If TagName = "table" Or TagName = "img" Then
'        Set ctlRg = ActiveForm.DHTMLEdit1.DOM.body.createControlRange
'        ctlRg.Add el
'        'Debug.Print ctlRg.length
'        ctlRg.collapse True
'        ctlRg.focus
'        ctlRg.Select
    
        Set rg = ActiveForm.DHTMLEdit1.DOM.selection.createRange
        rg.collapse
        rg.moveToElementText el
        rg.Select
        'rg.Move "character", -9999999
        el.Click
    Else
        Set rg = ActiveForm.DHTMLEdit1.DOM.selection.createRange
        rg.collapse
        rg.moveToElementText el
        
        rg.Select
        el.Click
    End If
    
    
    'colA=doc.all.tags("a")
    'Dim e As IHTMLEventObj
    'Set e = ActiveForm.DHTMLEdit1.DOM.parentWindow.event
    'window.event.srcElement
    'e.toElement = el
    'e.fromElement = el
    '----------------------------------------------------------------
    'Set rg = ActiveForm.DHTMLEdit1.DOM.selection.createRange
    'Set ctlRg = ActiveForm.DHTMLEdit1.DOM.selection.createRange
    
    'collapse True/False: Moves the insertion point to the beginning or end
    'of the current range. True Default. Moves the insertion point to
    'the beginning of the text range. False Moves the insertion point
    'to the end of the text range.
    'rg.collapse
    'rg.moveToElementText el
    'rg.Select
    'ctlRg.Select
    
'    Set ctlRg = ActiveForm.DHTMLEdit1.DOM.body.createControlRange
'    ctlRg.Add el
'    'Debug.Print ctlRg.length
'    'ctlRg.collapse True
'    ctlRg.focus
'    ctlRg.Select
    
    Exit Sub
    
'    Select Case ActiveForm.DHTMLEdit1.DOM.selection.Type
'       Case "Text", "None"
'            Set rg = ActiveForm.DHTMLEdit1.DOM.selection.createRange
'            'rg.collapse
'            rg.moveToElementText el
'            rg.Select
'            Debug.Print "Text selected"
'
'       Case "Control"
'            Set ctlRg = ActiveForm.DHTMLEdit1.DOM.selection.createRange
'            ctlRg.moveToElementText el
'            'ctlRg.commonParentElement.All
'            ctlRg.Select
'            Debug.Print "Control selected"
'    End Select
    '----------------------------------------------------------------
    
ErrHandler:
End Sub

