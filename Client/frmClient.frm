VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Administration Client"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   LinkTopic       =   "frmClient"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar statBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   65
      Top             =   7875
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraConn 
      Appearance      =   0  'Flat
      Caption         =   "Connection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      TabIndex        =   54
      Top             =   4980
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "Listen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   64
         Top             =   540
         Width           =   1095
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2400
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtIP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   56
         Text            =   "192.168.1.100"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraEnumWins 
      Appearance      =   0  'Flat
      Caption         =   "Enumerate Windows"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   60
      TabIndex        =   51
      Top             =   5940
      Width           =   9855
      Begin VB.CommandButton cmdWindowStuff 
         Caption         =   "Restore"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   9060
         TabIndex        =   62
         Top             =   1440
         Width           =   675
      End
      Begin VB.CommandButton cmdWindowStuff 
         Caption         =   "Minimize"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   8340
         TabIndex        =   60
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdWindowStuff 
         Caption         =   "Maximize"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   7560
         TabIndex        =   61
         Top             =   1440
         Width           =   795
      End
      Begin VB.CommandButton cmdWindowStuff 
         Caption         =   "Hide Window"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   8640
         TabIndex        =   58
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdWindowStuff 
         Caption         =   "Show Window"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   7560
         TabIndex        =   59
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Enumerate Visible Windows"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   39
         Left            =   7560
         TabIndex        =   53
         Top             =   240
         Width           =   2175
      End
      Begin MSComctlLib.ListView lstEnumWins 
         Height          =   1515
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2672
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Window Title"
            Object.Width           =   10348
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Window ID"
            Object.Width           =   2118
         EndProperty
      End
   End
   Begin VB.Frame fraProc 
      Appearance      =   0  'Flat
      Caption         =   "Proccess Viewer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   3720
      TabIndex        =   48
      Top             =   4500
      Width           =   6195
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Retrieve Proccess List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   19
         Left            =   4500
         TabIndex        =   57
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Terminate Process"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   34
         Left            =   4500
         TabIndex        =   49
         Top             =   600
         Width           =   1575
      End
      Begin MSComctlLib.ListView lstProcView 
         Height          =   1035
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   1826
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Proccess Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "PID"
            Object.Width           =   1765
         EndProperty
      End
   End
   Begin VB.Frame fraFileMan 
      Appearance      =   0  'Flat
      Caption         =   "File Manager"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   3720
      TabIndex        =   32
      Top             =   60
      Width           =   6195
      Begin MSWinsockLib.Winsock sckClient 
         Left            =   180
         Top             =   3780
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommandDiag 
         Left            =   3840
         Top             =   1020
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Create Directory"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   4500
         TabIndex        =   41
         Top             =   2220
         Width           =   1575
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Delete File/Folder"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   33
         Left            =   4500
         TabIndex        =   42
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "File Properties"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   20
         Left            =   4500
         TabIndex        =   40
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload File"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4500
         TabIndex        =   37
         Top             =   3540
         Width           =   1575
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Refresh List"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   31
         Left            =   4500
         TabIndex        =   33
         Top             =   1260
         Width           =   1575
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Execute Program"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4500
         TabIndex        =   46
         Top             =   2580
         Width           =   1575
      End
      Begin VB.ComboBox cmbDrives 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   240
         Width           =   5955
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Get Drive List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   30
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Rename File/Folder"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   32
         Left            =   4500
         TabIndex        =   43
         Top             =   1620
         Width           =   1575
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Get Drive Information"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   29
         Left            =   1980
         TabIndex        =   39
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdDownload 
         Caption         =   "Download File"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4500
         TabIndex        =   38
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4500
         TabIndex        =   35
         Text            =   "*.MPG; *.AVI"
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Search Drive"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   35
         Left            =   4500
         TabIndex        =   34
         Top             =   960
         Width           =   1575
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   4500
         TabIndex        =   36
         Top             =   3960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ImageList ImgListFF 
         Left            =   3720
         Top             =   3600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":015A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":02B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":040E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":0568
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":06C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":081C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":0976
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":0AD0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstFileFolder 
         Height          =   3315
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   5847
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         Icons           =   "ImgListFF"
         SmallIcons      =   "ImgListFF"
         ColHdrIcons     =   "ImgListFF"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   47
         Top             =   3960
         Width           =   675
      End
   End
   Begin VB.Frame fraWindowsCmds 
      Appearance      =   0  'Flat
      Caption         =   "Windows Commands"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3615
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Open Website"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   1260
         TabIndex        =   12
         Top             =   1860
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Logoff"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   1260
         TabIndex        =   63
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Change Wallpaper"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   21
         Left            =   1800
         TabIndex        =   20
         Top             =   3180
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Restore Original System Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   28
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   3375
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Set System Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   26
         Left            =   1800
         TabIndex        =   16
         Top             =   4140
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Send Error Message"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   38
         Left            =   1800
         TabIndex        =   7
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Hide System Clock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   24
         Left            =   1800
         TabIndex        =   18
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Show System Clock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   25
         Left            =   1800
         TabIndex        =   19
         Top             =   2220
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Hide Start Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   22
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Get Time/Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   1860
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Hide Taskbar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   2400
         TabIndex        =   27
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Hide Desktop"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   2400
         TabIndex        =   25
         Top             =   1860
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Shutdown"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   1260
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Get Clipboard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Swap Mouse Buttons"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1800
         TabIndex        =   3
         Top             =   540
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Close CD-ROM Tray"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Fix Mouse Buttons"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   1695
      End
      Begin VB.TextBox txtExecute 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Text            =   "netstat -?"
         Top             =   3540
         Width           =   2175
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Set Time/Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Show Taskbar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   2400
         TabIndex        =   28
         Top             =   900
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Show Desktop"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   2400
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Retrieve DOS Output"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   120
         TabIndex        =   24
         Top             =   3180
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Restart"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   1260
         TabIndex        =   23
         Top             =   900
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Show Start Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   23
         Left            =   120
         TabIndex        =   22
         Top             =   2220
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Get System Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   27
         Left            =   120
         TabIndex        =   17
         Top             =   4140
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Open CD-ROM Tray"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Get System Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   36
         Left            =   120
         TabIndex        =   11
         Top             =   3540
         Width           =   1155
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Take Desktop Picture"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   37
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtExecute2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   3840
         Width           =   2175
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Set Clipboard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   6
         Top             =   900
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub RtlMoveMemory Lib "KERNEL32.DLL" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function uncompress Lib "ZLIB.DLL" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Private Declare Function FindClose Lib "KERNEL32.DLL" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFileA Lib "KERNEL32.DLL" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

Private Const vbParseData As String = ""

Private CurrentDirectory As String
Private intFile As Integer
Private lngFileSize As Long
Private lngFileProg As Long
Private strFileName As String
Private RequestedFile As String
Private tmpString As String
Private ScreenShot As Boolean

Private Sub cmbDrives_Click()
    CurrentDirectory = Left$(cmbDrives.Text, 3): sckClient.SendData "29" & vbParseData & Left$(cmbDrives.Text, 3)
    cmdCommand(6).Enabled = True
    cmdCommand(29).Enabled = True
    cmdCommand(31).Enabled = True
    cmdCommand(35).Enabled = True
    cmdUpload.Enabled = True
End Sub

Private Sub cmdCommand_Click(Index As Integer)
    Dim tmpString As String
    Select Case Index
        Case 0, 1, 2, 3, 8, 10, 11, 12, 13, 14, 19, 22, 23, 24, 25, 27, 28, 30, 36, 39: sckClient.SendData Index & vbParseData
        Case 4:
                If (lstFileFolder.SelectedItem.Key <> "Previous") Then
                    If Mid$(lstFileFolder.SelectedItem.Text, 2, 2) = ":\" Then
                        sckClient.SendData Index & vbParseData & lstFileFolder.SelectedItem.Text
                    Else
                        sckClient.SendData Index & vbParseData & CurrentDirectory & lstFileFolder.SelectedItem.Text
                    End If
                End If
        Case 5: sckClient.SendData "4" & vbParseData & txtExecute.Text
        Case 6: tmpString = InputBox("Please specify a directory name.", "Create Directory")
                If LenB(tmpString) <> 0 Then
                    sckClient.SendData Index & vbParseData & CurrentDirectory & tmpString
                End If
        Case 9: sckClient.SendData Index & vbParseData & txtExecute.Text & vbParseData & txtExecute2.Text
        Case 15: sckClient.SendData "15" & vbParseData & "0"
        Case 16: sckClient.SendData "15" & vbParseData & "1"
        Case 17: sckClient.SendData "15" & vbParseData & "2"
        Case 20:
                 If (lstFileFolder.SelectedItem.Key <> "Previous") Then
                     If Mid$(lstFileFolder.SelectedItem.Text, 2, 2) = ":\" Then
                         sckClient.SendData "32" & vbParseData & lstFileFolder.SelectedItem.Text
                     Else
                         sckClient.SendData "32" & vbParseData & CurrentDirectory & lstFileFolder.SelectedItem.Text
                     End If
                 End If
        Case 26: sckClient.SendData Index & vbParseData & frmColor.sysColor(4).BackColor & ChrW$(2) & frmColor.sysColor(5).BackColor & ChrW$(2) & frmColor.sysColor(7).BackColor & ChrW$(2) & frmColor.sysColor(8).BackColor & ChrW$(2) & frmColor.sysColor(15).BackColor
        Case 29: sckClient.SendData "33" & vbParseData & Left$(cmbDrives.Text, 3)
        Case 31: sckClient.SendData "29" & vbParseData & CurrentDirectory: Exit Sub
        Case 32: tmpString = InputBox("Please specify a new name.", "Rename File")
                 If LenB(tmpString) <> 0 Then
                     If (lstFileFolder.SelectedItem.Key <> "Previous") Then
                        If Mid$(lstFileFolder.SelectedItem.Text, 2, 2) = ":\" Then
                            MsgBox "Can't rename search results."
                        Else
                            sckClient.SendData "31" & vbParseData & CurrentDirectory & lstFileFolder.SelectedItem.Text & vbParseData & CurrentDirectory & tmpString
                        End If
                    End If
                 End If
        Case 33:
                If Mid$(lstFileFolder.SelectedItem.Text, 2, 2) = ":\" Then
                    MsgBox "Can't delete search results."
                Else
                    If Left$(lstFileFolder.SelectedItem.Key, 9) = "Directory" Then
                        sckClient.SendData 5 & vbParseData & CurrentDirectory & lstFileFolder.SelectedItem.Text
                    ElseIf Left$(lstFileFolder.SelectedItem.Key, 4) = "File" Then
                        sckClient.SendData 20 & vbParseData & CurrentDirectory & lstFileFolder.SelectedItem.Text
                    End If
                End If
        Case 34: sckClient.SendData Index & vbParseData & lstProcView.SelectedItem.ListSubItems.Item(1).Text
        Case 35: sckClient.SendData Index & vbParseData & Left$(cmbDrives.Text, 3) & vbParseData & txtSearch.Text: Exit Sub
        Case 37: ScreenShot = True: sckClient.SendData Index & vbParseData & "0"
        Case 38: frmError.Enabled = True: frmError.Visible = True: Exit Sub
        Case Else: If LenB(txtExecute.Text) Then sckClient.SendData Index & vbParseData & txtExecute.Text
    End Select
    statBar.SimpleText = "Status: Executing Command..."
End Sub

Private Sub cmdConnect_Click()
    sckClient.Connect txtIP.Text, 6116
    statBar.SimpleText = "Status: Connecting..."
End Sub

Private Sub cmdDisconnect_Click()
    sckClient.Close
    statBar.SimpleText = "Status: Connection Idle"
End Sub

Private Sub cmdDownload_Click()
    If Mid$(lstFileFolder.SelectedItem.Text, 2, 2) <> ":\" Then
        CommandDiag.FileName = lstFileFolder.SelectedItem.Text
    End If
    CommandDiag.Filter = "All Files (*.*)|*.*"
    CommandDiag.ShowSave
    If LenB(CommandDiag.FileName) Then
        strFileName = CommandDiag.FileName
        If Mid$(lstFileFolder.SelectedItem.Text, 2, 2) = ":\" Then
            If Left$(lstFileFolder.SelectedItem.Key, 4) = "File" Then sckClient.SendData "REQU_FILE" & vbParseData & lstFileFolder.SelectedItem.Text: RequestedFile = lstFileFolder.SelectedItem.Text
        Else
            If Left$(lstFileFolder.SelectedItem.Key, 4) = "File" Then sckClient.SendData "REQU_FILE" & vbParseData & CurrentDirectory & lstFileFolder.SelectedItem.Text: RequestedFile = CurrentDirectory & lstFileFolder.SelectedItem.Text
        End If
    End If
    statBar.SimpleText = "Status: Executing Command..."
End Sub

Private Sub cmdUpload_Click()
    CommandDiag.FileName = vbNullString
    CommandDiag.Filter = "All Files (*.*)|*.*"
    CommandDiag.ShowOpen
    If LenB(CommandDiag.FileName) Then
        strFileName = CommandDiag.FileName
        ProgressBar.Max = FileLen(CommandDiag.FileName)
        lblPercent.Caption = "0%"
        sckClient.SendData "SEND_FILE" & vbParseData & CurrentDirectory & CommandDiag.FileTitle & vbParseData & FileLen(CommandDiag.FileName)
    End If
    statBar.SimpleText = "Status: Executing Command..."
End Sub

Private Sub cmdWindowStuff_Click(Index As Integer)
    sckClient.SendData "40" & vbParseData & lstEnumWins.SelectedItem.ListSubItems.Item(1).Text & vbParseData & Index
    statBar.SimpleText = "Status: Executing Command..."
End Sub

Private Sub Command1_Click()
    sckClient.LocalPort = 6116
    sckClient.Listen
    statBar.SimpleText = "Status: Listening..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmColor
    Unload frmError
    Unload Me
End Sub

Private Sub sckClient_ConnectionRequest(ByVal requestID As Long)
    sckClient.Close
    sckClient.Accept requestID
    statBar.SimpleText = "Status: Connected: Time - " & Format$(Now, "Hh:Nn:Ss AM/PM") & " Date - " & Format$(Now, "dddddd")
End Sub

Private Sub Form_Load()
    statBar.SimpleText = "Status: Connection Idle"
End Sub

Private Sub lstEnumWins_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdWindowStuff(0).Enabled = True
    cmdWindowStuff(3).Enabled = True
    cmdWindowStuff(5).Enabled = True
    cmdWindowStuff(6).Enabled = True
    cmdWindowStuff(9).Enabled = True
End Sub

Private Sub lstFileFolder_DblClick()
    Dim I As Integer
    
    If lstFileFolder.SelectedItem.Key = "Previous" Then
        For I = 1 To Len(CurrentDirectory)
            If Mid$(CurrentDirectory, Len(CurrentDirectory) - I, 1) = "\" Then
                CurrentDirectory = Left$(CurrentDirectory, Len(CurrentDirectory) - I)
                Exit For
            End If
        Next I
        sckClient.SendData "29" & vbParseData & CurrentDirectory
    ElseIf Left$(lstFileFolder.SelectedItem.Key, 9) = "Directory" Then
        CurrentDirectory = CurrentDirectory & lstFileFolder.SelectedItem.Text & "\"
        sckClient.SendData "29" & vbParseData & CurrentDirectory
    ElseIf Left$(lstFileFolder.SelectedItem.Key, 4) = "File" Then
        sckClient.SendData "4" & vbParseData & CurrentDirectory & lstFileFolder.SelectedItem.Text
    End If
End Sub

Private Sub lstFileFolder_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdCommand(4).Enabled = True
    cmdCommand(6).Enabled = True
    cmdCommand(20).Enabled = True
    cmdCommand(31).Enabled = True
    cmdCommand(32).Enabled = True
    cmdCommand(33).Enabled = True
    cmdDownload.Enabled = True
End Sub

Private Sub lstProcView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdCommand(34).Enabled = True
End Sub

Private Sub sckClient_Connect()
    statBar.SimpleText = "Status: Connected: Time - " & Format$(Now, "Hh:Nn:Ss AM/PM") & " Date - " & Format$(Now, "dddddd")
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Connection Failure", vbCritical, "Error"
    statBar.SimpleText = "Status: Connection Idle"
    sckClient.Close
End Sub

Private Function FileExists(ByVal strFileName As String) As Boolean
    Dim WFD As WIN32_FIND_DATA, hFile As Long
    
    hFile = FindFirstFileA(strFileName, WFD)
    FileExists = hFile <> -1
    FindClose hFile
End Function

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim I As Integer
    Dim tmpString2 As String
    Dim XNum As Integer
    Dim IconNum As Integer
    Dim FileBuffer As String
    Dim strDataParse() As String
    
    sckClient.GetData strData, vbString
    strDataParse = Split(strData, vbParseData)
    Debug.Print Len(strData)
    Select Case strDataParse(0)
    
        'File Transfer Commands
        Case "SEND_FILE": lngFileSize = Int(strDataParse(1)): ProgressBar.Max = lngFileSize: lblPercent.Caption = "0%": intFile = FreeFile
                          If FileExists(strFileName) Then
                              Kill strFileName: DoEvents
                          End If
                          Open strFileName For Binary Access Write As intFile: sckClient.SendData "ACPT_FILE" & vbParseData & RequestedFile: lngFileProg = 0
        
        Case "ACPT_FILE": intFile = FreeFile: Open strFileName For Binary Access Read As intFile: GoTo SendChunk
        Case "CHNK_FILE": GoTo SendChunk
        Case "DONE_FILE": Close intFile: cmdCommand_Click (31): ProgressBar.Value = 0: lblPercent.Caption = vbNullString: statBar.SimpleText = "Status: Command Complete": MsgBox "Upload Complete" & vbNewLine & vbNewLine & "Total Sent: " & Format$(ProgressBar.Max, "###,###,###,###") & " bytes", vbOKOnly, "Upload Status"
        Case "CHNK_DATA"
            strData = Mid$(strData, 11, Len(strData) - 10)
            If (lngFileProg + 4096) < lngFileSize Then
                lngFileProg = lngFileProg + 4096
                ProgressBar.Value = ProgressBar.Value + 4096
                lblPercent.Caption = Format$((ProgressBar.Value / ProgressBar.Max * 100), "##.##") & "%"
                Put intFile, , strData
                sckClient.SendData "CHNK_FILE"
            Else
                strData = Left$(strData, lngFileSize - lngFileProg)
                ProgressBar.Value = ProgressBar.Max
                lblPercent.Caption = "100%"
                Put intFile, , strData
                sckClient.SendData "DONE_FILE"
                Close intFile
                ProgressBar.Value = 0
                lblPercent.Caption = vbNullString
                If ScreenShot = True And strFileName = App.Path & "\CompressedSS.cmp" Then
                    ScreenShot = False
                    DecompressFile App.Path & "\CompressedSS.cmp", App.Path & "\ScreenShot.bmp"
                    MsgBox "Compressed Size: " & Format$(FileLen(App.Path & "\CompressedSS.cmp"), "###,###,### bytes") & vbNewLine & "Decompressed Size: " & Format$(FileLen(App.Path & "\ScreenShot.bmp"), "###,###,### bytes")
                    Kill App.Path & "\CompressedSS.cmp"
                    ShellExecute Me.hwnd, "Open", App.Path & "\ScreenShot.bmp", 0&, 0&, 3
                    DoEvents
                    sckClient.SendData "37" & vbParseData & "1" & vbParseData & "0"
                Else
                    statBar.SimpleText = "Status: Command Complete"
                    MsgBox "Download Complete" & vbNewLine & vbNewLine & "Total Received: " & Format$(ProgressBar.Max, "###,###,###,###") & " bytes", vbOKOnly, "Download Status"
                End If
            End If
    
        'MsgBox Data Info
        Case 0, 1, 2: MsgBox strDataParse(1), vbOKOnly, "Server Feedback"
        statBar.SimpleText = "Status: Command Complete"
        
        'Proccess List
        Case 3: lstProcView.ListItems.Clear
                For I = 1 To (UBound(strDataParse) - 1)
                    lstProcView.ListItems.Add(, "Proccess" & I, strDataParse(I)).SubItems(1) = strDataParse(I + 1)
                    I = I + 1
                Next I
                statBar.SimpleText = "Status: Command Complete"
        
        'Colors
        Case 4: frmColor.Enabled = True: frmColor.Visible = True
                frmColor.sysColor(4).BackColor = strDataParse(1)
                frmColor.sysColor(5).BackColor = strDataParse(2)
                frmColor.sysColor(7).BackColor = strDataParse(3)
                frmColor.sysColor(8).BackColor = strDataParse(4)
                frmColor.sysColor(15).BackColor = strDataParse(5)
                statBar.SimpleText = "Status: Command Complete"
                
        'Directory List
        Case 5: lstFileFolder.ListItems.Clear
                If Len(CurrentDirectory) > 3 Then If Mid$(strDataParse(1), 2, 1) <> ":" Then lstFileFolder.ListItems.Add(, "Previous", "..").SmallIcon = 1
                For I = 1 To (UBound(strDataParse) - 1)
                    If Left$(strDataParse(I), 1) = ChrW$(2) Then 'Directory
                        lstFileFolder.ListItems.Add(, "Directory" & I, Right$(strDataParse(I), (Len(strDataParse(I)) - 1))).SmallIcon = 1
                    Else 'File
                        Select Case LCase$(Mid$(strDataParse(I), InStrRev(strDataParse(I), ".") + 1))
                            Case "htm", "html", "txt", "doc", "ini": IconNum = 8
                            Case "exe", "bat", "com", "scr": IconNum = 3
                            Case "sys", "dll", "vxd", "cpl": IconNum = 4
                            Case "ogg", "mp3", "midi", "wav", "ram", "rm", "mp2", "mpga", "mid": IconNum = 6
                            Case "divx", "mpeg", "mpg", "avi", "asf", "swf", "wmv", "wma", "asx", "mov", "mpe", "qt": IconNum = 7
                            Case "jpg", "gif", "png", "bmp", "pdf", "jpe", "jpeg": IconNum = 5
                            Case "rar", "zip", "cab", "iso", "ace", "r00": IconNum = 9
                            Case Else: IconNum = 2
                        End Select
                        lstFileFolder.ListItems.Add(, "File" & I, strDataParse(I)).SmallIcon = IconNum
                    End If
                Next I
                If LenB(strDataParse(UBound(strDataParse))) <> 0 Then
                    tmpString = strDataParse(UBound(strDataParse))
                Else
                    tmpString = vbNullString
                End If
                
        'Drive List
        Case 6: cmbDrives.Clear
                For I = 1 To (UBound(strDataParse) - 1)
                    cmbDrives.AddItem strDataParse(I)
                Next I
                statBar.SimpleText = "Status: Command Complete"
        
        'File/Folder Properties
        Case 7: MsgBox "File Type: " & strDataParse(1) & vbNewLine & "File Size: " & Format$(strDataParse(2), "###,###,###,###") & " bytes" & vbNewLine & vbNewLine & "Read-Only: " & (strDataParse(3) = True) & vbNewLine & "Hidden: " & (strDataParse(4) = True) & vbNewLine & "System: " & (strDataParse(5) = True), vbOKOnly, "File/Folder Information"
        statBar.SimpleText = "Status: Command Complete"
        
        'Refresh Directory
        Case 8: cmdCommand_Click (31)
        statBar.SimpleText = "Status: Command Complete"
        
        'Drive Properties
        Case 9: MsgBox "Drive Label: " & strDataParse(1) & vbNewLine & "File System: " & strDataParse(2) & vbNewLine & vbNewLine & "Total Space: " & Format$(strDataParse(3), "###,###,###,###") & " bytes" & vbNewLine & "Free Space: " & Format$(strDataParse(4), "###,###,###,###") & " bytes" & vbNewLine & "Used Space: " & Format$((Int(strDataParse(3)) - Int(strDataParse(4))), "###,###,###,###") & " bytes", vbOKOnly, "Drive Information"
        statBar.SimpleText = "Status: Command Complete"

        'Refresh Process List
        Case 10: cmdCommand_Click (19)
        statBar.SimpleText = "Status: Command Complete"

        'System Information
        Case 11
            tmpString2 = "Windows Version: " & strDataParse(1) & vbNewLine & vbNewLine & "Registered Owner: " & strDataParse(2) & vbNewLine & "Registered Organization: " & strDataParse(3) & vbNewLine & "Product ID: " & strDataParse(4) & vbNewLine & vbNewLine & "Processor: " & strDataParse(5) & vbNewLine & "Memory Usage: " & Format$((strDataParse(6) / 1024), "###,###") & " / " & Format$((strDataParse(7) / 1024), "###,###") & " KB" & vbNewLine & vbNewLine & "Screen Resolution: " & strDataParse(8) & " by " & strDataParse(9) & " pixels" & vbNewLine & "Windows Uptime: " & CStr(Int(strDataParse(10) / 86400000)) & " Days, " & CStr(Int((strDataParse(10) Mod 86400000) / 3600000)) & " Hours, " & CStr(Int(((strDataParse(10) Mod 86400000) Mod 3600000) / 60000)) & " Minutes, " & CStr(Int((((strDataParse(10) Mod 86400000) Mod 3600000) Mod 60000) / 1000)) & " Seconds "
            MsgBox tmpString2, vbOKOnly, "System Information"
            statBar.SimpleText = "Status: Command Complete"
        
        'Screenshot
        Case 12: If LenB(strDataParse(1)) Then strFileName = App.Path & "\CompressedSS.cmp": sckClient.SendData "REQU_FILE" & vbParseData & strDataParse(1): RequestedFile = strDataParse(1): statBar.SimpleText = "Status: Command Complete"
        
        'Enumerate Windows
        Case 13: lstEnumWins.ListItems.Clear
                For I = 1 To (UBound(strDataParse) - 1)
                    lstEnumWins.ListItems.Add(, "Window" & I, strDataParse(I)).SubItems(1) = strDataParse(I + 1)
                    I = I + 1
                Next I
                statBar.SimpleText = "Status: Command Complete"
       
        'Command Complete
        Case 15: statBar.SimpleText = "Status: Command Complete"

        Case Else
                XNum = lstFileFolder.ListItems.Count
                If LenB(tmpString) <> 0 Then
                    If Left$(tmpString & strDataParse(0), 1) = ChrW$(2) Then 'Directory
                        lstFileFolder.ListItems.Add(, "Directory" & (1 + XNum), Right$(tmpString & strDataParse(0), (Len(tmpString & strDataParse(0)) - 1))).SmallIcon = 1
                    Else 'File
                        Select Case LCase$(Mid$(tmpString & strDataParse(0), InStrRev(tmpString & strDataParse(0), ".") + 1))
                            Case "htm", "html", "txt", "doc", "ini": IconNum = 8
                            Case "exe", "bat", "com", "scr": IconNum = 3
                            Case "sys", "dll", "vxd", "cpl": IconNum = 4
                            Case "ogg", "mp3", "midi", "wav", "ram", "rm", "mp2", "mpga", "mid": IconNum = 6
                            Case "divx", "mpeg", "mpg", "avi", "asf", "swf", "wmv", "wma", "asx", "mov", "mpe", "qt": IconNum = 7
                            Case "jpg", "gif", "png", "bmp", "pdf", "jpe", "jpeg": IconNum = 5
                            Case "rar", "zip", "cab", "iso", "ace", "r00": IconNum = 9
                            Case Else: IconNum = 2
                        End Select
                        lstFileFolder.ListItems.Add(, "File" & (XNum + 1), tmpString & strDataParse(0)).SmallIcon = IconNum
                    End If
                    XNum = XNum + 1
                End If
                
                For I = 1 To (UBound(strDataParse) - 1)
                    If Left$(strDataParse(I), 1) = ChrW$(2) Then 'Directory
                        lstFileFolder.ListItems.Add(, "Directory" & I + XNum, Right$(strDataParse(I), (Len(strDataParse(I)) - 1))).SmallIcon = 1
                    Else 'File
                        Select Case LCase$(Mid$(strDataParse(I), InStrRev(strDataParse(I), ".") + 1))
                            Case "htm", "html", "txt", "doc", "ini": IconNum = 8
                            Case "exe", "bat", "com", "scr": IconNum = 3
                            Case "sys", "dll", "vxd", "cpl": IconNum = 4
                            Case "ogg", "mp3", "midi", "wav", "ram", "rm", "mp2", "mpga", "mid": IconNum = 6
                            Case "divx", "mpeg", "mpg", "avi", "asf", "swf", "wmv", "wma", "asx", "mov", "mpe", "qt": IconNum = 7
                            Case "jpg", "gif", "png", "bmp", "pdf", "jpe", "jpeg": IconNum = 5
                            Case "rar", "zip", "cab", "iso", "ace", "r00": IconNum = 9
                            Case Else: IconNum = 2
                        End Select
                        lstFileFolder.ListItems.Add(, "File" & (I + XNum), strDataParse(I)).SmallIcon = IconNum
                    End If
                Next I

    End Select
    Erase strDataParse
    Exit Sub
SendChunk:
    FileBuffer = Space$(4096)
    If (ProgressBar.Value + 4096) > ProgressBar.Max Then
        ProgressBar.Value = ProgressBar.Max
        lblPercent.Caption = "100%"
    Else
        ProgressBar.Value = ProgressBar.Value + 4096
        lblPercent.Caption = Format$((ProgressBar.Value / ProgressBar.Max * 100), "##.##") & "%"
    End If
    Get intFile, , FileBuffer
    sckClient.SendData "CHNK_DATA" & vbParseData & FileBuffer
    Erase strDataParse
End Sub

Private Sub DecompressFile(ByVal FilePathIn As String, ByVal FilePathOut As String)
    Dim TheBytes()   As Byte
    Dim lngFileLen   As Long
    Dim BufferSize   As Long
    Dim TempBuffer() As Byte
    Dim intZlib As Integer
    
    ReDim TheBytes(FileLen(FilePathIn) - 1)
    intZlib = FreeFile
    Open FilePathIn For Binary Access Read As intZlib
        Get intZlib, , lngFileLen
        Get intZlib, , TheBytes()
    Close intZlib
    BufferSize = lngFileLen
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    ReDim TempBuffer(BufferSize)
    uncompress TempBuffer(0), BufferSize, TheBytes(0), UBound(TheBytes) + 1
    ReDim Preserve TheBytes(BufferSize - 1)
    RtlMoveMemory TheBytes(0), TempBuffer(0), BufferSize
    Erase TempBuffer
    intZlib = FreeFile
    If FileExists(FilePathOut) Then
        Kill FilePathOut
        DoEvents
    End If
    Open FilePathOut For Binary Access Write As intZlib
        Put intZlib, , TheBytes()
    Close intZlib
    Erase TheBytes
End Sub
