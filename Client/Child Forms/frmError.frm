VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Error Message"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Enabled         =   0   'False
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraButton 
      Appearance      =   0  'Flat
      Caption         =   "Message Button"
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
      Height          =   1155
      Left            =   60
      TabIndex        =   10
      Top             =   1140
      Width           =   3255
      Begin VB.OptionButton optMsgButtons 
         Caption         =   "Abort, Retry, Ignore"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   780
         Width           =   1635
      End
      Begin VB.OptionButton optMsgButtons 
         Caption         =   "Retry, Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1740
         TabIndex        =   15
         Top             =   780
         Width           =   1215
      End
      Begin VB.OptionButton optMsgButtons 
         Caption         =   "Yes, No, Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1740
         TabIndex        =   14
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton optMsgButtons 
         Caption         =   "Yes, No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1740
         TabIndex        =   13
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton optMsgButtons 
         Caption         =   "OK, Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton optMsgButtons 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraText 
      Appearance      =   0  'Flat
      Caption         =   "Text Message"
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
      Height          =   1335
      Left            =   60
      TabIndex        =   5
      Top             =   2340
      Width           =   3255
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send Message"
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
         Left            =   1620
         TabIndex        =   17
         Top             =   900
         Width           =   1515
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Message"
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
         Left            =   120
         TabIndex        =   18
         Top             =   900
         Width           =   1515
      End
      Begin VB.TextBox txtTitle 
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
         Left            =   720
         TabIndex        =   7
         Text            =   "OCX Missing"
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtPrompt 
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
         Left            =   720
         TabIndex        =   6
         Text            =   "MSWINSCK.OCX is not present."
         Top             =   540
         Width           =   2415
      End
      Begin VB.Label lblPrompt 
         Alignment       =   1  'Right Justify
         Caption         =   "Prompt:"
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
         Left            =   60
         TabIndex        =   9
         Top             =   540
         Width           =   555
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Title:"
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
         Left            =   180
         TabIndex        =   8
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame fraIcon 
      Appearance      =   0  'Flat
      Caption         =   "Message Icon"
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
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3255
      Begin VB.OptionButton optMsgIcon 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   16
         Left            =   120
         Picture         =   "frmError.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optMsgIcon 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   48
         Left            =   1680
         Picture         =   "frmError.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
      Begin VB.OptionButton optMsgIcon 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   32
         Left            =   900
         Picture         =   "frmError.frx":08A0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   675
      End
      Begin VB.OptionButton optMsgIcon 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   64
         Left            =   2460
         Picture         =   "frmError.frx":0CEA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MsgIcon As Integer
Private MsgButtons As Integer

Private Const vbParseData As String = ""

Private Sub Form_Unload(Cancel As Integer)
    frmError.Enabled = False
    frmError.Visible = False
    Unload Me
End Sub

Private Sub optMsgButtons_Click(Index As Integer)
    MsgButtons = Index
End Sub

Private Sub optMsgIcon_Click(Index As Integer)
    MsgIcon = Index
End Sub

Private Sub cmdSend_Click()
    Dim MsgIconButtons As Integer
    If MsgIcon = 0 Then MsgIcon = 16
    MsgIconButtons = MsgIcon + MsgButtons
    frmClient.sckClient.SendData "38" & vbParseData & txtPrompt.Text & vbParseData & MsgIconButtons & vbParseData & txtTitle.Text
    frmError.Enabled = False
    frmError.Visible = False
    frmClient.statBar.SimpleText = "Status: Executing Command..."
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim MsgIconButtons As Integer
    If MsgIcon = 0 Then MsgIcon = 16
    MsgIconButtons = MsgIcon + MsgButtons
    MsgBox txtPrompt.Text, MsgIconButtons, txtTitle.Text
End Sub
