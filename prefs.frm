VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form prefs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "prefs.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   2610
      Top             =   990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2115
      TabIndex        =   12
      Top             =   1845
      Width           =   915
   End
   Begin VB.CommandButton cApply 
      Caption         =   "&Apply"
      Height          =   330
      Left            =   3330
      TabIndex        =   11
      Top             =   1845
      Width           =   915
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1770
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   3122
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "prefs.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tLogPath"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "p_p"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkMini"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkListen"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkClrLog"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Messages"
      TabPicture(1)   =   "prefs.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tMsg2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkMsg2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "tMsg1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkMsg1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblNotice"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CheckBox chkClrLog 
         Caption         =   "Clear Logfile At StartUp"
         Height          =   195
         Left            =   225
         TabIndex        =   14
         Top             =   1125
         Width           =   2715
      End
      Begin VB.TextBox tMsg2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -72705
         TabIndex        =   10
         Top             =   810
         Width           =   1950
      End
      Begin VB.CheckBox chkMsg2 
         Caption         =   "Message on Disconnect:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74820
         TabIndex        =   9
         Top             =   855
         Width           =   2085
      End
      Begin VB.TextBox tMsg1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -72705
         TabIndex        =   8
         Top             =   450
         Width           =   1950
      End
      Begin VB.CheckBox chkMsg1 
         Caption         =   "Message on Connect: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74820
         TabIndex        =   7
         Top             =   495
         Width           =   1860
      End
      Begin VB.CheckBox chkListen 
         Caption         =   "Start Listen At Startup"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   810
         Width           =   2985
      End
      Begin VB.CheckBox chkMini 
         Caption         =   "Minimize to System Tray"
         Enabled         =   0   'False
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   1440
         Width           =   3210
      End
      Begin VB.PictureBox p_p 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3825
         ScaleHeight     =   225
         ScaleWidth      =   360
         TabIndex        =   3
         Top             =   405
         Width           =   420
         Begin VB.CommandButton cBrowse 
            Caption         =   "..."
            Height          =   240
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.TextBox tLogPath 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   675
         TabIndex        =   2
         Top             =   405
         Width           =   3075
      End
      Begin VB.Label lblNotice 
         Caption         =   "Note: The 'Message on Disconnect' will only be sent if you are the one that closes the connection. duh!"
         ForeColor       =   &H8000000D&
         Height          =   465
         Left            =   -74820
         TabIndex        =   13
         Top             =   1215
         Width           =   4065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Logfile:"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   450
         Width           =   525
      End
   End
End
Attribute VB_Name = "prefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cApply_Click()
SaveSetting "Kahadeva", "PortListen", "Msg1Txt", tMsg1.Text
SaveSetting "Kahadeva", "PortListen", "Msg2Txt", tMsg2.Text
SaveSetting "Kahadeva", "PortListen", "LogFile", tLogPath.Text
Unload Me
End Sub

Private Sub cBrowse_Click()
On Error Resume Next
With cdlg
    .DialogTitle = "Specify logfile"
    .Filter = "Logfiles | *.Log|All Files |*.*"
    .ShowOpen
End With
tLogPath.Text = cdlg.filename
End Sub

Private Sub cCancel_Click()
Unload Me
End Sub

Private Sub chkClrLog_Click()
Select Case chkClrLog.Value
    Case Checked
        SaveSetting "Kahadeva", "PortListen", "ClearLog?", "1"
    Case Unchecked
        SaveSetting "Kahadeva", "PortListen", "ClearLog?", "0"
End Select
End Sub

Private Sub chkListen_Click()
Select Case chkListen.Value
    Case Checked
        SaveSetting "Kahadeva", "PortListen", "StartListen?", "1"
    Case Else
        SaveSetting "Kahadeva", "PortListen", "StartListen?", "0"
End Select
End Sub

Private Sub chkMsg1_Click()
Select Case chkMsg1.Value
    Case Checked
        tMsg1.Enabled = True
        SaveSetting "Kahadeva", "PortListen", "Msg1?", "1"
    Case Else
        tMsg1.Enabled = False
        SaveSetting "Kahadeva", "PortListen", "Msg1?", "0"
End Select
End Sub

Private Sub chkMsg2_Click()
Select Case chkMsg2.Value
    Case Checked
        tMsg2.Enabled = True
        SaveSetting "Kahadeva", "PortListen", "Msg2?", "1"
    Case Else
        tMsg2.Enabled = False
        SaveSetting "Kahadeva", "PortListen", "Msg2?", "0"
End Select
End Sub

Private Sub Form_Load()
sRet = GetSetting("Kahadeva", "PortListen", "LogFile")
    tLogPath.Text = sRet
sRet = GetSetting("Kahadeva", "PortListen", "Msg1?")
    If sRet = "1" Then chkMsg1.Value = Checked Else chkMsg1.Value = Unchecked: tMsg1.Enabled = False
sRet = GetSetting("Kahadeva", "PortListen", "Msg2?")
    If sRet = "1" Then chkMsg2.Value = Checked Else chkMsg2.Value = Unchecked: tMsg2.Enabled = False
sRet = GetSetting("Kahadeva", "PortListen", "Msg1Txt")
    tMsg1.Text = sRet
sRet = GetSetting("Kahadeva", "PortListen", "Msg2Txt")
    tMsg2.Text = sRet
sRet = GetSetting("Kahadeva", "PortListen", "ClearLog?")
    If sRet = "1" Then chkClrLog.Value = Checked Else chkClrLog.Value = Unchecked
sRet = GetSetting("Kahadeva", "PortListen", "StartListen?")
    If sRet = "1" Then chkListen.Value = Checked Else chkListen.Value = Unchecked

End Sub

