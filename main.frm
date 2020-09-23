VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kahadeva Port Listener v0.66"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock tcp 
      Left            =   3150
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton cSendStr 
      Caption         =   "Send &Text"
      Height          =   330
      Left            =   2475
      TabIndex        =   4
      Top             =   3645
      Width           =   1140
   End
   Begin VB.CommandButton cExit 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   1215
      TabIndex        =   9
      Top             =   3645
      Width           =   1095
   End
   Begin VB.CommandButton cAbout 
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   8
      Top             =   3645
      Width           =   1095
   End
   Begin VB.Frame fra 
      Caption         =   "Connection Info"
      Height          =   1620
      Index           =   0
      Left            =   45
      TabIndex        =   16
      Top             =   1935
      Width           =   2265
      Begin VB.Frame Frame1 
         Height          =   60
         Left            =   90
         TabIndex        =   23
         Top             =   765
         Width           =   2085
      End
      Begin VB.Label lblPktSize 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1125
         TabIndex        =   25
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Pkt Size:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   24
         Top             =   1125
         Width           =   960
      End
      Begin VB.Label lblTime 
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   585
         TabIndex        =   22
         Top             =   900
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Time:"
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   900
         Width           =   390
      End
      Begin VB.Label lblBytes 
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1350
         TabIndex        =   20
         Top             =   1350
         Width           =   870
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes Recieved:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   19
         Top             =   1350
         Width           =   1230
      End
      Begin VB.Label lblHostname 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   540
         Width           =   1995
      End
      Begin VB.Label lblIP 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   315
         Width           =   2040
      End
   End
   Begin VB.CommandButton cmore 
      Caption         =   "&More..."
      Height          =   285
      Left            =   1305
      TabIndex        =   3
      Top             =   1620
      Width           =   870
   End
   Begin VB.CommandButton cSelect 
      Caption         =   "Copy All"
      Height          =   330
      Left            =   3690
      TabIndex        =   5
      Top             =   3645
      Width           =   1005
   End
   Begin VB.CommandButton cCopy 
      Caption         =   "&Copy"
      Height          =   330
      Left            =   4770
      TabIndex        =   6
      Top             =   3645
      Width           =   870
   End
   Begin VB.CommandButton cClear 
      Caption         =   "Cl&ear"
      Height          =   330
      Left            =   5715
      TabIndex        =   7
      Top             =   3645
      Width           =   870
   End
   Begin VB.CommandButton cListen 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   285
      Left            =   1305
      TabIndex        =   1
      Top             =   270
      Width           =   870
   End
   Begin VB.Frame fra 
      Height          =   1050
      Index           =   1
      Left            =   45
      TabIndex        =   13
      Top             =   720
      Width           =   2265
      Begin VB.CheckBox chkStayOnTop 
         Caption         =   "Stay On Top"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   585
         Width           =   2040
      End
      Begin VB.CheckBox chkLog 
         Caption         =   "Enable Log"
         Height          =   240
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "Defualt (Kahadeva.log)"
         Top             =   225
         Width           =   2085
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Status/Recieved Data"
      Height          =   3780
      Index           =   3
      Left            =   2385
      TabIndex        =   12
      Top             =   45
      Width           =   4290
      Begin RichTextLib.RichTextBox tdata 
         Height          =   3345
         Left            =   45
         TabIndex        =   15
         Top             =   180
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   5900
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"main.frx":0442
      End
   End
   Begin VB.Frame fra 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   2
      Left            =   45
      TabIndex        =   10
      Top             =   45
      Width           =   2265
      Begin VB.TextBox tPort 
         Height          =   285
         Left            =   540
         MaxLength       =   5
         TabIndex        =   0
         Top             =   225
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   270
         Width           =   360
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Kahadeva Port Listener v0.66 BETA
' Copyright © 2000 fezZe
' You may use any portion of this code in your applications as long i´ll get credits


Option Explicit
Dim Data As String, totalBytes As Long, Log As Boolean, First As Boolean

Private Sub cAbout_Click()
about.Show vbModal, Me
End Sub

Private Sub cClear_Click()
tdata.Text = ""
End Sub

Private Sub cCopy_Click()
Clipboard.Clear
Clipboard.SetText tdata.SelText, vbCFText
End Sub

Private Sub cExit_Click()
If tcp.State = 7 Or tcp.State = 2 Then
    If MsgBox("Port listen is active!" & vbCrLf & "Quit anyway?", vbYesNo + vbQuestion, "Quit Kahadeva") = vbYes Then _
    tcp.Close: End
Else
    If MsgBox("Are You Sure?", vbYesNo + vbQuestion, "Quit Kahadeva") = vbYes Then End
End If
End Sub

Private Sub chkLog_Click()
Select Case chkLog.Value
    Case Checked
        Log = True
        SaveSetting "Kahadeva", "PortListen", "Log", "1"
        Open LogFilePath For Append As #1
            Print #1, vbCrLf & "- Log Started: " & Date & " " & Time & " -" & vbCrLf
        Close #1
    Case Else
        Log = False
        SaveSetting "Kahadeva", "PortListen", "Log", "0"
         Open LogFilePath For Append As #1
            If totalBytes = 0 Then Print #1, vbCrLf & "   No Data Recieved"
            Print #1, vbCrLf & "- Log Closed: " & Date & " " & Time & " -" & vbCrLf
        Close #1
End Select
End Sub

Sub StartListen()
On Error GoTo ErrHandler
With tcp
            .Close
            .LocalPort = tPort.Text
            .Listen
          End With
        main.Caption = "Kahadeva Port Listener v0.66 - Listening"
        With tdata
            .SelColor = vbHighlight
            .SelText = "-> Listening on port: ": .SelColor = vbRed: .SelText = tcp.LocalPort & vbCrLf
        End With
        cListen.Caption = "Close"
        If Log = True Then
            Open LogFilePath For Append As #1
            Print #1, vbCrLf & "- Log Started: " & Date & " " & Time & " -" & vbCrLf
            Print #1, "-> Listening to port: " & tcp.LocalPort & vbCrLf
            Close #1
        End If
Exit Sub
ErrHandler:
MsgBox "An error occured!" & vbCrLf & vbCrLf & "Description: " & Err.Description, vbCritical
tcp.Close
End Sub

Private Sub chkStayOnTop_Click()
Select Case chkStayOnTop
    Case Checked
        SaveSetting "Kahadeva", "PortListen", "OnTop", "1"
        TopMost main, True
    Case Unchecked
        SaveSetting "Kahadeva", "PortListen", "OnTop", "0"
        TopMost main, False
End Select
End Sub

Private Sub cListen_Click()
On Error GoTo ErrHandler
Select Case cListen.Caption
    Case "Start"
          With tcp
            .Close
            .LocalPort = tPort.Text
            .Listen
          End With
        main.Caption = "Kahadeva Port Listener v0.66 - Listening"
        SaveSetting "Kahadeva", "PortListen", "Port", tPort.Text
        With tdata
            .SelColor = vbHighlight
            .SelText = "-> Listening on port: ": .SelColor = vbRed: .SelText = tcp.LocalPort & vbCrLf
        End With
        cListen.Caption = "Close"
        If Log = True Then
            Open LogFilePath For Append As #1
            Print #1, vbCrLf & "- Log Started: " & Date & " " & Time & " -" & vbCrLf
            Print #1, "-> Listening to port: " & tcp.LocalPort
            Close #1
        End If
    Case "Close"
            sRet = GetSetting("Kahadeva", "PortListen", "Msg2?")
         If sRet = "1" Then
            If tcp.State <> 7 Then GoTo JustClose
            sRet = GetSetting("Kahadeva", "PortListen", "Msg2Txt")
            tcp.SendData sRet
            With tdata
                .SelColor = vbHighlight
                .SelText = "-> Message Successfully Sent!" & vbCrLf
            End With
            Wait 0.7
            tcp.Close
         Else
            tcp.Close
            main.Caption = "Kahadeva Port Listener v0.66"
        End If
JustClose:
        tcp.Close
        cListen.Caption = "Start"
        main.Caption = "Kahadeva Port Listener v0.66"
        With tdata
            .SelColor = vbHighlight
            .SelText = "-> Listen Halted" & vbCrLf
        End With
    If Log = True Then
            Open LogFilePath For Append As #1
            Print #1, "-> Listen Halted"
            If totalBytes = 0 Then Print #1, vbCrLf & "  No Data Recieved"
            Print #1, vbCrLf & "- Log Closed: " & Date & " " & Time & " -" & vbCrLf
            Close #1
    End If
End Select
ErrHandler:
If Err.Number = 0 Then Exit Sub
MsgBox "An Error occured!" & vbCrLf & vbCrLf & "Description: " & Err.Description, vbCritical
End Sub

Private Sub cmore_Click()
prefs.Show vbModal, Me
End Sub

Private Sub cSelect_Click()
Clipboard.Clear
Clipboard.SetText tdata.Text, vbCFText
End Sub

Private Sub cSendStr_Click()
If tcp.State <> 7 Then Exit Sub
SendStr.Show vbModal, Me
End Sub

Private Sub Form_Activate()
If ShowNotice = True Then Exit Sub
If First Then Notice.Show vbModal, Me
End Sub

Private Sub Form_Load()
Dim Started As Long
totalBytes = 0
sRet = GetSetting("Kahadeva", "PortListen", "Started")
If sRet = "" Or sRet = "0" Then
    First = True
Else
    Started = sRet
End If
sRet = GetSetting("Kahadeva", "PortListen", "Port")
If Len(sRet) < 1 Then tPort.Text = "0" Else tPort.Text = sRet
sRet = GetSetting("Kahadeva", "PortListen", "LogFile")
If sRet = "" Then
    LogFilePath = "Kahadeva.log"
Else
    LogFilePath = sRet
End If
sRet = GetSetting("Kahadeva", "PortListen", "Log")
If sRet = "1" Then
  Log = True
  chkLog.Value = Checked
Else
    Log = False
    chkLog.Value = Unchecked
End If
sRet = GetSetting("Kahadeva", "PortListen", "ClearLog?")
    If sRet = "1" Then Open LogFilePath For Output As #1: Close #1
sRet = GetSetting("Kahadeva", "PortListen", "StartListen?")
    If sRet = "1" Then StartListen
Started = Started + 1
SaveSetting "Kahadeva", "PortListen", "Started", Started
sRet = GetSetting("Kahadeva", "PortListen", "OnTop")
If sRet = "1" Then
    TopMost main, True
    chkStayOnTop.Value = Checked
Else
    TopMost main, False
    chkStayOnTop.Value = Unchecked
End If
End Sub

Private Sub tcp_Close()
tcp.Close
With tdata
    .SelColor = vbRed
    .SelText = vbCrLf & vbCrLf & "-> Connection Closed by Client" & vbCrLf
End With
If Log = True Then
    Open LogFilePath For Append As #1
        Print #1, vbCrLf & "-> Connection Closed by Client : " & Date & " " & Time
    Close #1
End If
cListen.Caption = "Start"
main.Caption = "Kahadeva Port Listener v0.66"
StartListen
End Sub

Private Sub tcp_ConnectionRequest(ByVal requestID As Long)
If tcp.State <> 0 Then tcp.Close
tcp.Accept requestID
lblIP.Caption = tcp.RemoteHostIP
lblHostname.Caption = tcp.RemoteHost
If tcp.RemoteHost = "" Then lblHostname.Caption = "N/A": lblHostname.Enabled = False
With tdata
    .SelColor = vbHighlight
    .SelText = "-> Connection Established!" & vbCrLf & vbCrLf
End With
lblTime.Caption = Time
main.Caption = "Kahadeva Port Listener v0.66 - Client [" & tcp.RemoteHostIP & "]"
If Log = True Then
    Open LogFilePath For Append As #1
    Print #1, vbCrLf & "-> Connection Established: " & Date & " " & Time & " - [" & tcp.RemoteHostIP & "]" & vbCrLf
    Close #1
End If
sRet = GetSetting("Kahadeva", "PortListen", "Msg1?")
    If sRet = "1" Then
        sRet = GetSetting("Kahadeva", "PortListen", "Msg1Txt")
        tcp.SendData sRet
        With tdata
            .SelColor = vbHighlight
            .SelText = "-> Message Successfully Sent!" & vbCrLf & vbCrLf
        End With
    End If
End Sub

Private Sub tcp_DataArrival(ByVal bytesTotal As Long)
Dim TmpPkt As String
tcp.GetData Data, vbString
lblPktSize.Caption = bytesTotal
totalBytes = totalBytes + bytesTotal
lblBytes.Caption = totalBytes
TmpPkt = Mid(Data, 1, 3)
    If TmpPkt = "BN" Then
        With tdata
            .SelColor = vbRed
            .SelText = "-> WARNING: Might be a NetBus Pro/NetBus Clone Client! Check Your System For Trojans!" & vbCrLf
        End With
    Open LogFilePath For Append As #1
        Print #1, tcp.RemoteHostIP & "  :  " & Date & " " & Time & "  :  !! A NetBus Pro/NetBus Clone Client Attempted to Connect !!"
    Close #1: Exit Sub
    End If
With tdata
    .SelColor = vbBlack
    .SelText = Data & vbCrLf
End With
If Log = True Then
    Open LogFilePath For Append As #1
    Print #1, tcp.RemoteHostIP & " : " & Date & " " & Time & "  :   Recieved: " & Data
    Close #1
End If
End Sub



Private Sub tcp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Error in Winsock: " & vbCrLf & vbCrLf & "Description: " & Description
tcp.Close
End Sub

Private Sub tdata_Change()
tdata.SelStart = Len(tdata.Text)
End Sub
