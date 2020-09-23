VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Kahadeva Port Listener"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "about.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   900
      Top             =   180
   End
   Begin VB.Frame fra 
      Height          =   60
      Left            =   135
      TabIndex        =   4
      Top             =   810
      Width           =   4200
   End
   Begin VB.CommandButton cClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1665
      TabIndex        =   2
      Top             =   1800
      Width           =   1140
   End
   Begin VB.Label lblWinTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   510
      Left            =   45
      TabIndex        =   6
      Top             =   1215
      Width           =   4335
   End
   Begin VB.Label lblTimes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This Program Has Been Executed xxx Times"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   990
      Width           =   4200
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "about.frx":000C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2000 fezZe"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   450
      Width           =   4425
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.66"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   225
      Width           =   4425
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kahadeva Port Listener"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   4425
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount& Lib "kernel32" ()
Private Sub cClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Started As Long, TxtBit As String
sRet = GetSetting("Kahadeva", "PortListen", "Started")
Started = sRet
If Started = "1" Then
    TxtBit = " Time"
Else
    TxtBit = " Times"
End If
lblTimes.Caption = "Kahadeva Has Been Executed " & Started & TxtBit
TickCount
End Sub

Private Sub tmr_Timer()
TickCount
End Sub

Sub TickCount()
Dim h%, m%, ret&
    ret& = GetTickCount&
    h = Int(ret / 3600000)
    m = Int(ret / 60000) - (h * 60)
    lblWinTime.Caption = "Windows has been running for " & vbCr & Str$(h) & " Hours and" & Str$(m) & " Minutes."
End Sub
