VERSION 5.00
Begin VB.Form SendStr 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Send Text"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2385
      TabIndex        =   2
      Top             =   1800
      Width           =   1005
   End
   Begin VB.CommandButton cSend 
      Caption         =   "&Send"
      Height          =   330
      Left            =   3690
      TabIndex        =   1
      Top             =   1800
      Width           =   1005
   End
   Begin VB.TextBox tStr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4830
   End
End
Attribute VB_Name = "SendStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cCancel_Click()
Unload Me
End Sub

Private Sub cSend_Click()
With main.tcp
    .SendData tStr.Text
End With
With main.tdata
    .SelColor = vbHighlight
    .SelText = "-> Text Successfully Sent!" & vbCrLf
End With
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Send Text - [" & main.tcp.RemoteHostIP & "]"
End Sub
