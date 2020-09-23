VERSION 5.00
Begin VB.Form Notice 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Kahadeva Port Listener v0.66"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4380
   ControlBox      =   0   'False
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
   ScaleHeight     =   1320
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cClose 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   1485
      TabIndex        =   2
      Top             =   900
      Width           =   1410
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This Version is a Beta, so some features are disabled or, to be more specific, not implemented."
      ForeColor       =   &H000000FF&
      Height          =   465
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   315
      Width           =   4290
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Kahadeva Port Listener Version 0.66"
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   1
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4290
   End
End
Attribute VB_Name = "Notice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cClose_Click()
Unload Me
ShowNotice = True
prefs.Show vbModal, main
End Sub
