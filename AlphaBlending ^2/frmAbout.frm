VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblAppCopyright 
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label lblAppVersion 
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblAppTitle 
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Image imgAppIcon 
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Initialize form frmAbout, shows application
    ' information

    Caption = "About " & App.Title
    
    imgAppIcon.Picture = frmMain.Icon
    
    lblAppTitle.Caption = App.Title
    lblAppVersion.Caption = "Version " & App.Major & "." & _
        App.Minor & "." & App.Revision
    lblAppCopyright.Caption = App.LegalCopyright
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'When key Esc down unload window

    If KeyCode = 27 Then Unload Me
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub
