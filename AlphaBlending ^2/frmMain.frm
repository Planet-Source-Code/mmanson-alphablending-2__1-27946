VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5295
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Output"
      Height          =   4695
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   5055
      Begin VB.CheckBox chkAutorefresh 
         Caption         =   "Autorefresh/repaint on slide"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3480
         Width           =   4575
      End
      Begin MSComctlLib.Slider sldLevel 
         Height          =   2895
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   5106
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   43
         Max             =   255
         TickStyle       =   2
         TickFrequency   =   43
      End
      Begin VB.PictureBox picBlending 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   1920
         ScaleHeight     =   2385
         ScaleWidth      =   2385
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdBlending 
         Caption         =   "load"
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lblEnded 
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   4320
         Width           =   3135
      End
      Begin VB.Label lblStarted 
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   3960
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Ended at:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Started at:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3960
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSourceLoad 
      Caption         =   "load"
      Height          =   300
      Index           =   1
      Left            =   4680
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtSourcePath 
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   2
      Text            =   "Thunder.jpg"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Index           =   1
      Left            =   2760
      ScaleHeight     =   2385
      ScaleWidth      =   2385
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdSourceLoad 
      Caption         =   "load"
      Height          =   300
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtSourcePath 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "Sky.jpg"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Index           =   0
      Left            =   120
      ScaleHeight     =   2385
      ScaleWidth      =   2385
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2415
   End
   Begin VB.Menu mnuApp 
      Caption         =   "&App"
      Begin VB.Menu mnuAppExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About &"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Access to the self written C++ DLL
Private Declare Sub alphaBlend Lib "AlphaBlending.dll" ( _
    ByVal width As Long, _
    ByVal height As Long, _
    ByVal level As Long, _
    ByVal targetHDC As Long, _
    ByVal sourceHDC0 As Long, _
    ByVal sourceHDC1 As Long)

Private Declare Sub GetSystemTime Lib "kernel32" ( _
    lpSystemTime As SYSTEMTIME)
    
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private m_blnAppExit As Boolean


Private Sub Form_Load()
    Dim i As Byte

    m_blnAppExit = True

   'Set caption of form
    Caption = App.Title & " (v" & App.Major & "." & App.Minor & "." & _
        App.Revision & ")"
    
    mnuHelpAbout.Caption = "About &" & App.Title
    
   'Get path of images (application path + imagename)
    For i = 0 To 1 Step 1
        txtSourcePath(i).Text = App.Path & IIf( _
            Right(App.Path, 1) = "\", "", "\") & _
            txtSourcePath(i).Text
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'When key Esc down unload window

    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Abort unloading of form when creating image

    If m_blnAppExit = False Then Cancel = True
End Sub

Private Sub mnuAppExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub txtSourcePath_Change(Index As Integer)
   'Exist image in the defined directory? When yes/no
   ' enable/disable button cmdSourceLoad(Index)

    cmdSourceLoad(Index).Enabled = Dir(txtSourcePath(Index).Text) <> ""
End Sub

Private Sub cmdSourceLoad_Click(Index As Integer)
   'Paint - in textbox txtSourcePath(Index) defined -
   ' picture in picturebox picSource(Index)

    picSource(Index).PaintPicture _
        LoadPicture(txtSourcePath(Index).Text), _
        0, _
        0, _
        picSource(Index).width - Screen.TwipsPerPixelX, _
        picSource(Index).height - 2 * Screen.TwipsPerPixelY
End Sub

Private Sub sldLevel_KeyDown(KeyCode As Integer, Shift As Integer)
   'Bring some animation in this application

    Dim i As Integer, width As Integer, height As Integer
    
    width = picSource(0).width / Screen.TwipsPerPixelX - 2
    height = picSource(0).height / Screen.TwipsPerPixelY - 2
    
    If KeyCode = Asc("A") Then
       'Paint every fifth picture to get an animation
        For i = 0 To 255 Step 5
           'Call the function in the C++ DLL
            alphaBlend _
                width, _
                height, _
                i, _
                picBlending.hDC, _
                picSource(0).hDC, _
                picSource(1).hDC
           'Refresh picBlending so that it display the (re)painted
           ' picture
            picBlending.Refresh
        Next
    End If
End Sub

Private Sub sldLevel_Scroll()
    If Not chkAutorefresh.Value = 0 Then
       'Clear picturebox picBlending
        picBlending.Cls
    
       'Call the function in the C++ DLL
        alphaBlend _
            picSource(0).width / Screen.TwipsPerPixelX - 2, _
            picSource(0).height / Screen.TwipsPerPixelY - 2, _
            sldLevel.Value, _
            picBlending.hDC, _
            picSource(0).hDC, _
            picSource(1).hDC
       'Refresh picBlending so that it display the (re)painted
       ' picture
        picBlending.Refresh
    End If
End Sub

Private Sub cmdBlending_Click()
   'Block unloading of form
    m_blnAppExit = False
    
    Dim systime As SYSTEMTIME
    
   'Clear picturebox picBlending
    picBlending.Cls
    
    GetSystemTime systime
    
    lblStarted.Caption = Format(Now, "hh:mm:ss") & ":" & _
        systime.wMilliseconds
    
   'Call the function in the C++ DLL
    alphaBlend _
        picSource(0).width / Screen.TwipsPerPixelX - 2, _
        picSource(0).height / Screen.TwipsPerPixelY - 2, _
        sldLevel.Value, _
        picBlending.hDC, _
        picSource(0).hDC, _
        picSource(1).hDC
    
    GetSystemTime systime
    
    lblEnded.Caption = Format(Now, "hh:mm:ss") & ":" & _
        systime.wMilliseconds
    
   'Refresh picBlending so that it display the (re)painted
   ' picture
    picBlending.Refresh
    
    DoEvents
    
   'Unblock unloading of form
    m_blnAppExit = True
End Sub
