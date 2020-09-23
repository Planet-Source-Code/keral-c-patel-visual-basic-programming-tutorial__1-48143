VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Visual Basic Tutorial...."
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmMain.frx":11CA
   ScaleHeight     =   6120
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkCoding 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Coding"
      Height          =   855
      Left            =   2400
      Picture         =   "frmMain.frx":6115
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2070
      Width           =   1065
   End
   Begin VB.CheckBox chkVariables 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Variables"
      Height          =   855
      Left            =   2400
      Picture         =   "frmMain.frx":65C8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3150
      Width           =   1065
   End
   Begin VB.CheckBox chkFunctions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Functions"
      Height          =   855
      Left            =   2400
      Picture         =   "frmMain.frx":6A38
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4230
      Width           =   1065
   End
   Begin VB.CheckBox chkGUI 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&GUI"
      Height          =   855
      Left            =   4620
      Picture         =   "frmMain.frx":6E66
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2070
      Width           =   1065
   End
   Begin VB.CheckBox chkRegistry 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Registry"
      Height          =   855
      Left            =   4620
      Picture         =   "frmMain.frx":7293
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3150
      Width           =   1065
   End
   Begin VB.CheckBox chkMisc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Misc"
      Height          =   855
      Left            =   4620
      Picture         =   "frmMain.frx":76FF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4230
      Width           =   1065
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6135
      Left            =   -135
      TabIndex        =   2
      Top             =   1095
      Visible         =   0   'False
      Width           =   8625
      ExtentX         =   15214
      ExtentY         =   10821
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   4800
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   2640
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label lblClose 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   7815
      TabIndex        =   9
      Top             =   105
      Width           =   150
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tutorial For Visual-Basic6.0 By Keral."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   2505
      TabIndex        =   1
      Top             =   345
      Width           =   2745
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim effects As String

Private Sub chkCoding_Click()

    On Error Resume Next

    If chkCoding.Value = 1 Then

        DoEvents
        Me.Hide
        strURL = App.Path & "\Html\Coding.htm"
        frmSub.Show
        chkCoding.Value = 0

    End If

End Sub

Private Sub chkVariables_Click()

    On Error Resume Next

    If chkVariables.Value = 1 Then

        DoEvents
        Me.Hide
        strURL = App.Path & "\Html\Variables.htm"
        frmSub.Show
        chkVariables.Value = 0

    End If

End Sub

Private Sub chkFunctions_Click()

    On Error Resume Next

    If chkFunctions.Value = 1 Then

        DoEvents
        Me.Hide
        strURL = App.Path & "\Html\Functions.htm"
        frmSub.Show
        chkFunctions.Value = 0

    End If

End Sub

Private Sub chkGUI_Click()

    On Error Resume Next

    If chkGUI.Value = 1 Then

        DoEvents
        Me.Hide
        strURL = App.Path & "\Html\GUI.htm"
        frmSub.Show
        chkGUI.Value = 0

    End If

End Sub

Private Sub chkRegistry_Click()

    On Error Resume Next

    If chkRegistry.Value = 1 Then

        DoEvents
        Me.Hide
        strURL = App.Path & "\Html\Registry.htm"
        frmSub.Show
        chkRegistry.Value = 0

    End If

End Sub

Private Sub chkMisc_Click()

    On Error Resume Next

    If chkMisc.Value = 1 Then

        DoEvents
        Me.Hide
        strURL = App.Path & "\Html\Misc.htm"
        frmSub.Show
        chkMisc.Value = 0

    End If

End Sub

Private Sub Form_Load()

    On Error Resume Next
    effects = GetSetting(App.EXEName, "Effects", "Status", "Y")

    If effects = "Y" Then

        'For BackGround and mouse effects
        WebBrowser1.Visible = True
        WebBrowser1.Navigate App.Path & "\Html\kcp.html"
        'for music
        Call MultiMedia
        Timer1.Enabled = True
        'for butterflies
        frmFly1.Show
        frmFly2.Show

    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    'This makes the form dragable
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next
    MMControl1.Command = "Close"

End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

End Sub

Private Sub lblClose_Click()

    On Error Resume Next
    Call formclear

End Sub

Private Sub Timer1_Timer()

    On Error Resume Next
    'This Checks if the Music is Playing or whether it is Stopped?
    'If it is not playing or it has stopped than it will Start it Again.

    If MMControl1.Mode = mciModeStop Then

        Call MultiMedia

    End If

    'Place the butterflies on top so that if lostfocus or gotfocus events occur
    'then also this will take care of it
    SetWindowPos frmFly1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos frmFly2.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub MultiMedia()

    On Error Resume Next

    If MMControl1.Mode = mciModePlay Then MMControl1.Command = "Close"

    DoEvents
    'To randomize the Song
    Call frmSplash.randoming
    'To play the Song************************************************************
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.FileName = resultname 'This we got from scanning the hard-disks in frmSplash
    MMControl1.Command = "Open"
    MMControl1.Command = "Prev"
    MMControl1.Command = "Play"
    '****************************************************************************

End Sub

