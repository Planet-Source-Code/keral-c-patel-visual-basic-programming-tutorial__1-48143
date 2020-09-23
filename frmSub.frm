VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmSub 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSub.frx":0000
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   767
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optOff 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "O&FF"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   540
      TabIndex        =   5
      Top             =   3120
      Width           =   630
   End
   Begin VB.OptionButton optOn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&ON"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   540
      TabIndex        =   4
      Top             =   2790
      Width           =   630
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6045
      Left            =   1785
      TabIndex        =   0
      Top             =   480
      Width           =   9390
      ExtentX         =   16563
      ExtentY         =   10663
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
   Begin VB.Shape Shape2 
      Height          =   450
      Left            =   300
      Top             =   975
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   300
      Top             =   2280
      Width           =   1140
   End
   Begin VB.Label lblEffects 
      BackStyle       =   0  'Transparent
      Caption         =   "Effects:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   450
      TabIndex        =   3
      Top             =   2385
      Width           =   825
   End
   Begin VB.Label lblHead 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   15
      Width           =   11475
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   495
      TabIndex        =   1
      Top             =   990
      Width           =   630
   End
End
Attribute VB_Name = "frmSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    On Error Resume Next
    optOn.BackColor = RGB(230, 190, 210)
    optOff.BackColor = RGB(230, 190, 210)
    WebBrowser1.Navigate strURL

End Sub

Private Sub lblBack_Click()

    On Error Resume Next
    Unload Me
    frmMain.Show
    DoEvents

End Sub

Private Sub lblHead_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    'This makes the form dragable
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

End Sub

Private Sub optOff_Click()

    Dim msgret As Byte
    msgret = MsgBox("It will Effect On Restarting The Application", vbOKOnly, "VB Tutorial")
    'Six because it is return value if YES is pressed
    SaveSetting App.EXEName, "Effects", "Status", "N"

End Sub

Private Sub optOn_Click()

    Dim msgret As Byte
    msgret = MsgBox("It will Effect On Restarting The Application", vbOKOnly, "VB Tutorial")
    'Six because it is return value if YES is pressed
    SaveSetting App.EXEName, "Effects", "Status", "Y"

End Sub

