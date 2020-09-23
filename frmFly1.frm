VERSION 5.00
Begin VB.Form frmFly1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmFly1.frx":0000
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   405
      Top             =   540
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   540
   End
End
Attribute VB_Name = "frmFly1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim scrwidth As Integer
Dim scrheight As Integer

Private Sub Form_Activate()

    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub Form_Load()

    On Error Resume Next
    Shape frmFly1, RGB(255, 255, 255) 'so that white color will be transperent
    scrwidth = Screen.Width
    scrheight = Screen.Height

End Sub

Private Sub Timer1_Timer()

    If Me.Left >= 0 Then

        Me.Left = Me.Left - 50

    Else

        DoEvents
        Me.Left = scrwidth - 480 ' the width of our form

    End If

End Sub

Private Sub Timer2_Timer()

    If Me.Top >= 0 Then

        Me.Top = Me.Top - 50

    Else

        DoEvents
        Me.Top = scrheight - 435 ' the height of our form

    End If

End Sub

