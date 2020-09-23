VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4035
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4035
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox filenames 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4215
      TabIndex        =   6
      Top             =   5595
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.ListBox just 
      Height          =   300
      Left            =   4215
      TabIndex        =   5
      Top             =   5310
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.FileListBox files 
      Height          =   240
      Left            =   4530
      Pattern         =   "*.mp3"
      System          =   -1  'True
      TabIndex        =   4
      Top             =   5385
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.DirListBox folders2 
      Height          =   360
      Left            =   3645
      TabIndex        =   3
      Top             =   5265
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.DirListBox folders 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   2
      Top             =   5295
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.ListBox LstDrv 
      Height          =   300
      Left            =   3705
      TabIndex        =   1
      Top             =   5595
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3090
      Top             =   5505
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":7128
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2205
      Left            =   300
      TabIndex        =   7
      Top             =   1335
      Width           =   7605
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   645
      TabIndex        =   0
      Top             =   465
      Width           =   75
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'For showing the Read-Me file when first run occurs
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
'For finding the drives
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Dim strname As String

Private Sub Form_Load()

    'Error Handling
    On Error Resume Next
    'This function will get registry setting
    'If you are running this for first time then it will return "XYZ"
    strname = GetSetting(App.EXEName, "Info", "Name", "XYZ")

    If strname = "XYZ" Then 'and this Condition Will Be True

        strname = InputBox("Please Enter Your Name Here", "Visual Basic Tutorial", "XYZ")
        'Here it will save the Information
        SaveSetting App.EXEName, "Info", "Name", strname
        ShellExecute Me.hwnd, vbNullString, App.Path & "\HTML\Read-Me.htm", vbNullString, "C:\", SW_SHOWNORMAL

    End If

    'Now it will Read the Information from the Registry again
    strname = GetSetting(App.EXEName, "Info", "Name", "XYZ")
    lblName = "Welcome " & strname ' and display it on label
    'Now for effects setting
    Dim effects As String
    effects = GetSetting(App.EXEName, "Effects", "Status", "Y")

    If effects = "Y" Then

        Timer1.Enabled = True

    Else

        Load frmMain
        Unload Me 'here we are unloading it becasue we don't need songs collection
        frmMain.Visible = True

    End If

End Sub

Private Sub Timer1_Timer()

    On Error Resume Next
    'First we will disable this timer
    Timer1.Enabled = False
    'Then we will find all the drives
    Call FindDrives
    'Then we will find the mp3's in the directories on that drives
    Dim i As Byte

    For i = 0 To LstDrv.ListCount - 1

        folders.Path = LstDrv.List(i)
        Call filesearch
        DoEvents

    Next

    Load frmMain
    Me.Hide 'We are not unloading this form
    'So that we can again take a random song from the filelist and play it
    'Otherwise we can even store the filenames in a collection but it will make
    'this tutorial more complex and I wanted to keep it Simple
    frmMain.Visible = True

End Sub

Private Sub FindDrives()

    On Error Resume Next
    Dim strSave As String
    Dim ret&
    Dim drv As Byte
    'Create a buffer to store all the drives
    strSave = String$(255, Chr$(0))
    'Get all the drives
    ret& = GetLogicalDriveStrings(255, strSave)
    'Extract the drives from the buffer and print them on the form

    For drv = 1 To 100 'I don't think you will have more then 100 Drives

        If Left$(strSave, InStr(1, strSave, Chr$(0))) = Chr$(0) Then Exit For

        LstDrv.AddItem Left$(strSave, InStr(1, strSave, Chr$(0)) - 1)
        strSave = Right$(strSave, Len(strSave) - InStr(1, strSave, Chr$(0)))

    Next

End Sub

Private Sub filesearch()

    On Error Resume Next
    Dim cntr As Integer, xlist As Integer
    Dim yfolders As Integer, zfiles As Integer
    just.Clear
    just.AddItem folders.Path
    files.Path = folders.Path

    For cntr = 0 To files.ListCount - 1

        filenames.AddItem files.Path & "\" & files.List(cntr)

    Next

    Do Until xlist = just.ListCount

        folders2.Path = just.List(xlist)

        If folders2.ListCount > 0 Then

            For yfolders = 0 To folders2.ListCount - 1

                just.AddItem folders2.List(yfolders)
                files.Path = folders2.List(yfolders)

                For zfiles = 0 To files.ListCount - 1

                    filenames.AddItem files.Path & "\" & files.List(zfiles)

                Next

            Next

        End If

        xlist = xlist + 1

    Loop

End Sub

Public Sub randoming()

    On Error Resume Next
    Dim icntr As Integer
    Randomize
    icntr = Int((Rnd * filenames.ListCount) + 1)
    resultname = filenames.List(icntr) 'resultname is the Global String Variable

End Sub

