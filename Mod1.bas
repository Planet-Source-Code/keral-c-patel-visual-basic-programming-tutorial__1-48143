Attribute VB_Name = "Mod1"
Option Explicit
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Const RGN_DIFF = 4
Dim CurRgn As Long, TempRgn As Long
Public strURL As String
Public resultname As String
'for topmost window
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Declare Sub SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Sub formclear()

    On Error Resume Next
    Unload frmMain
    Unload frmSplash
    Unload frmSub
    Unload frmFly1
    Unload frmFly2

End Sub

Public Function Shape(BG As Form, TransperentColor)

    'This is for Butterflies
    On Error Resume Next
    Dim X As Integer, Y As Integer
    Dim success As Boolean
    CurRgn = CreateRectRgn(0, 0, BG.ScaleWidth, BG.ScaleHeight)

    While Y <= BG.ScaleHeight

        While X <= BG.ScaleWidth

            If GetPixel(BG.hdc, X, Y) = TransperentColor Then

                TempRgn = CreateRectRgn(X, Y, X + 1, Y + 1)
                success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
                DeleteObject (TempRgn)

            End If

            X = X + 1

        Wend

        Y = Y + 1
        X = 0

    Wend

    success = SetWindowRgn(BG.hwnd, CurRgn, True)
    DeleteObject (CurRgn)

End Function

