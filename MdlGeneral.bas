Attribute VB_Name = "MdlGeneral"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Const ICON_SMALL As Long = 0

Private Const WM_GETICON As Long = 127
Private Const WM_GETTEXT As Long = &HD
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_SETICON As Long = 128

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CYCAPTION As Long = 4 'Height of windows caption
Private Const SM_CXBORDER As Long = 5 'Width of no-sizable borders
Private Const SM_CYBORDER As Long = 6 'Height of non-sizable borders
Private Const SM_CXDLGFRAME As Long = 7 'Width of dialog box borders
Private Const SM_CYDLGFRAME As Long = 8 'Height of dialog box borders
Private Const SM_CYSMCAPTION As Long = 51

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1

Public Function GetTextByHandle(hWnd As Long) As String

    Dim TextLenght As Long

    TextLenght = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
    GetTextByHandle = Space(TextLenght)
    Call SendMessage(hWnd, WM_GETTEXT, TextLenght + 1, ByVal GetTextByHandle)

End Function

Public Function GetCaptionHeight(ToolWindow As Boolean) As Long

    If Not ToolWindow Then
        GetCaptionHeight = GetSystemMetrics(SM_CYCAPTION)
    Else
        GetCaptionHeight = GetSystemMetrics(SM_CYSMCAPTION)
    End If

End Function

Public Function GetBorderWidth() As Long

    GetBorderWidth = GetSystemMetrics(SM_CXBORDER)

End Function

Public Function GetBorderHeight() As Long

    GetBorderHeight = GetSystemMetrics(SM_CYBORDER)

End Function

Public Function GetThickFrameWidth() As Long

    GetThickFrameWidth = GetSystemMetrics(SM_CXDLGFRAME)

End Function

Public Function GetThickFrameHeight() As Long

    GetThickFrameHeight = GetSystemMetrics(SM_CYDLGFRAME)

End Function

Public Sub CloneIconByHandle(hWnd1 As Long, hWnd2 As Long)

    Dim IconhWnd As Long

    IconhWnd = SendMessage(hWnd1, WM_GETICON, ICON_SMALL, ByVal 0&)
    Call SendMessage(hWnd2, WM_SETICON, ICON_SMALL, ByVal IconhWnd)

End Sub

Public Sub HideWindow(hWnd As Long)

    Call ShowWindow(hWnd, SW_HIDE)

End Sub

Public Sub UnHideWindow(hWnd As Long)

    Call ShowWindow(hWnd, SW_SHOWNORMAL)

End Sub
