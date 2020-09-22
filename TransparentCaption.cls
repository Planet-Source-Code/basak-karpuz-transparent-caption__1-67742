VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransparentCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'coded by Ramci
'email RaMCi_GeLiyo@HotMaiL.CoM

'i really thought much on this project
'cuz both of the forms must be
'active or inactive together
'so this would cause a infinite loop :s
'also the less codes of line took much to think on 8-)

'it does not matter weather
'ur form is sizable or toolbox or themed
'or including menu or what ever ...

'this can be used by skinners ;)
'u can just skin the TCaption
'thats what in my mind ;D

'there is something bad that is
'u need to InitializeCaption in Initialize
'but one can easyly modify it ...

'just leave me a credit pls :$

'and at last thanks Paul Caton
'for his great subclasser (h)

Implements MISubclass

Private GSubclass As MGSubclass

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)

Private Const HWND_TOPMOST As Long = (-1)
Private Const HWND_NOTOPMOST As Long = (-2)

Private Const LWA_ALPHA As Long = &H2

Private Const SC_CLOSE As Long = 61536

Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_FRAMECHANGED As Long = &H20

Private Const WM_MOVE As Long = 3
Private Const WM_NCACTIVATE As Long = 134
Private Const WM_NCLBUTTONDOWN As Long = 161
Private Const WM_SIZE As Long = 5
Private Const WM_SYSCOMMAND As Long = 274
Private Const WM_WINDOWPOSCHANGED As Long = 71

Private Const WS_BORDER As Long = 8388608
Private Const WS_CAPTION As Long = 12582912
Private Const WS_THICKFRAME As Long = 262144
Private Const WS_VISIBLE As Long = 268435456

Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_TOOLWINDOW As Long = &H80&

Private TCaption_Handle As Long
Private TCaption_Activated As Boolean
Private TCaption_Border_Width As Long
Private TCaption_Border_Height As Long
Private TCaption_ThickFrame_Width As Long
Private TCaption_ThickFrame_Height As Long
Private TCaption_Caption_Height As Long
Private TCaption_ExStyle As Long
Private TCaption_RECT As RECT
Private TCaption_Style As Long
Private TCaption_Transparency As Long

Private TCaption_SysMenu_Handle As Long
Private TCaption_SysMenu_ExStyle As Long

Private Form_Handle As Long
Private Form_Activated As Boolean
Private Form_Loaded As Boolean
Private Form_Parent_Handle As Long
Private Form_Rect As RECT
Private Form_Style As Long
Private Form_ExStyle As Long

Private Sub Class_Initialize()

    If GSubclass Is Nothing Then Set GSubclass = New MGSubclass

End Sub

Private Sub Class_Terminate()

    If Not GSubclass Is Nothing Then Set GSubclass = Nothing

End Sub

Public Property Let FormHandle(hWnd As Long)

    Form_Handle = hWnd

End Property

Public Property Let CaptionTransparency(Percentage As Long)

    TCaption_Transparency = 255 * (1 - (Val(Percentage) / 100))

End Property

Public Sub InitializeCaption()

    Call GSubclass.Attach_Message(Me, Form_Handle, WM_NCACTIVATE)
    Call GSubclass.Attach_Message(Me, Form_Handle, WM_WINDOWPOSCHANGED)

End Sub

Public Sub TerminiateCaption()

    Call GSubclass.Detach_Message(Me, Form_Handle, WM_NCACTIVATE)
    Call GSubclass.Detach_Message(Me, Form_Handle, WM_WINDOWPOSCHANGED)
    Call DestroyTransparentCaption

End Sub

Private Sub CreateTransparentCaption()

    Form_Parent_Handle = GetParent(Form_Handle)
    Form_Style = GetWindowLong(Form_Handle, GWL_STYLE)
    Form_ExStyle = GetWindowLong(Form_Handle, GWL_EXSTYLE)
    Call GetWindowRect(Form_Handle, TCaption_RECT)
    Call SetWindowLong(Form_Handle, GWL_STYLE, WS_VISIBLE)
    Call SetWindowLong(Form_Handle, GWL_EXSTYLE, 0&)
    If Form_Style And WS_CAPTION Then
        TCaption_Caption_Height = GetCaptionHeight(Form_ExStyle And WS_EX_TOOLWINDOW)
    End If
    If Form_Style And WS_BORDER Then
        TCaption_Border_Width = GetBorderWidth
        TCaption_Border_Height = GetBorderHeight
    End If
    If Form_Style And WS_THICKFRAME Then
        TCaption_ThickFrame_Width = GetThickFrameWidth
        TCaption_ThickFrame_Height = GetThickFrameHeight
    End If
    Call GetWindowRect(Form_Handle, Form_Rect)
    TCaption_Handle = CreateWindowEx(0, "#32770", GetTextByHandle(Form_Handle), WS_VISIBLE, TCaption_RECT.Left, TCaption_RECT.Top, TCaption_RECT.Right - TCaption_RECT.Left, TCaption_RECT.Bottom - TCaption_RECT.Top, Form_Parent_Handle, 0, App.hInstance, 0)
    Call CloneIconByHandle(Form_Handle, TCaption_Handle)
    Call SetWindowLong(TCaption_Handle, GWL_STYLE, Form_Style Or WS_VISIBLE)
    If CBool(TCaption_Transparency) Then
        Call SetWindowLong(TCaption_Handle, GWL_EXSTYLE, Form_ExStyle Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(TCaption_Handle, 0, TCaption_Transparency, LWA_ALPHA)
    Else
        Call SetWindowLong(TCaption_Handle, GWL_EXSTYLE, Form_ExStyle)
    End If
    Call SetWindowPos(TCaption_Handle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED Or SWP_NOACTIVATE)
    Call SetWindowPos(Form_Handle, HWND_NOTOPMOST, TCaption_RECT.Left + TCaption_Border_Width + TCaption_ThickFrame_Width, TCaption_RECT.Top + TCaption_Caption_Height + TCaption_Border_Height + TCaption_ThickFrame_Height, TCaption_RECT.Right - TCaption_RECT.Left - 2 * (TCaption_Border_Width + TCaption_ThickFrame_Width), TCaption_RECT.Bottom - TCaption_RECT.Top - TCaption_Caption_Height - 2 * (TCaption_Border_Height + TCaption_ThickFrame_Height), SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOACTIVATE)
    Call GSubclass.Attach_Message(Me, TCaption_Handle, WM_MOVE)
    Call GSubclass.Attach_Message(Me, TCaption_Handle, WM_NCACTIVATE)
    Call GSubclass.Attach_Message(Me, TCaption_Handle, WM_NCLBUTTONDOWN)
    Call GSubclass.Attach_Message(Me, TCaption_Handle, WM_SIZE)
    Call GSubclass.Attach_Message(Me, TCaption_Handle, WM_SYSCOMMAND)
    Call MISubclass_WindowProc(TCaption_Handle, WM_NCACTIVATE, True, 0)

End Sub

Private Sub DestroyTransparentCaption()

    Call SetWindowLong(Form_Handle, GWL_STYLE, Form_Style Or WS_VISIBLE)
    Call SetWindowLong(Form_Handle, GWL_EXSTYLE, Form_ExStyle)
    Call SetWindowPos(Form_Handle, HWND_NOTOPMOST, TCaption_RECT.Left, TCaption_RECT.Top, TCaption_RECT.Right - TCaption_RECT.Left, TCaption_RECT.Bottom - TCaption_RECT.Top, SWP_FRAMECHANGED)
    Call GSubclass.Detach_Message(Me, TCaption_Handle, WM_MOVE)
    Call GSubclass.Detach_Message(Me, TCaption_Handle, WM_NCLBUTTONDOWN)
    Call GSubclass.Detach_Message(Me, TCaption_Handle, WM_NCACTIVATE)
    Call GSubclass.Detach_Message(Me, TCaption_Handle, WM_SIZE)
    Call DestroyWindow(TCaption_Handle)

End Sub

Private Function MISubclass_WindowProc(ByVal lHwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Static Pending As Boolean

    Select Case lHwnd
        Case Form_Handle
            Select Case iMsg
                Case WM_NCACTIVATE
                    Form_Activated = CBool(wParam)
                    'Pending means this WM_* is send form other window
                    If Pending Then
                        'if this WM_* is sent by TCaption
                        'Set top or no top windows respect to their state
                        'i mean top for active and no top for inactive
                        If Form_Activated Then
                            If Not TCaption_Activated Then
                                Call SetWindowPos(TCaption_Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
                            End If
                            Call SetWindowPos(lHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
                        Else
                            If TCaption_Activated Then
                                Call SetWindowPos(TCaption_Handle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
                            End If
                            Call SetWindowPos(lHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
                        End If
                    Else
                        'if this WM_* belongs to Form share with TCaption
                        Pending = True
                        Call MISubclass_WindowProc(TCaption_Handle, WM_NCACTIVATE, wParam, 0)
                        Pending = False
                    End If
                    MISubclass_WindowProc = GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
                Case WM_WINDOWPOSCHANGED
                    'when WM_WINDOWPOSCHANGED passes for the 1st time
                    'form is initialized, it means form is loaded ;)
                    If Not Form_Loaded Then
                        Form_Loaded = True
                        Call CreateTransparentCaption
                    End If
                    MISubclass_WindowProc = GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
                Case Else
                    'WM_SYSCOMMAND with SC_CLOSE will be passed here ...
                    MISubclass_WindowProc = GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
            End Select
        Case TCaption_Handle
            Select Case iMsg
                Case WM_NCACTIVATE
                    TCaption_Activated = CBool(wParam)
                    If Not Pending Then
                        'if this WM_* belong to TCaption share with Form ;D
                        Pending = True
                        Call MISubclass_WindowProc(Form_Handle, WM_NCACTIVATE, wParam, 0)
                        Pending = False
                    End If
                    MISubclass_WindowProc = GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
                Case WM_NCLBUTTONDOWN
                    'This is for fastening resize and move events
                    Call HideWindow(Form_Handle)
                    MISubclass_WindowProc = GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
                    Call UnHideWindow(Form_Handle)
                Case WM_MOVE
                    'Let TCaption move
                    MISubclass_WindowProc = GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
                    Call GetWindowRect(lHwnd, TCaption_RECT)
                    'Move Form
                    Call SetWindowPos(Form_Handle, 0, TCaption_RECT.Left + TCaption_Border_Width + TCaption_ThickFrame_Width, TCaption_RECT.Top + TCaption_Caption_Height + TCaption_Border_Height + TCaption_ThickFrame_Height, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE)
                Case WM_SIZE
                    'Let resize TCaption
                    MISubclass_WindowProc = GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
                    Call GetWindowRect(lHwnd, TCaption_RECT)
                    'Resize Form
                    Call SetWindowPos(Form_Handle, 0, 0, 0, TCaption_RECT.Right - TCaption_RECT.Left - 2 * (TCaption_Border_Width + TCaption_ThickFrame_Width), TCaption_RECT.Bottom - TCaption_RECT.Top - TCaption_Caption_Height - 2 * (TCaption_Border_Height + TCaption_ThickFrame_Height), SWP_NOMOVE Or SWP_NOACTIVATE)
                Case WM_SYSCOMMAND
                    Select Case wParam
                        Case SC_CLOSE
                            'idk why CS_CLOSE does not work
                            'but np send it to Form ;)
                            'it will TerminiateCaption and its self
                            Call MISubclass_WindowProc(Form_Handle, iMsg, wParam, 0)
                        Case Else
                            MISubclass_WindowProc = GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
                    End Select
                Case Else
                    MISubclass_WindowProc = GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
            End Select
    End Select

End Function

Private Property Get MISubclass_MsgResponse() As EMsgResponse

    Select Case GSubclass.p_CurrentMessage
        Case WM_MOVE, WM_NCACTIVATE, WM_NCLBUTTONDOWN, _
            WM_SIZE, WM_SYSCOMMAND, WM_WINDOWPOSCHANGED
            MISubclass_MsgResponse = emrConsume
        Case Else
            MISubclass_MsgResponse = emrPreProcess
    End Select

End Property

Private Property Let MISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property
