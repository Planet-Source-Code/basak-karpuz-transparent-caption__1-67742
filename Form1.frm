VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   1620
   ClientLeft      =   270
   ClientTop       =   645
   ClientWidth     =   2130
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   2130
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   225
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   270
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000000&
      Caption         =   "MsGBox"
      Height          =   510
      Left            =   630
      TabIndex        =   0
      Top             =   810
      Width           =   780
   End
   Begin VB.Menu m1 
      Caption         =   "Menu1"
      Begin VB.Menu m11 
         Caption         =   "Menu11"
      End
      Begin VB.Menu m12 
         Caption         =   "Menu12"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TransCap As TransparentCaption

Private Sub Command1_Click()

    Call MsgBox("Clicked", vbInformation, "TransparentCaption")

End Sub

Private Sub Form_Initialize()

    Set TransCap = New TransparentCaption
    Let TransCap.FormHandle = hWnd
    Let TransCap.CaptionTransparency = 25
    Call TransCap.InitializeCaption

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call TransCap.TerminiateCaption

End Sub

Private Sub Form_Terminate()

    Set TransCap = Nothing

End Sub

