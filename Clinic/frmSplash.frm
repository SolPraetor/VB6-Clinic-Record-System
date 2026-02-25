VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   15885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   15885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFade 
      Interval        =   20
      Left            =   840
      Top             =   240
   End
   Begin VB.Timer tmrLoad 
      Interval        =   50
      Left            =   240
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar pbSplash 
      Align           =   2  'Align Bottom
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   6135
      Width           =   15885
      _ExtentX        =   28019
      _ExtentY        =   2778
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblWelcome 
      Caption         =   "Clinic record system by JEMCtech Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4800
      TabIndex        =   1
      Top             =   2760
      Width           =   6375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowLong Lib "user32" _
Alias "GetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" _
(ByVal hwnd As Long, _
ByVal crKey As Long, _
ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2

Private FadeValue As Integer

Private Sub Form_Load()
    SetWindowLong Me.hwnd, GWL_EXSTYLE, _
        GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED

    FadeValue = 255
    SetLayeredWindowAttributes Me.hwnd, 0, CByte(FadeValue), LWA_ALPHA

    pbSplash.Min = 0
    pbSplash.Max = 100
    pbSplash.Value = 0

    tmrLoad.Interval = 30
    tmrLoad.Enabled = True

    tmrFade.Interval = 15
    tmrFade.Enabled = False

End Sub

Private Sub Label1_Click()

End Sub

Private Sub tmrLoad_Timer()

    If pbSplash.Value < 100 Then
        pbSplash.Value = pbSplash.Value + 2
    Else
        tmrLoad.Enabled = False
        tmrFade.Enabled = True
    End If

End Sub

Private Sub tmrFade_Timer()

    FadeValue = FadeValue - 8

    If FadeValue <= 0 Then
        FadeValue = 0
        SetLayeredWindowAttributes Me.hwnd, 0, 0, LWA_ALPHA
        tmrFade.Enabled = False
        Unload Me
        Exit Sub
    End If

    If FadeValue > 255 Then FadeValue = 255
    If FadeValue < 0 Then FadeValue = 0

    SetLayeredWindowAttributes Me.hwnd, 0, CByte(FadeValue), LWA_ALPHA

End Sub

