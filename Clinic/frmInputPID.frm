VERSION 5.00
Begin VB.Form frmInputPID 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Certificate"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPID 
      Height          =   495
      Left            =   240
      MaxLength       =   4
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblPID 
      Caption         =   " Enter Patient ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmInputPID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'General
Option Explicit

Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Public PIDValue As String

Private Sub Form_Unload(Cancel As Integer)

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If

End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            App.Path & "\ClinicRecord.mdb"
End Sub

Private Sub txtPID_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cmdConfirm_Click()
    If Trim(txtPID.Text) = "" Then
        MsgBox "Please enter Patient ID.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtPID.Text) Then
        MsgBox "Patient ID must contain numbers only.", vbExclamation
        Exit Sub
    End If
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT ID FROM patient_master WHERE ID = " & CLng(txtPID.Text), _
                 cn, adOpenStatic, adLockReadOnly

    If rs.EOF Then
        MsgBox "No Patient ID Found.", vbCritical
        rs.Close
        Set rs = Nothing
        txtPID.SetFocus
        Exit Sub
    End If

    rs.Close
    Set rs = Nothing

    PIDValue = Trim(txtPID.Text)
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    PIDValue = ""
    Unload Me
End Sub


