VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmChanger 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8430
   ClientLeft      =   -15
   ClientTop       =   -15
   ClientWidth     =   15510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   15510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15615
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         TabIndex        =   3
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3120
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   4200
         Width           =   2775
      End
      Begin VB.TextBox txtUser 
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
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   1
         Top             =   3600
         Width           =   2775
      End
      Begin MSAdodcLib.Adodc loginado 
         Height          =   375
         Left            =   2760
         Top             =   1320
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\cjrd\Desktop\Clinic\ClinicRecord.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\cjrd\Desktop\Clinic\ClinicRecord.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Select * FROM clinic_master"
         Caption         =   "loginado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Image Command1 
         Height          =   375
         Left            =   6000
         Picture         =   "frmChanger.frx":0000
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Datamex College of Saint Adeline"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   9360
         TabIndex        =   6
         Top             =   480
         Width           =   5295
      End
      Begin VB.Image Image3 
         Height          =   1695
         Left            =   240
         Picture         =   "frmChanger.frx":2B88
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2010
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   8520
         Left            =   0
         Picture         =   "frmChanger.frx":5794
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15600
      End
   End
End
Attribute VB_Name = "frmChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()

If txtUser.Text = "" Or txtPass.Text = "" Then
    MsgBox "Enter new username and password.", vbExclamation
    Exit Sub
End If

If Len(txtUser.Text) < 5 Or Len(txtPass.Text) < 5 Then
    MsgBox "Username and Password must be at least 5 characters long.", vbExclamation
    txtUser.SetFocus
    Exit Sub
End If

If txtUser.Text = "User" Then
    MsgBox "Credentials cannot be the same as default.", vbExclamation
    txtUser.SetFocus
    txtUser.Text = ""
    txtPass.Text = ""
    Exit Sub
End If

If txtUser.Text = "Admin" Then
    MsgBox "Credentials cannot be the same as administrator.", vbExclamation
    txtUser.SetFocus
    txtUser.Text = ""
    txtPass.Text = ""
    Exit Sub
End If

loginado.RecordSource = "SELECT * FROM clinic_master WHERE ID=1"
loginado.Refresh

    If MsgBox("Confirm Changes?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

If Not loginado.Recordset.EOF Then
    loginado.Recordset!UserName = txtUser.Text
    loginado.Recordset!Password = txtPass.Text
    loginado.Recordset!FirstLogin = False
    loginado.Recordset.Update
End If

MsgBox "Credentials successfully changed!", vbInformation

frmLogin.Show
Unload Me

End Sub

Private Sub Command1_Click()

If txtPass.PasswordChar = "*" Then
    txtPass.PasswordChar = ""
    Command1.BorderStyle = 1
Else
    txtPass.PasswordChar = "*"
    Command1.BorderStyle = 0
End If

End Sub

Private Sub FilterInput(KeyAscii As Integer, _
                        ByVal AllowNumbers As Boolean, _
                        ByVal AllowExtraChars As String)

    If KeyAscii = vbKeyBack Then Exit Sub

    If (KeyAscii >= 65 And KeyAscii <= 90) Or _
       (KeyAscii >= 97 And KeyAscii <= 122) Or _
       (KeyAscii = 164 Or KeyAscii = 165) Then Exit Sub
       
    If AllowNumbers Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    End If

    If InStr(AllowExtraChars, Chr(KeyAscii)) > 0 Then Exit Sub
    
    KeyAscii = 0

End Sub

Private Sub txtPass_Validate(Cancel As Boolean)
    If txtPass.Text = "" Then Exit Sub
    If txtPass.Text Like "*[!A-Za-z0-9]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, ""
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, ""
End Sub

Private Sub txtUser_Validate(Cancel As Boolean)
    If txtUser.Text = "" Then Exit Sub
    If txtUser.Text Like "*[!A-Za-z0-9]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub
