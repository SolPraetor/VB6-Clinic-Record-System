VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   15600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   9255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15615
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
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
         Left            =   3720
         TabIndex        =   8
         Top             =   6720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdTerminate 
         Caption         =   "Terminate"
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
         Left            =   3720
         TabIndex        =   7
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
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
         Left            =   3720
         TabIndex        =   6
         Top             =   5280
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc loginado 
         Height          =   375
         Left            =   0
         Top             =   0
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
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3120
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox txtUser 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   0
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   1695
         Left            =   240
         Picture         =   "frmLogin.frx":BDDC2
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2010
      End
      Begin VB.Label Label3 
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
         TabIndex        =   5
         Top             =   480
         Width           =   5295
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
         Left            =   1680
         TabIndex        =   4
         Top             =   3840
         Width           =   1575
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
         Left            =   1680
         TabIndex        =   3
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Image Command1 
         Height          =   375
         Left            =   6000
         Picture         =   "frmLogin.frx":C09CE
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   8520
         Left            =   0
         Picture         =   "frmLogin.frx":C3556
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15600
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogin_Click()
loginado.RecordSource = "SELECT * FROM clinic_master WHERE Username='" & Replace(txtUser.Text, "'", "''") & "'"
loginado.Refresh

If Not loginado.Recordset.EOF Then

    If StrComp(loginado.Recordset!UserName, txtUser.Text, vbBinaryCompare) = 0 _
    And StrComp(loginado.Recordset!Password, txtPass.Text, vbBinaryCompare) = 0 _
    And loginado.Recordset!FirstLogin = True Then
        MsgBox "First login detected! You must change the default username and password.", vbInformation
        frmChanger.Show
        Unload Me
        Exit Sub
    End If

    If StrComp(loginado.Recordset!UserName, txtUser.Text, vbBinaryCompare) = 0 _
    And StrComp(loginado.Recordset!Password, txtPass.Text, vbBinaryCompare) = 0 Then

        If loginado.Recordset!ID = 1 Then
            frmSplash.Show
            Unload Me
        ElseIf loginado.Recordset!ID = 2 Then
            cmdReset.Visible = True
        End If

    Else
        MsgBox "Invalid Credentials!", vbCritical
        txtPass.Text = ""
        txtPass.SetFocus
    End If

Else
    MsgBox "Invalid Credentials!", vbCritical
    txtPass.Text = ""
    txtPass.SetFocus
End If

End Sub


Private Sub cmdTerminate_Click()
    End
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

Private Sub cmdReset_Click()

Dim ans As Integer

ans = MsgBox("Reset credentials to default? (User / 1234)", vbYesNo + vbQuestion)

If ans = vbYes Then

    loginado.RecordSource = "SELECT * FROM clinic_master WHERE ID=1"
    loginado.Refresh

    If Not loginado.Recordset.EOF Then
        loginado.Recordset!UserName = "User"
        loginado.Recordset!Password = "1234"
        loginado.Recordset!FirstLogin = True
        loginado.Recordset.Update
    End If

    MsgBox "Credentials reset to default!", vbInformation
    txtUser.Text = ""
    txtPass.Text = ""
    cmdReset.Visible = False
End If

End Sub

