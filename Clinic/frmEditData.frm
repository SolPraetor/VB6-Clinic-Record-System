VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEditData 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   15420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraControlPanel 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   -120
      TabIndex        =   9
      Top             =   -120
      Width           =   15600
      Begin VB.CommandButton cmdInv 
         Caption         =   "Check Inventory"
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
         Left            =   7680
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cmbID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditData 
         Caption         =   "Edit Data"
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
         Left            =   3960
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdEditPrescription 
         Caption         =   "Edit Medication"
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
         Left            =   5760
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Return"
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
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
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
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Patient Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   9720
         TabIndex        =   33
         Top             =   360
         Width           =   5535
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00800000&
         BorderColor     =   &H00800000&
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   0
         Top             =   0
         Width           =   15600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Patient: None"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4560
         TabIndex        =   27
         Top             =   1440
         Width           =   7335
      End
      Begin VB.Label lblID 
         Caption         =   "Select ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.Frame fraEditPrescription 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   15480
      Begin VB.TextBox txtGivenMed 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   3360
         MaxLength       =   1
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
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
         Left            =   5040
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cboMedicine 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtDiagnosis 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   2160
         Width           =   7335
      End
      Begin VB.TextBox txtTreatment 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   7920
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2160
         Width           =   7335
      End
      Begin VB.Label lblTreatment 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Treatment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   8280
         TabIndex        =   32
         Top             =   1560
         Width           =   6975
      End
      Begin VB.Label lblDiagnosis 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Diagnosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   31
         Top             =   1560
         Width           =   7095
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0000C000&
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   6960
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblMedicine 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Given:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblMedicineG 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Medicine:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraEditDataWindow 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   15480
      Begin VB.TextBox txtContact 
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
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   37
         Top             =   5400
         Width           =   4935
      End
      Begin VB.TextBox txtAge 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox txtAddress 
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   35
         Top             =   3720
         Width           =   4935
      End
      Begin VB.TextBox txtName 
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   34
         Top             =   360
         Width           =   4935
      End
      Begin VB.ComboBox txtSymptom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   29
         Text            =   "txtSymptom"
         Top             =   4680
         Width           =   4935
      End
      Begin VB.ComboBox cboSex 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3000
         Width           =   4935
      End
      Begin VB.TextBox txtComplain 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   9360
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1080
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   2280
         TabIndex        =   22
         Top             =   1320
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   """MM/dd/yyyy"""
         Format          =   52953091
         CurrentDate     =   46073
      End
      Begin VB.TextBox txtDOB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Don't Remove, Put under Date Dropdown"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblComplain 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Patient Explanation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   10320
         TabIndex        =   30
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblAge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblDOB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Birthdate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblSex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblSymptom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Symptom:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label lblContact 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H0000C000&
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   9360
         Top             =   240
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frmEditData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'General
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Public OpenPrescription As Boolean
Private MedChanged As Boolean
Private Stopper As Boolean

'Main Logic
Private Sub Form_Load()
    fraEditDataWindow.Visible = True
    fraEditPrescription.Visible = False
    
    If OpenPrescription Then
        fraEditDataWindow.Visible = False
        fraEditPrescription.Visible = True
    Else
        fraEditDataWindow.Visible = True
    End If

    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ClinicRecord.mdb"

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master ORDER BY ID ASC", cn, adOpenKeyset, adLockOptimistic

    LoadDropdown
    LoadCombo
    LoadSymptoms
    
    cboSex.Clear
    cboSex.AddItem "Male"
    cboSex.AddItem "Female"
    cboSex.ListIndex = 0

    If cmbID.ListCount > 0 Then
        cmbID.ListIndex = 0
        LoadRecord CLng(cmbID.Text)
    End If
    
    dtpDOB.MaxDate = Date
    dtpDOB.MinDate = DateAdd("yyyy", -120, Date)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
End Sub

'Helper Codes
Private Sub LoadSymptoms()
    Dim rsSym As ADODB.Recordset
    Set rsSym = New ADODB.Recordset
    
    rsSym.Open "SELECT Symptom FROM symptom_master ORDER BY Symptom ASC", cn, adOpenStatic, adLockReadOnly
    
    txtSymptom.Clear
    
    Do Until rsSym.EOF
        txtSymptom.AddItem rsSym!Symptom
        rsSym.MoveNext
    Loop
    
    rsSym.Close
    Set rsSym = Nothing
End Sub

Private Sub ShowEditRecord()
    If rs.EOF Or rs.BOF Then Exit Sub

    txtName.Text = rs!Name
    txtAddress.Text = rs!Address
    txtSymptom.Text = rs!Symptom
    txtAge.Text = IIf(IsNull(rs!Age), "", rs!Age)
    If Not IsNull(rs!DOB) Then
        dtpDOB.Value = rs!DOB
    Else
        dtpDOB.Value = Date
    End If
    
    If Not IsNull(rs!Sex) Then
        If rs!Sex = "Male" Then
            cboSex.ListIndex = 0
        ElseIf rs!Sex = "Female" Then
            cboSex.ListIndex = 1
        Else
            cboSex.ListIndex = -1
        End If
    Else
        cboSex.ListIndex = -1
    End If
    
    txtContact.Text = rs!Contact
    txtComplain.Text = rs!Complain
End Sub

Private Sub ShowFrame(fra As Frame)
    fraEditDataWindow.Visible = False
    fraEditPrescription.Visible = False
    fra.Visible = True
End Sub

Private Sub Clear()
    txtName.Text = ""
    txtAddress.Text = ""
    txtAge.Text = ""
    dtpDOB.Value = Date
    cboSex.ListIndex = -1
    txtSymptom.Text = ""
    txtContact.Text = "09"
    txtComplain.Text = ""
    cboMedicine.ListIndex = -1
    txtTreatment.Text = ""
    txtDiagnosis.Text = ""
End Sub

Private Sub LoadCombo()
    Dim rsMed As New ADODB.Recordset
    rsMed.CursorLocation = adUseClient
    rsMed.Open "SELECT MedName FROM medicine_master ORDER BY MedName ASC", cn, adOpenStatic, adLockReadOnly

    cboMedicine.Clear
    Do While Not rsMed.EOF
        cboMedicine.AddItem rsMed!MedName
        rsMed.MoveNext
    Loop
    
    cboMedicine.ListIndex = -1

    rsMed.Close
    Set rsMed = Nothing
End Sub

Private Sub LoadDropdown()
    Dim rsIDs As New ADODB.Recordset
    rsIDs.CursorLocation = adUseClient
    rsIDs.Open "SELECT ID FROM patient_master ORDER BY ID ASC", cn, adOpenStatic, adLockReadOnly

    cmbID.Clear
    Do While Not rsIDs.EOF
        cmbID.AddItem rsIDs!ID
        rsIDs.MoveNext
    Loop

    rsIDs.Close
    Set rsIDs = Nothing
End Sub

Private Sub LoadPrescription(ByVal patientID As Long)
    Dim rsP As New ADODB.Recordset
    Dim found As Boolean
    rsP.CursorLocation = adUseClient
    
    rsP.Open "SELECT Medicine, Diagnosis, Treatment FROM patient_master WHERE ID = " & patientID, _
             cn, adOpenKeyset, adLockReadOnly

    cboMedicine.ListIndex = -1
    MedChanged = False
    Stopper = False
    found = False

    If Not rsP.EOF Then
        
    If Not IsNull(rsP!Medicine) Or Not IsNull(rsP!Diagnosis) Or Not IsNull(rsP!Treatment) Then
        Stopper = True
            
            Dim i As Integer
            For i = 0 To cboMedicine.ListCount - 1
                If cboMedicine.List(i) = rsP!Medicine Then
                    cboMedicine.ListIndex = i
                    found = True
                    Stopper = True
                    Exit For
                End If
            Next i
        End If

        txtDiagnosis.Text = IIf(IsNull(rsP!Diagnosis), "", rsP!Diagnosis)
        txtTreatment.Text = IIf(IsNull(rsP!Treatment), "", rsP!Treatment)
    End If

    rsP.Close
    Set rsP = Nothing

End Sub

Private Sub LoadRecord(ByVal patientID As Long)
    If rs.State = adStateOpen Then
        If rs.EditMode <> adEditNone Then rs.CancelUpdate
        rs.Close
    End If

    rs.Open "SELECT * FROM patient_master WHERE ID = " & patientID, _
            cn, adOpenKeyset, adLockOptimistic

    If Not rs.EOF Then
        ShowEditRecord
        LoadPrescription patientID
    Else
        Clear
    End If
End Sub

Private Sub CalculateAge()

    If dtpDOB.Value > Date Then
        dtpDOB.Value = Date
        Exit Sub
    End If

    Dim Age As Integer
    Age = Year(Date) - Year(dtpDOB.Value)
    If Month(Date) < Month(dtpDOB.Value) _
    Or (Month(Date) = Month(dtpDOB.Value) And Day(Date) < Day(dtpDOB.Value)) Then
        Age = Age - 1
    End If

    If Age < 0 Or Age > 121 Then
        Exit Sub
    End If

    txtAge.Text = Age

End Sub

Public Sub InitializeForm()
    fraEditDataWindow.Visible = True
    fraEditPrescription.Visible = False
    
    If OpenPrescription Then
        fraEditDataWindow.Visible = False
        fraEditPrescription.Visible = True
    Else
        fraEditDataWindow.Visible = True
        fraEditPrescription.Visible = False
    End If
    
End Sub

'Function Codes
Private Function Cleaner(ByVal txt As String) As String
    Dim result As String
    Dim i As Long
    Dim checker As String
    Dim lastSpace As Boolean
    
    For i = 1 To Len(txt)
        checker = Mid$(txt, i, 1)

        If checker = " " Then
            If Not lastSpace Then
                result = result & " "
                lastSpace = True
            End If
        Else
            result = result & checker
            lastSpace = False
        End If
    Next i
    
    Cleaner = result
End Function

Private Function CapitalizeName(ByVal s As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim capitalizeNext As Boolean
    
    result = ""
    capitalizeNext = True
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        
        If capitalizeNext And ch >= "a" And ch <= "z" Then
            ch = UCase$(ch)
            capitalizeNext = False
        ElseIf Not capitalizeNext And ch >= "A" And ch <= "Z" Then
            ch = LCase$(ch)
        ElseIf ch = " " Or ch = "-" Or ch = "'" Or ch = "." Then
            capitalizeNext = True
        Else
            capitalizeNext = False
        End If
        
        result = result & ch
    Next i
    
    CapitalizeName = result
End Function

'Navigation Codes
Private Sub cmdUpdate_Click()
    'Cleaner
    txtName.Text = Cleaner(txtName.Text)
    txtAddress.Text = Cleaner(txtAddress.Text)
    txtComplain.Text = Cleaner(txtComplain.Text)
    txtTreatment.Text = Cleaner(txtTreatment.Text)
    txtDiagnosis.Text = Cleaner(txtDiagnosis.Text)
    
    If rs.EOF Or rs.BOF Then
        MsgBox "No patient record selected.", vbExclamation
        Exit Sub
    End If
    
        If Trim(txtName.Text) = "" Then
        MsgBox "Patient's name cannot be Empty.", vbExclamation
        Exit Sub
    End If
    
        If Trim(txtAddress.Text) = "" Then
        MsgBox "Please enter the Patient Address.", vbExclamation
        Exit Sub
    End If
    
    If rs.State <> adStateOpen Then
        rs.Open "SELECT * FROM patient_master WHERE ID = " & cmbID.Text, _
                cn, adOpenKeyset, adLockOptimistic
    End If

    If rs.EOF Then
        MsgBox "Patient record no longer exists.", vbCritical
        Exit Sub
    End If

    If MsgBox("Are you sure you want to update this patient record?", vbYesNo + vbQuestion) = vbNo Then
        If rs.State = adStateOpen Then
            If rs.EditMode <> adEditNone Then rs.CancelUpdate
        End If
        Exit Sub
    End If
    
    If txtContact.Text = "09" Then
        txtContact.Text = ""
    End If

    If txtContact.Text <> "" Then
        If Not txtContact.Text Like "09#########" Then
            MsgBox "Invalid Contact Details.", vbExclamation
            txtContact.SetFocus
            Exit Sub
        End If
    End If

    rs!Name = Trim(txtName.Text)
    rs!Address = Trim(txtAddress.Text)

    If Trim(txtAge.Text) <> "" Then
        rs!Age = CLng(txtAge.Text)
    Else
        rs!Age = Null
    End If

    rs!DOB = dtpDOB.Value
    rs!Sex = cboSex.Text
    rs!Symptom = Trim(txtSymptom.Text)
    rs!Contact = Trim(txtContact.Text)
    rs!Complain = Trim(txtComplain.Text)
    If Stopper Then
        rs!Treatment = Trim(txtTreatment.Text)
        rs!Diagnosis = Trim(txtDiagnosis.Text)

        If cboMedicine.ListIndex <> -1 Then
            rs!Medicine = cboMedicine.Text
        End If
    Else
    
    MsgBox "This patient has no existing medication. " & _
           "Please create a medication on main window.", vbExclamation
           Exit Sub
End If

    Dim selectedMed As String
    selectedMed = ""
    If cboMedicine.ListIndex <> -1 Then selectedMed = cboMedicine.Text

    If selectedMed <> "" And Trim(txtGivenMed.Text) <> "" Then
        Dim QtyToDeduct As Long
        If Not IsNumeric(txtGivenMed.Text) Then
            MsgBox "Please enter a valid quantity to give.", vbExclamation
            Exit Sub
        End If

        QtyToDeduct = CLng(txtGivenMed.Text)
        If QtyToDeduct <= 0 Then
            MsgBox "Quantity must be greater than zero.", vbExclamation
            Exit Sub
        End If

        Dim rsStock As New ADODB.Recordset
        rsStock.Open "SELECT * FROM medicine_master WHERE MedName='" & Replace(selectedMed, "'", "''") & "'", _
                     cn, adOpenDynamic, adLockOptimistic

        If rsStock.EOF Then
            MsgBox "Medicine not found in stock.", vbCritical
            rsStock.Close
            Set rsStock = Nothing
            Exit Sub
        End If

        If rsStock!StockQty < QtyToDeduct Then
            MsgBox "Not enough stock available!" & vbCrLf & _
                   "Available: " & rsStock!StockQty, vbCritical
            rsStock.Close
            Set rsStock = Nothing
            Exit Sub
        End If

        If MsgBox("Deduct " & QtyToDeduct & " from stock for medicine: " & selectedMed & "?", _
                  vbYesNo + vbQuestion) = vbNo Then
            txtGivenMed.Text = ""
            rsStock.Close
            Set rsStock = Nothing
            MsgBox "Update Cancelled"
            Exit Sub
        End If

        rsStock!StockQty = rsStock!StockQty - QtyToDeduct
        rsStock.Update
        rsStock.Close
        Set rsStock = Nothing

        rs!Medicine = selectedMed
    End If

    rs.Update

    MsgBox "Patient record updated successfully!", vbInformation

    If MsgBox("Continue Editing?", vbYesNo + vbQuestion) = vbNo Then
        Unload Me
        frmUserDB.Show
    End If
End Sub

Private Sub cmbID_Click()
    If cmbID.ListIndex <> -1 Then
        Dim newID As Long
        newID = CLng(cmbID.Text)

        LoadRecord newID
        Label1.Caption = "Selected Patient: " & rs!Name
    End If

End Sub

Private Sub cmdEditData_Click()
    ShowFrame fraEditDataWindow
End Sub

Private Sub cmdEditPrescription_Click()
    ShowFrame fraEditPrescription
End Sub

Private Sub cmdReturn_Click()
    If MsgBox("Return to dashboard?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            If rs.EditMode <> adEditNone Then rs.CancelUpdate
            rs.Close
        End If
        Set rs = Nothing
    End If

    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If

    frmUserDB.OpenLog = True
    frmEditData.Hide
    Unload Me

End Sub

Private Sub cmdInv_Click()
    frmMedicineInventory.Show vbModal
End Sub

Private Sub cboMedicine_Change()
    If cboMedicine.ListIndex <> -1 Then
        MedChanged = True
    Else
        MedChanged = False
    End If
End Sub

Private Sub cmdRefresh_Click()
    LoadCombo
End Sub

Private Sub dtpDOB_Change()
    CalculateAge
End Sub

Private Sub txtContact_Change()
    Dim i As Integer
    Dim temp As String
    
    For i = 1 To Len(txtContact.Text)
        If Mid(txtContact.Text, i, 1) Like "[0-9]" Then
            temp = temp & Mid(txtContact.Text, i, 1)
        End If
    Next i
    
    If txtContact.Text <> temp Then
        txtContact.Text = temp
        txtContact.SelStart = Len(temp)
    End If

End Sub

'Key Ascii Codes
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

Private Sub txtContact_KeyPress(KeyAscii As Integer)
    If txtContact.SelStart < 3 And KeyAscii = vbKeyBack Then
        KeyAscii = 0
    End If

  If KeyAscii = vbKeyBack Then Exit Sub
    If Len(txtContact.Text) = 0 Then
        If Chr(KeyAscii) <> "0" Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

    If Len(txtContact.Text) = 1 Then
        If Chr(KeyAscii) <> "9" Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

End Sub

Private Sub txtDiagnosis_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, " ,.-#'():/"
End Sub

Private Sub txtDiagnosis_Validate(Cancel As Boolean)
    If txtDiagnosis.Text = "" Then Exit Sub
    If txtDiagnosis.Text Like "*[!A-Za-z0-9 ,.\#'():/-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, False, " .-'"
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, " ,.-#"
End Sub

Private Sub txtCondition_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, False, " ,.-()"
End Sub

Private Sub txtSymptom_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, False, " "
    
        If KeyAscii = vbKeyBack Then Exit Sub

    If Len(txtSymptom.Text) >= 15 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtComplain_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, " ,.-#'()"
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    If txtName.Text = "" Then Exit Sub
    If txtName.Text Like "*[!A-Za-zñÑ .'-]*" Then
        MsgBox "Invalid characters in Name.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtAddress_Validate(Cancel As Boolean)
    If txtAddress.Text = "" Then Exit Sub
    If txtAddress.Text Like "*[!A-Za-z0-9 ,.\#-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtSymptom_Validate(Cancel As Boolean)
    If txtSymptom.Text Like "*[!A-Za-zñÑ]*" Then
        MsgBox "Invalid characters in Symptom.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtComplain_Validate(Cancel As Boolean)
    If txtComplain.Text = "" Then Exit Sub
    If txtComplain.Text Like "*[!A-Za-z0-9 ,.\#'()-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtGivenMed_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    
End Sub

Private Sub dtpDOB_GotFocus()
    txtDOB.SetFocus
End Sub

Private Sub txtName_LostFocus()
    txtName.Text = CapitalizeName(txtName.Text)
End Sub

Private Sub txtTreatment_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, " ,.-#'()"
End Sub

Private Sub txtTreatment_Validate(Cancel As Boolean)
    If txtTreatment.Text = "" Then Exit Sub
    If txtTreatment.Text Like "*[!A-Za-z0-9 ,.\#'()-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub
