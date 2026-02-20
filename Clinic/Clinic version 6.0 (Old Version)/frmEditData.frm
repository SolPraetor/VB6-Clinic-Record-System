VERSION 5.00
Begin VB.Form frmEditData 
   Caption         =   "Edit Data Window"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraControlPanel 
      Height          =   2055
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8415
      Begin VB.ComboBox cmbID 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditData 
         Caption         =   "Edit Data"
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
         Left            =   3600
         TabIndex        =   30
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdEditPrescription 
         Caption         =   "Edit Prescription"
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
         Left            =   5280
         TabIndex        =   29
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Return"
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
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
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
         Left            =   1920
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblID 
         Caption         =   "Select ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   32
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Frame fraEditPrescription 
      Height          =   6495
      Left            =   0
      TabIndex        =   20
      Top             =   1920
      Width           =   8415
      Begin VB.TextBox txtDiagnosis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtTreatment 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   4200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox txtMed 
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
         Left            =   3000
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblDiagnosis 
         Caption         =   "Diagnosis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblTreatment 
         Caption         =   "Treatment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   27
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblMedicine 
         Caption         =   "Medicine Given"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame fraEditDataWindow 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   8415
      Begin VB.TextBox txtName 
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
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtAge 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtDOB 
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
         Left            =   1920
         TabIndex        =   3
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtAddress 
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
         Left            =   1920
         TabIndex        =   4
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox txtSex 
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
         Left            =   1920
         TabIndex        =   5
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox txtAllergy 
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
         Left            =   1920
         TabIndex        =   6
         Top             =   4560
         Width           =   2175
      End
      Begin VB.TextBox txtCondition 
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
         Left            =   1920
         TabIndex        =   7
         Top             =   5400
         Width           =   2175
      End
      Begin VB.TextBox txtComplain 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   4200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtContact 
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
         Left            =   6000
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblAge 
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblDOB 
         Caption         =   "DOB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblSex 
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lblAllergy 
         Caption         =   "Allergy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label lblCondition 
         Caption         =   "Condition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label lblComplain 
         Caption         =   "Patient Explanation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   11
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label lblContact 
         Caption         =   "Contact"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmEditData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'General
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Public OpenPrescriptionOnLoad As Boolean

'Necessary Codes
Private Sub Form_Load()
    fraEditDataWindow.Visible = True
    fraEditPrescription.Visible = False
    
    If OpenPrescriptionOnLoad Then
        fraEditDataWindow.Visible = False
        fraEditPrescription.Visible = True
    Else
        fraEditDataWindow.Visible = True
    End If

    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ClinicRecord.mdb"

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master ORDER BY ID ASC", cn, adOpenKeyset, adLockOptimistic

    LoadIDDropdown

    If cmbID.ListCount > 0 Then
        cmbID.ListIndex = 0
        LoadRecordByID cmbID.Text
    End If
End Sub

Private Sub ShowFrame(fra As Frame)
    fraEditDataWindow.Visible = False
    fraEditPrescription.Visible = False
    fra.Visible = True
    
End Sub

Private Sub ClearEditFields()
    txtName.Text = ""
    txtAddress.Text = ""
    txtAge.Text = ""
    txtDOB.Text = ""
    txtSex.Text = ""
    txtAllergy.Text = ""
    txtCondition.Text = ""
    txtContact.Text = ""
    txtComplain.Text = ""
    txtMed.Text = ""
    txtTreatment.Text = ""
    txtDiagnosis.Text = ""
End Sub

Private Sub ShowEditRecord()
    If rs.EOF Or rs.BOF Then Exit Sub

    txtName.Text = rs!Name
    txtAddress.Text = rs!Address
    txtAge.Text = IIf(IsNull(rs!Age), "", rs!Age)
    txtDOB.Text = IIf(IsNull(rs!DOB), "", Format(rs!DOB, "mm/dd/yyyy"))
    txtSex.Text = rs!Sex
    txtAllergy.Text = rs!Allergy
    txtCondition.Text = rs!Condition
    txtContact.Text = rs!Contact
    txtComplain.Text = rs!Complain
End Sub

Private Sub LoadIDDropdown()
    Dim rsIDs As New ADODB.Recordset
    rsIDs.Open "SELECT ID FROM patient_master ORDER BY ID ASC", cn, adOpenStatic, adLockReadOnly

    cmbID.Clear
    Do While Not rsIDs.EOF
        cmbID.AddItem rsIDs!ID
        rsIDs.MoveNext
    Loop

    rsIDs.Close
    Set rsIDs = Nothing
End Sub

Private Sub LoadPrescriptionByID(ByVal patientID As Long)
    Dim rsP As New ADODB.Recordset
    rsP.Open "SELECT Medicine, Diagnosis, Treatment FROM patient_master WHERE ID = " & patientID, _
             cn, adOpenKeyset, adLockReadOnly

    If Not rsP.EOF Then
        txtMed.Text = IIf(IsNull(rsP!Medicine), "", rsP!Medicine)
        txtDiagnosis.Text = IIf(IsNull(rsP!Diagnosis), "", rsP!Diagnosis)
        txtTreatment.Text = IIf(IsNull(rsP!Treatment), "", rsP!Treatment)
    End If

    rsP.Close
    Set rsP = Nothing
End Sub

Private Sub LoadRecordByID(ByVal patientID As Long)
    If rs.State = adStateOpen Then
        If rs.EditMode <> adEditNone Then rs.CancelUpdate
        rs.Close
    End If

    rs.Open "SELECT * FROM patient_master WHERE ID = " & patientID, _
            cn, adOpenKeyset, adLockOptimistic

    If Not rs.EOF Then
        ShowEditRecord
        LoadPrescriptionByID patientID
    Else
        ClearEditFields
    End If
End Sub


'Button Codes
Private Sub cmdUpdate_Click()
    If rs.EOF Or rs.BOF Then
        MsgBox "No patient record selected.", vbExclamation
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

    If Trim(txtAge.Text) <> "" And Not IsNumeric(txtAge.Text) Then
        MsgBox "Age must be numeric.", vbExclamation
        txtAge.SetFocus
        Exit Sub
    End If

    If Trim(txtDOB.Text) <> "" And Not IsDate(txtDOB.Text) Then
        MsgBox "DOB must be a valid date.", vbExclamation
        txtDOB.SetFocus
        Exit Sub
    End If

    If MsgBox("Are you sure you want to update this patient record?", vbYesNo + vbQuestion) = vbNo Then
        If rs.State = adStateOpen Then
            If rs.EditMode <> adEditNone Then rs.CancelUpdate
        End If
        Exit Sub
    End If

    rs!Name = txtName.Text
    rs!Address = txtAddress.Text

    If Trim(txtAge.Text) <> "" Then
        rs!Age = CLng(txtAge.Text)
    Else
        rs!Age = Null
    End If

    If Trim(txtDOB.Text) <> "" Then
        rs!DOB = CDate(txtDOB.Text)
    Else
        rs!DOB = Null
    End If

    rs!Sex = txtSex.Text
    rs!Allergy = txtAllergy.Text
    rs!Condition = txtCondition.Text
    rs!Contact = txtContact.Text
    rs!Complain = txtComplain.Text
    rs!Medicine = txtMed.Text
    rs!Treatment = txtTreatment.Text
    rs!Diagnosis = txtDiagnosis.Text
    rs.Update

    MsgBox "Patient record updated successfully!", vbInformation

    If MsgBox("Continue Editing?", vbYesNo + vbQuestion) = vbNo Then
        Unload Me
    End If
End Sub


Private Sub cmbID_Click()
    If cmbID.ListIndex <> -1 Then
        Dim newID As Long
        newID = CLng(cmbID.Text)

        If rs.State = adStateOpen Then
            If rs.EditMode <> adEditNone Then
                If MsgBox("You have unsaved changes. Cancel changes and switch record?", vbYesNo + vbExclamation) = vbNo Then
                    Exit Sub
                Else
                    rs.CancelUpdate
                End If
            End If
        End If

        SelectedPatientID = newID
        LoadRecordByID SelectedPatientID
    End If
End Sub


Private Sub cmdEditData_Click()
    ShowFrame fraEditDataWindow
End Sub

Private Sub cmdEditPrescription_Click()
    ShowFrame fraEditPrescription
End Sub

Private Sub cmdReturn_Click()
    If MsgBox("Are you sure you want to return to Dashboard?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    frmEditData.Hide
    frmUserDB.Show
End Sub

'Exit Codes
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
End Sub


