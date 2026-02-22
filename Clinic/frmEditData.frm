VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEditData 
   Caption         =   "Edit Data Window"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraControlPanel 
      Height          =   2055
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdInv 
         Caption         =   "Check Inventory"
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
         Left            =   6600
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbID 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   28
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
         Left            =   3240
         TabIndex        =   27
         Top             =   360
         Width           =   1455
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
         Left            =   4800
         TabIndex        =   26
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
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1455
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
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   1455
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
         TabIndex        =   29
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Frame fraEditDataWindow 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   8415
      Begin VB.ComboBox cboSex 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   3120
         Width           =   2175
      End
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
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1200
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
         TabIndex        =   3
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox txtSymptom 
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
         TabIndex        =   5
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
         TabIndex        =   7
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
         MaxLength       =   11
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   1920
         TabIndex        =   32
         Top             =   2160
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   """MM/dd/yyyy"""
         Format          =   143917059
         CurrentDate     =   46073
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
         Left            =   120
         TabIndex        =   16
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
         Left            =   120
         TabIndex        =   15
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
         Left            =   120
         TabIndex        =   14
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
         Left            =   120
         TabIndex        =   13
         Top             =   3840
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
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblSymptom 
         Caption         =   "Symptom"
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
         TabIndex        =   11
         Top             =   4680
         Width           =   1575
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame fraEditPrescription 
      Height          =   6495
      Left            =   0
      TabIndex        =   18
      Top             =   1920
      Width           =   8415
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   615
         Left            =   5040
         TabIndex        =   34
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cboMedicine 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   600
         Width           =   2055
      End
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   1800
         Width           =   3975
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   480
         Width           =   2535
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
Public OpenPrescription As Boolean
Private MedChanged As Boolean

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

'Necessary Codes
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
    txtCondition.Text = ""
    txtContact.Text = ""
    txtComplain.Text = ""
    cboMedicine.ListIndex = -1
    txtTreatment.Text = ""
    txtDiagnosis.Text = ""
End Sub

Private Sub ShowEditRecord()
    If rs.EOF Or rs.BOF Then Exit Sub

    txtName.Text = rs!Name
    txtAddress.Text = rs!Address
    txtAge.Text = IIf(IsNull(rs!age), "", rs!age)
    If Not IsNull(rs!dob) Then
        dtpDOB.Value = rs!dob
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
    
    txtSymptom.Text = rs!Symptom
    txtCondition.Text = rs!Condition
    txtContact.Text = rs!Contact
    txtComplain.Text = rs!Complain
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
    rsP.CursorLocation = adUseClient
    rsP.Open "SELECT Medicine, Diagnosis, Treatment FROM patient_master WHERE ID = " & patientID, _
             cn, adOpenKeyset, adLockReadOnly

    cboMedicine.ListIndex = -1
    MedChanged = False

    If Not rsP.EOF Then
        If Not IsNull(rsP!Medicine) Then
            Dim i As Integer
            For i = 0 To cboMedicine.ListCount - 1
                If cboMedicine.List(i) = rsP!Medicine Then
                    cboMedicine.ListIndex = i
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

'Command Codes
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

    If MsgBox("Are you sure you want to update this patient record?", vbYesNo + vbQuestion) = vbNo Then
        If rs.State = adStateOpen Then
            If rs.EditMode <> adEditNone Then rs.CancelUpdate
        End If
        Exit Sub
    End If
    
    Dim ContactNo As String
    ContactNo = Trim(txtContact.Text)

    If ContactNo <> "" And Not ContactNo Like "09#########" Then
        MsgBox "Invalid ContactNo number format.", vbExclamation
        txtContact.SetFocus
        Exit Sub
    End If

    rs!Name = txtName.Text
    rs!Address = txtAddress.Text

    If Trim(txtAge.Text) <> "" Then
        rs!age = CLng(txtAge.Text)
    Else
        rs!age = Null
    End If

    rs!dob = dtpDOB.Value
    rs!Sex = cboSex.Text
    rs!Symptom = txtSymptom.Text
    rs!Condition = txtCondition.Text
    rs!Contact = txtContact.Text
    rs!Complain = txtComplain.Text
    rs!Treatment = txtTreatment.Text
    rs!Diagnosis = txtDiagnosis.Text

    Dim selectedMed As String
    selectedMed = ""

    If cboMedicine.ListIndex <> -1 Then
        selectedMed = cboMedicine.Text
    End If

    If selectedMed <> "" Then
        If IsNull(rs!Medicine) Or rs!Medicine <> selectedMed Then
            Dim rsStock As New ADODB.Recordset
            rsStock.Open "SELECT * FROM medicine_master WHERE MedName='" & Replace(selectedMed, "'", "''") & "'", _
                         cn, adOpenDynamic, adLockOptimistic
            If Not rsStock.EOF Then
                If rsStock!StockQty <= 0 Then
                    MsgBox "Medicine is OUT OF STOCK!", vbCritical
                    rsStock.Close
                    Set rsStock = Nothing
                    Exit Sub
                End If
                rsStock!StockQty = rsStock!StockQty - 1
                rsStock.Update
            End If
            rsStock.Close
            Set rsStock = Nothing
        End If
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

    Unload Me

    frmUserDB.OpenLog = True
    frmUserDB.Show
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdInv_Click()
    frmMedicineInventory.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
End Sub

'New Codes
Private Sub dtpDOB_Change()
    CalculateAge
End Sub

Private Sub CalculateAge()

    If dtpDOB.Value > Date Then
        dtpDOB.Value = Date
        Exit Sub
    End If

    Dim age As Integer
    age = Year(Date) - Year(dtpDOB.Value)
    If Month(Date) < Month(dtpDOB.Value) _
    Or (Month(Date) = Month(dtpDOB.Value) And Day(Date) < Day(dtpDOB.Value)) Then
        age = age - 1
    End If

    If age < 0 Or age > 121 Then
        Exit Sub
    End If

    txtAge.Text = age

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
