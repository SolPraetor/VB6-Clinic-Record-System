VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMedicineInventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   13635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMedInv 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.TextBox txtAlertStatus 
         Height          =   735
         Left            =   7920
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
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
         Left            =   9600
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DGMedicine 
         Height          =   5175
         Left            =   6840
         TabIndex        =   1
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   9128
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraAddPanel 
         Height          =   6495
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6735
         Begin VB.TextBox txtMedID 
            Height          =   615
            Left            =   2400
            TabIndex        =   14
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtMedName 
            Height          =   615
            Left            =   2400
            TabIndex        =   13
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox txtQty 
            Height          =   615
            Left            =   2400
            TabIndex        =   12
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox txtManufacturer 
            Height          =   615
            Left            =   2400
            TabIndex        =   11
            Top             =   3240
            Width           =   1575
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
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
            TabIndex        =   10
            Top             =   240
            Width           =   1215
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
            Left            =   2760
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
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
            Left            =   1440
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
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
            Left            =   4080
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
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
            Left            =   5400
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblMID 
            Caption         =   "Medicine ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label lblMedName 
            Caption         =   "Medicine Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   17
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label lblQty 
            Caption         =   "Stock Quantity"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   16
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label lblManufacturer 
            Caption         =   "Manufacturer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   15
            Top             =   3360
            Width           =   1935
         End
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status"
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
         Left            =   6840
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMedicineInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            App.Path & "\ClinicRecord.mdb"

    txtMedID.Locked = True
    txtMedID.Text = GetNextMedID
End Sub

Private Function GetNextMedID() As Long
    Dim rsID As ADODB.Recordset
    Set rsID = New ADODB.Recordset

    rsID.CursorLocation = adUseClient
    rsID.Open "SELECT MAX(MedID) AS MaxID FROM medicine_master", cn, _
              adOpenForwardOnly, adLockReadOnly

    If rsID.EOF Or IsNull(rsID!MaxID) Then
        GetNextMedID = 1
    Else
        GetNextMedID = rsID!MaxID + 1
    End If

    rsID.Close
    Set rsID = Nothing
End Function

Private Sub LoadData()
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open "SELECT MedID, MedName, Manufacturer, StockQty, " & _
            "IIF(StockQty <= 10, 'LOW', 'OK') AS AlertStatus " & _
            "FROM medicine_master ORDER BY MedID ASC", _
            cn, adOpenStatic, adLockReadOnly

    Set DGMedicine.DataSource = rs
End Sub

Private Sub cmdAdd_Click()
    Dim stockValue As Long
    Dim rsCheck As ADODB.Recordset

    If txtMedName.Text = "" Or txtManufacturer.Text = "" Or txtQty.Text = "" Then
        MsgBox "Please complete all fields.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtQty.Text) Then
        MsgBox "Stock must be numeric.", vbExclamation
        Exit Sub
    End If
    stockValue = CLng(txtQty.Text)

    Set rsCheck = New ADODB.Recordset
    rsCheck.CursorLocation = adUseClient
    rsCheck.Open "SELECT * FROM medicine_master WHERE MedName='" & _
                 Replace(txtMedName.Text, "'", "''") & "'", cn, adOpenStatic, adLockReadOnly

    If Not rsCheck.EOF Then
        MsgBox "Medicine name already exists!", vbExclamation
        rsCheck.Close
        Set rsCheck = Nothing
        Exit Sub
    End If
    rsCheck.Close
    Set rsCheck = Nothing

    If MsgBox("Add this medicine?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    cn.Execute "INSERT INTO medicine_master (MedID, MedName, Manufacturer, StockQty) VALUES (" & _
               GetNextMedID & ", '" & Replace(txtMedName.Text, "'", "''") & "', '" & _
               Replace(txtManufacturer.Text, "'", "''") & "', " & stockValue & ")"

    MsgBox "Medicine added successfully!", vbInformation

    If Not frmUserDB Is Nothing Then
        frmUserDB.LoadMedicineCombo frmUserDB.cboMedicine
    End If

    Call ClearFields
    Call LoadData
End Sub

Private Sub cmdUpdate_Click()
    Dim stockValue As Long

    If txtMedID.Text = "" Then
        MsgBox "Select a record from the grid first.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtQty.Text) Then
        MsgBox "Stock must be numeric.", vbExclamation
        Exit Sub
    End If
    stockValue = CLng(txtQty.Text)

    If MsgBox("Update this medicine?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    cn.Execute "UPDATE medicine_master SET " & _
               "MedName='" & Replace(txtMedName.Text, "'", "''") & "', " & _
               "Manufacturer='" & Replace(txtManufacturer.Text, "'", "''") & "', " & _
               "StockQty=" & stockValue & " " & _
               "WHERE MedID=" & txtMedID.Text

    MsgBox "Medicine updated successfully!", vbInformation

    If Not frmUserDB Is Nothing Then
        frmUserDB.LoadMedicineCombo frmUserDB.cboMedicine
    End If

    Call ClearFields
    Call LoadData
End Sub

Private Sub cmdDelete_Click()
    If txtMedID.Text = "" Then
        MsgBox "Select a record first.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Delete this medicine?", vbYesNo + vbCritical) = vbYes Then
        cn.Execute "DELETE FROM medicine_master WHERE MedID=" & txtMedID.Text
        MsgBox "Deleted successfully!", vbInformation

        If Not frmUserDB Is Nothing Then
            frmUserDB.LoadMedicineCombo frmUserDB.cboMedicine
        End If

        Call ClearFields
        Call LoadData
    End If
End Sub

Private Sub DGMedicine_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If DGMedicine.ApproxCount = 0 Then Exit Sub

    txtMedID.Text = DGMedicine.Columns(0).Text
    txtMedName.Text = DGMedicine.Columns(1).Text
    txtManufacturer.Text = DGMedicine.Columns(2).Text
    txtQty.Text = DGMedicine.Columns(3).Text
    txtAlertStatus.Text = DGMedicine.Columns(4).Text
    If txtAlertStatus.Text = "LOW" Then
        MsgBox "Warning: Stock for this medicine is LOW!", vbExclamation, "Low Stock Alert"
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearFields
End Sub

Private Sub ClearFields()
    txtMedID.Text = GetNextMedID
    txtMedName.Text = ""
    txtManufacturer.Text = ""
    txtQty.Text = ""
End Sub

Private Sub cmdLoad_Click()
    Call LoadData
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
