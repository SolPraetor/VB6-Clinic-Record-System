VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMedicineInventory 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7095
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17310
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   17310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMedInv 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17295
      Begin VB.Frame fraAddPanel 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   9375
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   7575
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
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
            Left            =   240
            TabIndex        =   15
            Top             =   3960
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
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
            Left            =   240
            TabIndex        =   14
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
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
            Left            =   240
            TabIndex        =   13
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox txtManufacturer 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            Top             =   2640
            Width           =   2535
         End
         Begin VB.TextBox txtQty 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4560
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   11
            Top             =   1920
            Width           =   2535
         End
         Begin VB.TextBox txtMedName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4560
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   10
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox txtMedID 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
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
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Update"
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
            Left            =   240
            TabIndex        =   7
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblManufacturer 
            BackStyle       =   0  'Transparent
            Caption         =   "Manufacturer:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1680
            TabIndex        =   19
            Top             =   2760
            Width           =   2295
         End
         Begin VB.Label lblQty 
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Quantity:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1680
            TabIndex        =   18
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label lblMedName 
            BackStyle       =   0  'Transparent
            Caption         =   "Medicine Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1680
            TabIndex        =   17
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label lblMID 
            BackStyle       =   0  'Transparent
            Caption         =   "Medicine ID:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1680
            TabIndex        =   16
            Top             =   600
            Width           =   2175
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00800000&
            BorderColor     =   &H00800000&
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   4575
            Left            =   120
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdOrder 
         Caption         =   "Input Medicine"
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
         Left            =   14640
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtAlertStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton cmdLoad 
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
         Left            =   12720
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DGMedicine 
         Height          =   5535
         Left            =   7920
         TabIndex        =   1
         Top             =   1200
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   9763
         _Version        =   393216
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Tahoma"
            Size            =   9.75
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
      Begin VB.Line Line1 
         X1              =   7680
         X2              =   7680
         Y1              =   9360
         Y2              =   0
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   8160
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00800000&
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   7920
         Top             =   240
         Width           =   9135
      End
   End
End
Attribute VB_Name = "frmMedicineInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'General
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

'Main Logic
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            App.Path & "\ClinicRecord.mdb"

    txtMedID.Locked = True
    ClearGrid
    LoadData
End Sub

'Helper Codes
Private Sub Clear()
    txtMedID.Text = ""
    txtMedName.Text = ""
    txtManufacturer.Text = ""
    txtQty.Text = ""
    txtAlertStatus.Text = ""
    
End Sub

Private Sub ClearGrid()

    Set rs = New ADODB.Recordset
    rs.Fields.Append "MedID", adInteger
    rs.Fields.Append "MedName", adVarChar, 50
    rs.Fields.Append "Manufacturer", adVarChar, 50
    rs.Fields.Append "StockQty", adInteger
    rs.Fields.Append "AlertStatus", adVarChar, 12
    
    rs.Open
    
    Set DGMedicine.DataSource = rs
End Sub

Public Sub LoadData()
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open "SELECT MedID, MedName, Manufacturer, StockQty, " & _
            "IIF(StockQty <= 10, 'LOW', 'OK') AS AlertStatus " & _
            "FROM medicine_master ORDER BY MedID ASC", _
            cn, adOpenStatic, adLockReadOnly

    Set DGMedicine.DataSource = rs
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

'Function Codes
Private Function Cleaner(ByVal txt As String) As String 'Get rid of Multi Spacing
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

Private Function GetNextMedID() As Long
    Set rs = New ADODB.Recordset

    rs.CursorLocation = adUseClient
    rs.Open "SELECT MAX(MedID) AS MaxID FROM medicine_master", cn, _
              adOpenForwardOnly, adLockReadOnly

    If rs.EOF Or IsNull(rs!MaxID) Then
        GetNextMedID = 1
    Else
        GetNextMedID = rs!MaxID + 1
    End If

    rs.Close
    Set rs = Nothing
End Function

'Navigation Codes
Private Sub cmdAdd_Click()
    frmOrderMedicine.txtMedID.Text = GetNextMedID
    frmOrderMedicine.Show vbModal
    Clear
End Sub

Private Sub cmdOrder_Click()
    If txtMedID.Text = "" Then
        MsgBox "Select a record first.", vbExclamation
        Exit Sub
    End If

    frmOrderMedicine.txtMedID.Text = txtMedID.Text
    frmOrderMedicine.txtMedName.Text = txtMedName.Text
    frmOrderMedicine.txtMedName.Locked = True
    frmOrderMedicine.txtManufacturer.Text = txtManufacturer.Text
    frmOrderMedicine.txtManufacturer.Locked = True
    Clear
    frmOrderMedicine.Show vbModal
End Sub

Private Sub cmdDelete_Click()

    Dim ExpiredQty As Long
    Dim CurrentQty As Long
    Dim userInput As String
    Dim tempVal As Double
    Dim ValidInput As Boolean

    If txtMedID.Text = "" Then
        MsgBox "Select a record first.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtQty.Text) Then Exit Sub
    CurrentQty = CLng(txtQty.Text)

    ValidInput = False

    Do While Not ValidInput

        userInput = InputBox("Enter number of expired medicines to remove:", "Remove Expired Medicine")

        If userInput = "" Then Exit Sub

        If Not IsNumeric(userInput) Then
            MsgBox "Please enter a valid number.", vbExclamation

        Else
            tempVal = Val(userInput)

            If tempVal <= 0 Then
                MsgBox "Quantity must be greater than zero.", vbExclamation

            ElseIf tempVal > 500 Then
                MsgBox "Invalid Quantity.", vbCritical

            ElseIf tempVal > CurrentQty Then
                MsgBox "Expired quantity cannot exceed current stock!", vbCritical

            Else
                ExpiredQty = CLng(tempVal)
                ValidInput = True
            End If

        End If

    Loop

    If MsgBox("Remove " & ExpiredQty & " expired medicine(s) from stock?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    cn.Execute "UPDATE medicine_master SET StockQty = StockQty - " & ExpiredQty & _
               " WHERE MedID = " & CLng(txtMedID.Text)

    MsgBox "Expired medicines removed from stock!", vbInformation

    Clear
    LoadData

End Sub
Private Sub cmdUpdate_Click()
    txtMedName.Text = Trim(Cleaner(txtMedName.Text))
    txtManufacturer.Text = Trim(Cleaner(txtManufacturer.Text))
    
    If txtMedID.Text = "" Then
        MsgBox "Select a record first.", vbExclamation
        Exit Sub
    End If

    If txtMedName.Locked = True Then
        txtMedName.Locked = False
        txtManufacturer.Locked = False
        cmdUpdate.Default = True
        DGMedicine.Enabled = False
        cmdLoad.Enabled = False
        cmdDelete.Enabled = False
        cmdClear.Enabled = False
        cmdClose.Enabled = False
        cmdAdd.Enabled = False
        cmdOrder.Enabled = False
        
        MsgBox "You can now edit the fields. Press Enter/Escape to save/cancel changes. ", vbInformation
        Exit Sub
    End If

    cn.Execute "UPDATE medicine_master SET " & _
               "MedName='" & Replace(txtMedName.Text, "'", "''") & "', " & _
               "Manufacturer='" & Replace(txtManufacturer.Text, "'", "''") & "' " & _
               "WHERE MedID=" & txtMedID.Text

    MsgBox "Record updated successfully!", vbInformation

    txtMedName.Locked = True
    txtManufacturer.Locked = True
    cmdUpdate.Default = False
    DGMedicine.Enabled = True
    cmdLoad.Enabled = True
    cmdDelete.Enabled = True
    cmdClear.Enabled = True
    cmdClose.Enabled = True
    cmdAdd.Enabled = True
    cmdOrder.Enabled = True
    LoadData

End Sub

Private Sub cmdClear_Click()
    Call Clear
    ClearGrid
End Sub

Private Sub cmdLoad_Click()
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Call LoadData
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        
        If txtMedName.Locked = False Then
            
            txtMedName.Text = DGMedicine.Columns(1).Text
            txtManufacturer.Text = DGMedicine.Columns(2).Text
            
            txtMedName.Locked = True
            txtManufacturer.Locked = True
            DGMedicine.Enabled = True
        cmdLoad.Enabled = True
        cmdDelete.Enabled = True
        cmdClear.Enabled = True
        cmdClose.Enabled = True
        cmdAdd.Enabled = True
        cmdOrder.Enabled = True
            
            MsgBox "Edit cancelled.", vbInformation
            
        End If
        
    End If

End Sub

