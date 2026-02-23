VERSION 5.00
Begin VB.Form frmOrderMedicine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Order Window"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order Medicine"
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
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
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
      Left            =   2400
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtManufacturer 
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtQty 
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtMedName 
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtMedID 
      Height          =   615
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   1575
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
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblQty 
      Caption         =   "Order Quantity"
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
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
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
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
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
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmOrderMedicine"
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
End Sub

Private Sub cmdOrder_Click()
    Dim OrderQty As Long
    Dim CurrentStock As Long
    Dim NewStock As Long
    Dim MedName As String
    
    If Not IsNumeric(txtQty.Text) Or txtQty.Text = "" Then
        MsgBox "Enter order quantity.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(txtMedID.Text) Or txtMedID.Text = "" Then
        MsgBox "Invalid Medicine ID.", vbExclamation
        Exit Sub
    End If
    
    OrderQty = CLng(txtQty.Text)
    MedName = Trim(txtMedName.Text)
    
    If OrderQty <= 0 Then
        MsgBox "Order quantity must be greater than zero.", vbExclamation
        Exit Sub
    End If
    
    If MedName = "" Or Trim(txtManufacturer.Text) = "" Then
        MsgBox "Medicine name and manufacturer are required for new records.", vbExclamation
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    rs.Open "SELECT StockQty FROM medicine_master WHERE MedID=" & txtMedID.Text, _
            cn, adOpenForwardOnly, adLockReadOnly

    If rs.EOF Then
        If MsgBox("Medicine does not exist." & vbCrLf & _
                  "Add new medicine: " & MedName & " ?", _
                  vbYesNo + vbQuestion, "Insert New Medicine") = vbNo Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        
        cn.Execute "INSERT INTO medicine_master (MedID, MedName, Manufacturer, StockQty) VALUES (" & _
                   txtMedID.Text & ", '" & _
                   Replace(txtMedName.Text, "'", "''") & "', '" & _
                   Replace(txtManufacturer.Text, "'", "''") & "', " & _
                   OrderQty & ")"
                   
        MsgBox "New medicine added successfully!", vbInformation
        
        NewStock = OrderQty
        
    Else
        CurrentStock = rs!StockQty
        NewStock = CurrentStock + OrderQty
        
        If MsgBox("Confirm order of " & OrderQty & _
                  " units for medicine: " & MedName & " ?", _
                  vbYesNo + vbQuestion, "Update Stock") = vbNo Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        
        cn.Execute "UPDATE medicine_master SET StockQty=" & NewStock & _
                   " WHERE MedID=" & txtMedID.Text
                   
        MsgBox "Stock updated successfully!", vbInformation
        
    End If

    rs.Close
    Set rs = Nothing

    txtQty.Text = NewStock
    
    If Not frmMedicineInventory Is Nothing Then
        frmMedicineInventory.LoadData
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
    If Not frmMedicineInventory Is Nothing Then
        frmMedicineInventory.LoadData
    End If
End Sub
