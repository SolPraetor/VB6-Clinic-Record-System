VERSION 5.00
Begin VB.Form frmOrderMedicine 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtManufacturer 
      Height          =   615
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtQty 
      Height          =   615
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtMedName 
      Height          =   615
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtMedID 
      Height          =   615
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblManufacturer 
      Caption         =   "Manufacturer:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblQty 
      Caption         =   "Input Quantity:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblMedName 
      Caption         =   "Medicine Name:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblMID 
      Caption         =   "Medicine ID:"
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
'General
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

'Main Logic
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            App.Path & "\ClinicRecord.mdb"
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

'Order Codes
Private Sub cmdOrder_Click()
    Dim OrderQty As Long
    Dim CurrentStock As Long
    Dim NewStock As Long
    Dim MedName As String
    Dim Manufacturer As String
    Const MaxStock As Long = 200

    txtMedName.Text = Cleaner(txtMedName.Text)
    txtManufacturer.Text = Cleaner(txtManufacturer.Text)

    If txtQty.Text = "" Or Not IsNumeric(txtQty.Text) Then
        MsgBox "Enter a valid numeric order quantity.", vbExclamation
        Exit Sub
    End If

    OrderQty = CLng(txtQty.Text)
    
    If OrderQty <= 0 Then
        MsgBox "Quantity must be greater than zero.", vbExclamation
        Exit Sub
    End If
    
    If OrderQty > MaxStock Then
        MsgBox "You cannot input more than the current stock limit.", vbExclamation
        Exit Sub
    End If

    If txtMedID.Text = "" Or Not IsNumeric(txtMedID.Text) Then
        MsgBox "Invalid Medicine ID.", vbExclamation
        Exit Sub
    End If

    MedName = Trim(txtMedName.Text)
    Manufacturer = Trim(txtManufacturer.Text)

    If MedName = "" Or Manufacturer = "" Then
        MsgBox "Medicine name and manufacturer are required.", vbExclamation
        Exit Sub
    End If

    If MedName Like "*[!A-Za-z0-9 .'-]*" Or Manufacturer Like "*[!A-Za-z0-9 &.'-]*" Then
        MsgBox "Medicine name and manufacturer contain invalid characters.", vbExclamation
        Exit Sub
    End If

    If MedName Like "*[A-Za-z0-9]*" = False Then
        MsgBox "Medicine name must contain at least one letter.", vbExclamation
        Exit Sub
    End If
    
    If Manufacturer Like "*[A-Za-z0-9]*" = False Then
        MsgBox "Manufacturer must contain at least one letter.", vbExclamation
        Exit Sub
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT StockQty FROM medicine_master WHERE MedID=" & txtMedID.Text, _
            cn, adOpenDynamic, adLockOptimistic

    If rs.EOF Then
    
    Set rs = New ADODB.Recordset

    rs.Open "SELECT * FROM medicine_master WHERE MedName='" & _
                 Replace(MedName, "'", "''") & _
                 "' AND Manufacturer='" & _
                 Replace(Manufacturer, "'", "''") & "'", _
                 cn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        MsgBox "This medicine with the same manufacturer already exists.", vbExclamation
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    rs.Close
    Set rs = Nothing

    If MsgBox("Add new medicine: " & MedName & " ?", _
                vbYesNo + vbQuestion, "Insert New Medicine") = vbNo Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

        If OrderQty > MaxStock Then
            MsgBox "Cannot add more than " & MaxStock & " units to inventory stock.", vbExclamation
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        
        cn.Execute "INSERT INTO medicine_master (MedID, MedName, Manufacturer, StockQty) VALUES (" & _
                   txtMedID.Text & ", '" & _
                   Replace(MedName, "'", "''") & "', '" & _
                   Replace(Manufacturer, "'", "''") & "', " & _
                   OrderQty & ")"
        
        NewStock = OrderQty
        MsgBox "New medicine added successfully!", vbInformation
        
    Else
        
        CurrentStock = rs!StockQty
        NewStock = CurrentStock + OrderQty

        If NewStock > MaxStock Then
            MsgBox "Total stock cannot exceed " & MaxStock & " units.", vbExclamation
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If

        If MsgBox("Confirm order of " & OrderQty & _
                  " units for medicine: " & MedName & " ?", _
                  vbYesNo + vbQuestion, "Update Stock") = vbNo Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        
        rs!StockQty = NewStock
        rs.Update
        MsgBox "Stock updated successfully!", vbInformation
        
    End If

    txtQty.Text = NewStock
    frmOrderMedicine.Hide
    Unload Me
    
End Sub
Private Sub cmdClose_Click()
    frmOrderMedicine.Hide
    Unload Me
End Sub


