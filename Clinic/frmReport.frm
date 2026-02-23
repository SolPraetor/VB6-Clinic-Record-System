VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Logs"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   13620
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Log"
      Height          =   735
      Left            =   5880
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdMedicalCert 
      Caption         =   "Generate Medical Certificate"
      Height          =   735
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdLowStock 
      Caption         =   "Stocks"
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdInventory 
      Caption         =   "Full Inventory"
      Height          =   735
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Daily Log"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DGReport 
      Height          =   6015
      Left            =   7800
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10610
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
End
Attribute VB_Name = "frmReport"
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

Private Sub cmdLog_Click()
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient


    rs.Open "SELECT ID, Name, Complain, Diagnosis, Treatment, Medicine " & _
                  "FROM patient_master ORDER BY ID ASC", _
                  cn, adOpenStatic, adLockReadOnly

    If rs.EOF Then
        MsgBox "No records found.", vbInformation
        Exit Sub
    End If

    Set DGReport.DataSource = rs
    MsgBox "Daily Consultation Log Loaded.", vbInformation
    
End Sub

Private Sub cmdInventory_Click()
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open "SELECT MedID, MedName, Manufacturer, StockQty, " & _
                  "IIF(StockQty <= 10, 'LOW', 'OK') AS AlertStatus " & _
                  "FROM medicine_master ORDER BY MedName ASC", _
                  cn, adOpenStatic, adLockReadOnly

    If rs.EOF Then
        MsgBox "No inventory records found.", vbInformation
        Exit Sub
    End If

    Set DGReport.DataSource = rs
    MsgBox "Inventory Report Loaded.", vbInformation

End Sub

Private Sub cmdLowStock_Click()
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open "SELECT MedID, MedName, Manufacturer, StockQty, " & _
                  "IIF(StockQty <= 10, 'LOW', 'OK') AS AlertStatus " & _
                  "FROM medicine_master " & _
                  "WHERE StockQty <= 10 " & _
                  "ORDER BY StockQty ASC", _
                  cn, adOpenStatic, adLockReadOnly

    If rs.EOF Then
        MsgBox "No low stock medicines.", vbInformation
        Exit Sub
    End If

    Set DGReport.DataSource = rs
    MsgBox "Low Stock Report Loaded.", vbInformation

End Sub


Private Sub cmdMedicalCert_Click()
    On Error GoTo PrintError
    Dim PID As String

    PID = InputBox("Enter Patient ID to generate certificate:")
    If PID = "" Then Exit Sub

    Dim rsCert As New ADODB.Recordset
    rsCert.Open "SELECT Name, Diagnosis FROM patient_master WHERE ID = " & PID, _
                cn, adOpenStatic, adLockReadOnly

    If rsCert.EOF Then
        MsgBox "Patient not found.", vbCritical
        rsCert.Close
        Exit Sub
    End If

    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.Print "                 MEDICAL CERTIFICATE"
    Printer.Print ""
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.Print "This is to certify that " & rsCert!Name & _
                  " was examined on " & Format(Date, "mmmm dd, yyyy") & "."
    Printer.Print ""
    Printer.Print "Diagnosis: " & rsCert!Diagnosis
    Printer.Print ""
    Printer.Print "He/She is advised to take proper medication and rest."
    Printer.Print ""
    Printer.Print ""
    Printer.Print "Physician Signature: ___________________________"
    Printer.EndDoc
    
    MsgBox "Medical Certificate Printed Successfully!", vbInformation
    rsCert.Close
    Set rsCert = Nothing
    Exit Sub
    
PrintError:
    If Err.Number = 482 Then
        MsgBox "Printing canceled by user.", vbInformation
    Else
        MsgBox "Printer error: " & Err.Description, vbCritical
    End If
    Err.Clear
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo PrintError
    If rs Is Nothing Then
        MsgBox "No report loaded.", vbExclamation
        Exit Sub
    End If

    If rs.State <> adStateOpen Then
        MsgBox "Report not available.", vbExclamation
        Exit Sub
    End If

    If rs.EOF Then
        MsgBox "Report is empty.", vbExclamation
        Exit Sub
    End If

    If Printers.Count = 0 Then
        MsgBox "No printer installed.", vbCritical
        Exit Sub
    End If

    rs.MoveFirst
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print "CLINIC REPORT"
    Printer.FontBold = False
    Printer.Print "Generated: " & Format(Now, "mmmm dd, yyyy hh:mm AM/PM")
    Printer.Print String(100, "-")
    Printer.Print ""

    Dim field As Integer
    Dim rowText As String

    Do While Not rs.EOF

        rowText = ""

        For field = 0 To rs.Fields.Count - 1
            rowText = rowText & rs.Fields(field).Name & ": " & _
                      rs.Fields(field).Value & "   "
        Next field

        Printer.Print rowText
        Printer.Print String(80, "-")

        rs.MoveNext
    Loop

    Printer.EndDoc

    MsgBox "Printing successful.", vbInformation
    Exit Sub

PrintError:
    If Err.Number = 482 Then
        MsgBox "Printing canceled by user.", vbInformation
    Else
        MsgBox "Printer error: " & Err.Description, vbCritical
    End If
    Err.Clear
End Sub

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

Private Sub cmdExportPDF_Click()
    If rs Is Nothing Then
        MsgBox "No report loaded.", vbExclamation
        Exit Sub
    End If

    If rs.State <> adStateOpen Then
        MsgBox "Report not available.", vbExclamation
        Exit Sub
    End If

    If rs.EOF Then
        MsgBox "Report is empty.", vbExclamation
        Exit Sub
    End If

    Dim prt As Printer
    Dim found As Boolean
    found = False

    For Each prt In Printers
        If prt.DeviceName = "Microsoft Print to PDF" Then
            Set Printer = prt
            found = True
            Exit For
        End If
    Next

    If Not found Then
        MsgBox "Microsoft Print to PDF not found.", vbCritical
        Exit Sub
    End If

    Printer.ScaleMode = vbTwips
    Printer.CurrentX = 1000
    Printer.CurrentY = 1000

    rs.MoveFirst
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print "CLINIC REPORT"
    Printer.FontBold = False
    Printer.Print "Generated: " & Format(Now, "mmmm dd, yyyy hh:mm AM/PM")
    Printer.Print String(100, "-")
    Printer.Print ""

    Dim field As Integer
    Dim rowText As String
    
    Do While Not rs.EOF

        rowText = ""

        For field = 0 To rs.Fields.Count - 1
            If IsNull(rs.Fields(field).Value) Then
                rowText = rowText & ""
            Else
                rowText = rowText & rs.Fields(i).Value
            End If

            rowText = rowText & "    "
        Next field

        Printer.Print rowText
        If Printer.CurrentY > Printer.ScaleHeight - 1000 Then
            Printer.NewPage
            Printer.CurrentX = 1000
            Printer.CurrentY = 1000
        End If

        rs.MoveNext
    Loop

    Printer.EndDoc

    MsgBox "PDF file generated. Please choose save location.", vbInformation

End Sub



