VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUserDB 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10515
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   18225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   18225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Form1"
      Height          =   10695
      Left            =   2640
      TabIndex        =   7
      Top             =   -120
      Width           =   15735
      Begin VB.Frame fraPrescription 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   10575
         Left            =   -120
         TabIndex        =   41
         Top             =   120
         Width           =   15735
         Begin VB.TextBox txtGivenMed 
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
            Left            =   2880
            MaxLength       =   1
            TabIndex        =   67
            Top             =   2640
            Width           =   2655
         End
         Begin VB.TextBox txtPID 
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
            Left            =   2880
            MaxLength       =   5
            TabIndex        =   66
            Top             =   1200
            Width           =   2655
         End
         Begin VB.TextBox txtDiagnosis 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6135
            Left            =   360
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            Top             =   4080
            Width           =   7380
         End
         Begin VB.TextBox txtTreatment 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6135
            Left            =   8040
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            Top             =   4080
            Width           =   7335
         End
         Begin VB.CommandButton cmdPConfirm 
            Caption         =   "Confirm"
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
            Left            =   5760
            TabIndex        =   46
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdPEdit 
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
            Left            =   12840
            TabIndex        =   45
            Top             =   1920
            Width           =   2535
         End
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
            Left            =   12840
            TabIndex        =   44
            Top             =   1080
            Width           =   2535
         End
         Begin VB.ComboBox cboMedicine 
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
            Left            =   2880
            TabIndex        =   43
            Text            =   "cboMedicine"
            Top             =   1920
            Width           =   2655
         End
         Begin VB.CommandButton cmdMedRefresh 
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
            Height          =   495
            Left            =   5760
            TabIndex        =   42
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label lblSelectedSymptom 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Symptom: None"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   855
            Left            =   7440
            TabIndex        =   56
            Top             =   2640
            Width           =   5055
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Medicine Details"
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
            Left            =   0
            TabIndex        =   55
            Top             =   120
            Width           =   3735
         End
         Begin VB.Label lblPID 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Patient ID:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   54
            Top             =   1320
            Width           =   2535
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
            Left            =   360
            TabIndex        =   53
            Top             =   3480
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
            Left            =   7680
            TabIndex        =   52
            Top             =   3480
            Width           =   7695
         End
         Begin VB.Label lblMedicineG 
            BackStyle       =   0  'Transparent
            Caption         =   "Medicine Given:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   51
            Top             =   2640
            Width           =   2535
         End
         Begin VB.Label lblMedicine 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Medicine:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   50
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Label lblSelectedID 
            BackColor       =   &H00C0C0C0&
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
            ForeColor       =   &H00000080&
            Height          =   855
            Left            =   7440
            TabIndex        =   49
            Top             =   1320
            Width           =   5175
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00800000&
            BorderColor     =   &H00C00000&
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   15735
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   10695
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   15735
         Begin VB.CommandButton cmdLog 
            Caption         =   "Consultation Log"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton cmdInventory 
            Caption         =   "Full Inventory"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton cmdLowStock 
            Caption         =   "Low Stocks"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton cmdMedicalCert 
            Caption         =   "Generate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7680
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Generate Medical Certificate"
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print Log"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   1320
            Width           =   2295
         End
         Begin MSDataGridLib.DataGrid DGReport 
            Height          =   8175
            Left            =   360
            TabIndex        =   63
            Top             =   2280
            Width           =   15015
            _ExtentX        =   26485
            _ExtentY        =   14420
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   22
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
               Size            =   12
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Clinic Report"
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
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   2655
         End
         Begin VB.Shape Shape17 
            BorderColor     =   &H00800000&
            Height          =   855
            Left            =   360
            Top             =   1200
            Width           =   12135
         End
         Begin VB.Shape Shape14 
            BackColor       =   &H00800000&
            BorderColor     =   &H00800000&
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   0
            Top             =   120
            Width           =   15735
         End
      End
      Begin VB.Frame fraRemovePatient 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   10695
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   15615
         Begin VB.CommandButton cmdRConfirm 
            Caption         =   "Confirm"
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
            Left            =   3960
            TabIndex        =   18
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtRID 
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
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   17
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CommandButton cmdRefreshArchive 
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
            Height          =   735
            Left            =   4680
            TabIndex        =   16
            Top             =   2040
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid DGArchive 
            Height          =   7575
            Left            =   240
            TabIndex        =   19
            Top             =   2880
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   13361
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   22
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
               Size            =   12
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
         Begin VB.Label Label5 
            BackColor       =   &H0000C000&
            Height          =   135
            Left            =   240
            TabIndex        =   74
            Top             =   2040
            Width           =   4335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Archive Patient Data"
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
            Left            =   -120
            TabIndex        =   23
            Top             =   240
            Width           =   4455
         End
         Begin VB.Label lblRID 
            BackStyle       =   0  'Transparent
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
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H0000C000&
            Caption         =   "Patient Archive Logs"
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
            Height          =   615
            Left            =   240
            TabIndex        =   21
            Top             =   2160
            Width           =   4335
         End
         Begin VB.Label lblSelectedPatientRID 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Selected Patient: None"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Left            =   5520
            TabIndex        =   20
            Top             =   1200
            Width           =   7815
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00800000&
            BorderColor     =   &H00800000&
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   0
            Top             =   120
            Width           =   15615
         End
      End
      Begin VB.Frame fraPatientLog 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   10695
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   15615
         Begin VB.TextBox Text5 
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
            Left            =   240
            MaxLength       =   50
            TabIndex        =   68
            Top             =   1440
            Width           =   4815
         End
         Begin VB.CommandButton Command9 
            Caption         =   "&Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5160
            TabIndex        =   12
            Top             =   1440
            Width           =   975
         End
         Begin VB.CommandButton cmdEdit 
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
            Height          =   735
            Left            =   12000
            TabIndex        =   11
            Top             =   1080
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
            Height          =   735
            Left            =   13800
            TabIndex        =   10
            Top             =   1080
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid DGData 
            Height          =   8415
            Left            =   240
            TabIndex        =   13
            Top             =   2040
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   14843
            _Version        =   393216
            BackColor       =   16777215
            HeadLines       =   1
            RowHeight       =   22
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
               Size            =   12
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
         Begin VB.Label lblLogs 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Data Logs"
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
            Left            =   -120
            TabIndex        =   14
            Top             =   240
            Width           =   3975
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00800000&
            BorderColor     =   &H00800000&
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   0
            Top             =   120
            Width           =   15615
         End
      End
      Begin VB.Frame fraAddPatient 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   10695
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   15735
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
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   1080
            Width           =   1935
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
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   72
            Top             =   2640
            Width           =   7095
         End
         Begin VB.TextBox txtName 
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
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   71
            Top             =   1920
            Width           =   3975
         End
         Begin VB.TextBox txtID 
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
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   70
            Top             =   1200
            Width           =   1695
         End
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
            Left            =   10560
            MaxLength       =   11
            TabIndex        =   69
            Top             =   2640
            Width           =   2535
         End
         Begin VB.ComboBox txtSymptom 
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
            ItemData        =   "frmUserDB.frx":0000
            Left            =   10560
            List            =   "frmUserDB.frx":0002
            TabIndex        =   29
            Text            =   "txtSymptom"
            Top             =   1920
            Width           =   2535
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
            Left            =   10560
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1200
            Width           =   2535
         End
         Begin VB.CommandButton cmdAddConfirm 
            Caption         =   "Confirm"
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
            Left            =   13680
            TabIndex        =   27
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txtComplain 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6135
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   4320
            Width           =   15015
         End
         Begin MSComCtl2.DTPicker dtpDOB 
            Height          =   495
            Left            =   6720
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
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
            Format          =   143917059
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
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "Don't Remove, Put under Date Dropdown"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblContact 
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
            Height          =   615
            Left            =   9000
            TabIndex        =   40
            Top             =   2640
            Width           =   1575
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
            Left            =   240
            TabIndex        =   39
            Top             =   3600
            Width           =   4095
         End
         Begin VB.Label lblSymptom 
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
            Height          =   495
            Left            =   9000
            TabIndex        =   38
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblSex 
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
            Height          =   375
            Left            =   9000
            TabIndex        =   37
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblAddress 
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
            Height          =   495
            Left            =   240
            TabIndex        =   36
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label lblDOB 
            BackStyle       =   0  'Transparent
            Caption         =   "DOB:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5760
            TabIndex        =   35
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lblAge 
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
            Height          =   495
            Left            =   5760
            TabIndex        =   34
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblName 
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
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lblID 
            BackStyle       =   0  'Transparent
            Caption         =   "ID:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblAddPatient 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Add Patient Details"
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
            Left            =   -120
            TabIndex        =   31
            Top             =   240
            Width           =   4215
         End
         Begin VB.Shape Shape8 
            BackColor       =   &H00800000&
            BorderColor     =   &H00800000&
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   0
            Top             =   120
            Width           =   15735
         End
         Begin VB.Shape Shape9 
            BackColor       =   &H0000C000&
            BorderColor     =   &H0000C000&
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   240
            Top             =   3480
            Width           =   4095
         End
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H00000080&
         BorderColor     =   &H00000080&
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   0
         Top             =   120
         Width           =   15735
      End
   End
   Begin VB.Frame fraUserControlPanel 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000A&
      Height          =   10695
      Left            =   0
      TabIndex        =   6
      Top             =   -120
      Width           =   2655
      Begin VB.Timer timDisplay 
         Interval        =   1000
         Left            =   0
         Top             =   10200
      End
      Begin VB.CommandButton cmdReport 
         Appearance      =   0  'Flat
         Caption         =   "Create Clinic Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6360
         Width           =   2295
      End
      Begin VB.CommandButton cmdPatientLog 
         Caption         =   "View Patient Log"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4680
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddPatient 
         Caption         =   "Add Patient"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CommandButton cmdPrescription 
         Caption         =   "Administer Medicine"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3840
         Width           =   2295
      End
      Begin VB.CommandButton cmdRemovePatient 
         Caption         =   "Archive Patient Log"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5520
         Width           =   2295
      End
      Begin VB.CommandButton cmdLogout 
         Caption         =   "Logout"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   0
         Top             =   8160
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   2295
         Left            =   120
         Picture         =   "frmUserDB.frx":0004
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00 PM"
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
         Left            =   480
         TabIndex        =   64
         Top             =   9840
         Width           =   1815
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00000080&
         BorderColor     =   &H00000080&
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   2640
         Top             =   120
         Width           =   12975
      End
      Begin VB.Label lblCpanelDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "mm/dd/yy"
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
         Left            =   600
         TabIndex        =   8
         Top             =   9360
         Width           =   1575
      End
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00800000&
      BorderColor     =   &H00800000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmUserDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'General
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim SelectID As Long
Public OpenLog As Boolean
Dim trigger As CommandButton
Dim ReportType As String

Private Sub Command9_Click()
    Dim searchValue As String
    Dim filterStr As String
    
    searchValue = Trim(Text5.Text)

    If rs Is Nothing Then
        MsgBox "Record not loaded", vbInformation
        Exit Sub
    End If

    If searchValue = "" Then
        rs.Filter = ""
        Set DGData.DataSource = rs
        DGData.Refresh
        Exit Sub
    End If

    filterStr = "Name LIKE '%" & searchValue & "%' OR Sex LIKE '" & searchValue & "%'"
    
    If IsNumeric(searchValue) Then
        filterStr = filterStr & " OR ID = " & CLng(searchValue) & " OR Age = " & CLng(searchValue)
        
    End If

    rs.Filter = filterStr

    If rs.EOF Then
        MsgBox "No record found!", vbInformation
        rs.Filter = ""
        Exit Sub
    End If

    Set DGData.DataSource = rs
    DGData.Refresh
End Sub

'Main Logic
Private Sub Form_Load()
    fraAddPatient.Visible = False
    fraRemovePatient.Visible = False
    fraPrescription.Visible = False
    fraPatientLog.Visible = False
    Frame1.Visible = False
    cmdPConfirm.Enabled = False

    lblCpanelDate.Caption = Date

    If OpenLog Then
        fraPatientLog.Visible = True
    Else
        fraMain.Visible = True
    End If

    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ClinicRecord.mdb"

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master", cn, adOpenDynamic, adLockOptimistic

    LoadCombo cboMedicine
    LoadSymptoms
    
    cboSex.Clear
    cboSex.AddItem "Male"
    cboSex.AddItem "Female"
    cboSex.ListIndex = 0

    dtpDOB.MaxDate = Date
    dtpDOB.MinDate = DateAdd("yyyy", -120, Date)
        
    Cleanup
    
End Sub

'Helper Codes
Public Sub Cleanup()
    On Error Resume Next
    cn.Execute "DELETE FROM archive_master WHERE ArchiveDate < DateAdd('yyyy', -4, Date())"
End Sub

Private Sub ShowRecord() 'Check and display records
    If rs Is Nothing Then Exit Sub
    If rs.EOF Or rs.BOF Then Exit Sub

    txtName.Text = rs!Name
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
    
    txtAddress.Text = rs!Address
    txtSymptom.Text = rs!Symptom
    txtContact.Text = rs!Contact
    txtComplain.Text = rs!Complain
    txtTreatment.Text = IIf(IsNull(rs!Treatment), "", rs!Treatment)
    txtDiagnosis.Text = IIf(IsNull(rs!Diagnosis), "", rs!Diagnosis)
End Sub

Private Sub ShowFrame(fra As Frame)

    If fra.Visible Then
        fra.Visible = False
        Exit Sub
    End If
    
    fraAddPatient.Visible = False
    fraRemovePatient.Visible = False
    fraPatientLog.Visible = False
    fraPrescription.Visible = False
    Frame1.Visible = False
    fra.Visible = True

End Sub

Private Sub Clear() 'Clear filled inputs and assign default inputs
    txtName.Text = ""
    txtAge.Text = ""
    dtpDOB.Value = Date
    txtAddress.Text = ""
    cboSex.ListIndex = -1
    txtContact.Text = "09"
    txtComplain.Text = ""
    txtDiagnosis.Text = ""
    txtTreatment.Text = ""
    txtID.Text = ""
    txtRID.Text = ""
    txtPID.Text = ""
    txtGivenMed = ""
    lblSelectedID.Caption = "Selected Patient: None"
    lblSelectedSymptom.Caption = "Patient Symptom: None"
    lblSelectedPatientRID.Caption = "Selected Patient: None"
    SelectID = 0
    cmdPConfirm.Enabled = False
End Sub

Public Sub LoadCombo(cbo As ComboBox) 'Loads combo boxes
    Set rs = New ADODB.Recordset

    If cn Is Nothing Then
        Set cn = New ADODB.Connection
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ClinicRecord.mdb"
    End If

    rs.CursorLocation = adUseClient
    rs.Open "SELECT MedName FROM medicine_master ORDER BY MedName ASC", _
               cn, adOpenStatic, adLockReadOnly

    cbo.Clear
    Do While Not rs.EOF
        cbo.AddItem rs!MedName
        rs.MoveNext
    Loop
    
        cbo.ListIndex = -1

    rs.Close
    Set rs = Nothing
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
    
    If Age < 16 Or Age > 50 Then
        txtAge.Text = ""
        Exit Sub
    End If

    txtAge.Text = Age

End Sub

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


Private Function CapitalizeName(ByVal s As String) As String 'Capitalizes names
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

'Add Patient Codes
Private Sub cmdAddPatient_Click()
    ShowFrame fraAddPatient
    Clear
    Highlight cmdAddPatient, fraAddPatient
End Sub

Private Sub dtpDOB_Change()
    CalculateAge
End Sub

Private Sub Text5_GotFocus()
    Command9.Default = True
    
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

Private Sub cmdAddConfirm_Click()
    If Trim(txtID.Text) = "" Then
        MsgBox "Please enter valid Student ID.", vbExclamation
        Exit Sub
    End If

    If Len(txtID.Text) < 5 Then
        MsgBox "Invalid Student ID.", vbCritical
        txtID.SetFocus
        Exit Sub
    End If
    
    Set rs = New ADODB.Recordset

    rs.Open "SELECT ID FROM patient_master WHERE ID=" & txtID.Text, cn, adOpenStatic, adLockReadOnly

    If Not rs.EOF Then
        MsgBox "The Patient ID already exists! Please enter a unique ID.", vbExclamation
        rs.Close
        Set rs = Nothing
        txtID.SetFocus
     Exit Sub
    End If

    Set rs = New ADODB.Recordset
    rs.Open "patient_master", cn, adOpenDynamic, adLockOptimistic

    'Cleaner
    txtName.Text = Cleaner(txtName.Text)
    txtAddress.Text = Cleaner(txtAddress.Text)
    txtComplain.Text = Cleaner(txtComplain.Text)
    
    If Trim(txtName.Text) = "" Then
        MsgBox "Please enter the Patient Name.", vbExclamation
        Exit Sub
    End If
    
    If Val(txtAge.Text) < 16 Or Val(txtAge.Text) > 50 Then
        MsgBox "Invalid Age.", vbExclamation
        Exit Sub
    End If
    
        If Trim(cboSex.Text) = "" Then
        MsgBox "Please enter the Patient Gender.", vbExclamation
        Exit Sub
    End If
    
        If Trim(txtAddress.Text) = "" Then
        MsgBox "Please enter the Patient Address.", vbExclamation
        Exit Sub
    End If
    

    If MsgBox("Are you sure you want to add this patient record?", vbYesNo + vbQuestion) = vbNo Then
        rs.CancelUpdate
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
    
    rs.AddNew
    rs!ID = txtID.Text
    rs!Name = CapitalizeName(Trim(txtName.Text))
    If IsNumeric(txtAge.Text) And Val(txtAge.Text) >= 0 Then rs!Age = CLng(txtAge.Text)
    rs!DOB = dtpDOB.Value
    rs!Address = Trim(txtAddress.Text)
    rs!Sex = cboSex.Text
    rs!Symptom = Trim(txtSymptom.Text)
    rs!Contact = Trim(txtContact.Text)
    rs!Complain = Trim(txtComplain.Text)

    rs.Update
    MsgBox "Patient record saved successfully!"
    rs.MoveLast
    ShowRecord
    Clear
End Sub

'Prescription Codes
Private Sub cmdPrescription_Click()
    ShowFrame fraPrescription
    Clear
    Highlight cmdPrescription, fraPrescription
End Sub

Private Sub cmdInv_Click()
    frmMedicineInventory.Show vbModal
End Sub

Private Sub cmdMedRefresh_Click()
    LoadCombo cboMedicine
    Clear
End Sub

Private Sub txtName_LostFocus()
    txtName.Text = CapitalizeName(txtName.Text)
End Sub

Private Sub txtPID_Change()
    cmdPConfirm.Enabled = False

    If Trim(txtPID.Text) = "" Then
        lblSelectedID.Caption = "Selected Patient: None"
        lblSelectedSymptom.Caption = "Patient Symptom: None"
        Exit Sub
    End If

    If Not IsNumeric(txtPID.Text) Then
        lblSelectedID.Caption = ""
        lblSelectedSymptom.Caption = ""
        Exit Sub
    End If

    Dim patientID As Long
    patientID = CLng(txtPID.Text)

    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT Name, Symptom, Medicine, Treatment, GivenMed, Diagnosis FROM patient_master WHERE ID = " & patientID, _
                cn, adOpenForwardOnly, adLockReadOnly

    If rs.EOF Then
        lblSelectedID.Caption = "No patient found"
        lblSelectedSymptom.Caption = "Patient Symptom: None"
        Exit Sub
    End If

    lblSelectedID.Caption = "Selected Patient: " & rs!Name
    lblSelectedSymptom.Caption = "Patient Symptom: " & rs!Symptom
    txtDiagnosis.Text = IIf(IsNull(rs!Diagnosis), "", rs!Diagnosis)
    txtTreatment.Text = IIf(IsNull(rs!Treatment), "", rs!Treatment)
    txtGivenMed.Text = IIf(IsNull(rs!GivenMed), "", rs!GivenMed)
    cboMedicine.Text = IIf(IsNull(rs!Medicine), "", rs!Medicine)
    

    If Trim("" & rs!Medicine) <> "" Or _
        Trim("" & rs!Treatment) <> "" Or _
        Trim("" & rs!Diagnosis) <> "" Then
       
        lblSelectedID.Caption = lblSelectedID.Caption
        lblSelectedSymptom.Caption = lblSelectedSymptom.Caption
        MsgBox "Patient already has medication"
        cmdPConfirm.Enabled = False
        txtDiagnosis.Locked = True
        txtTreatment.Locked = True
        txtGivenMed.Locked = True
        cboMedicine.Locked = True
    Else
        cmdPConfirm.Enabled = True
        txtDiagnosis.Locked = False
        txtTreatment.Locked = False
        txtGivenMed.Locked = False
        cboMedicine.Locked = False
        SelectID = patientID
    End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub cmdPEdit_Click()
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master", cn, adOpenStatic, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "No records found.", vbInformation
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    frmEditData.OpenPrescription = True
    frmEditData.Show vbModal
End Sub

Private Sub cmdPConfirm_Click()
    'Cleaner
    txtTreatment.Text = Cleaner(txtTreatment.Text)
    txtDiagnosis.Text = Cleaner(txtDiagnosis.Text)

    If Trim(txtDiagnosis.Text) = "" Then
        MsgBox "Please enter Patient Diagnosis.", vbExclamation
        Exit Sub
    End If

    If Trim(txtTreatment.Text) = "" Then
        MsgBox "Please enter the Patient treatment plan.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Are you sure you want to save this patient's medication?", _
              vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    If cboMedicine.Text <> "" Then

        If Trim(txtGivenMed.Text) = "" Or Not IsNumeric(txtGivenMed.Text) Then
            MsgBox "Please enter a valid quantity to give.", vbExclamation
            Exit Sub
        End If

        Dim QtyToDeduct As Long
        QtyToDeduct = CLng(txtGivenMed.Text)

        If QtyToDeduct <= 0 Then
            MsgBox "Quantity must be greater than zero.", vbExclamation
            Exit Sub
        End If

        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM medicine_master WHERE MedName='" & _
                Replace(cboMedicine.Text, "'", "''") & "'", _
                cn, adOpenDynamic, adLockOptimistic

        If rs.EOF Then
            MsgBox "Medicine not found.", vbCritical
            rs.Close
            Exit Sub
        End If

        If rs!StockQty < QtyToDeduct Then
            MsgBox "Not enough stock available!" & vbCrLf & _
                   "Available: " & rs!StockQty, vbCritical
            rs.Close
            Exit Sub
        End If

        If IsNumeric(txtGivenMed.Text) And Val(txtGivenMed.Text) >= 0 Then rs!GivenMed = CLng(txtGivenMed.Text)
        rs!StockQty = rs!StockQty - QtyToDeduct
        rs.Update
        rs.Close

    End If


    '========================
    'UPDATE PATIENT RECORD
    '========================
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master WHERE ID = " & _
            SelectID, cn, adOpenDynamic, adLockOptimistic

    If rs.EOF Then
        MsgBox "Patient record no longer exists.", vbCritical
        rs.Close
        Exit Sub
    End If

    rs!Medicine = cboMedicine.Text
    rs!Treatment = txtTreatment.Text
    rs!Diagnosis = txtDiagnosis.Text
    rs.Update

    rs.Close
    Set rs = Nothing


    '========================
    'SUCCESS MESSAGE
    '========================
    MsgBox "Medication saved and stock updated successfully!", vbInformation

    ShowRecord
    Clear

End Sub

Private Sub cboMedicine_Change()
    txtGivenMed.Enabled = (cboMedicine.Text <> "")
End Sub
'Archive Codes
Private Sub cmdRemovePatient_Click()
    ShowFrame fraRemovePatient
    Clear
    Highlight cmdRemovePatient, fraRemovePatient
    
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ClinicRecord.mdb"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM archive_master ORDER BY ID ASC", cn, adOpenStatic, adLockReadOnly

    Set DGArchive.DataSource = rs

    Dim col As Integer
    For col = 0 To 12
        DGArchive.Columns(col).Width = 1500
    Next col
    Cleanup
End Sub

Private Sub txtRID_Change()
    If Trim(txtRID.Text) = "" Then
        lblSelectedPatientRID.Caption = "Selected Patient: None"
        Exit Sub
    End If

    If Not IsNumeric(txtRID.Text) Then
        lblSelectedPatientRID.Caption = ""
        Exit Sub
    End If

    Dim patientID As Long
    
    patientID = CLng(txtRID.Text)
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT Name FROM patient_master WHERE ID = " & patientID, _
                cn, adOpenForwardOnly, adLockReadOnly

    If rs.EOF Then
        lblSelectedPatientRID.Caption = "No patient found"
    Else
        lblSelectedPatientRID.Caption = "Selected Patient: " & rs!Name
    End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub cmdRConfirm_Click()
    Dim patientID As Long
    Dim Treatment As String, Diagnosis As String, Medicine As String
    Dim DOB As String, Age As String
    Dim Archive As String

    If Trim(txtRID.Text) = "" Or Not IsNumeric(txtRID.Text) Then
        MsgBox "Enter a valid Patient ID.", vbExclamation
        Exit Sub
    End If

    patientID = CLng(txtRID.Text)
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master WHERE ID=" & patientID, cn, adOpenDynamic, adLockOptimistic

    If rs.EOF Then
        MsgBox "No record found with that ID.", vbInformation
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion) = vbNo Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If IsNull(rs!Treatment) Then
        Treatment = "NULL"
    Else
        Treatment = "'" & Replace$(rs!Treatment, "'", "''") & "'"
    End If

    If IsNull(rs!Diagnosis) Then
        Diagnosis = "NULL"
    Else
        Diagnosis = "'" & Replace$(rs!Diagnosis, "'", "''") & "'"
    End If

    If IsNull(rs!Medicine) Then
        Medicine = "NULL"
    Else
        Medicine = "'" & Replace$(rs!Medicine, "'", "''") & "'"
    End If

    If IsNull(rs!DOB) Then
        DOB = "NULL"
    Else
        DOB = "#" & Format$(rs!DOB, "mm/dd/yyyy") & "#"
    End If

    If IsNull(rs!Age) Then
        Age = "NULL"
    Else
        Age = rs!Age
    End If

    Archive = "INSERT INTO archive_master (ID, Name, Address, Age, DOB, Sex, Symptom, Contact, Complain, Treatment, Diagnosis, Medicine, ArchiveDate) VALUES (" & _
          rs!ID & ", '" & Replace(rs!Name, "'", "''") & "', '" & Replace(rs!Address, "'", "''") & "', " & _
          Age & ", " & DOB & ", '" & Replace(rs!Sex, "'", "''") & "', '" & Replace(rs!Symptom, "'", "''") & "', '" & _
          Replace(rs!Contact, "'", "''") & "', '" & Replace(rs!Complain, "'", "''") & "', " & _
          Treatment & ", " & Diagnosis & ", " & Medicine & ", #" & Format(Date, "mm/dd/yyyy") & "#)"

    cn.Execute Archive
    rs.Delete
    rs.Close
    Set rs = Nothing

    MsgBox "Record archived and deleted successfully!"
    Clear
End Sub

Private Sub cmdRefreshArchive_Click()
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ClinicRecord.mdb"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM archive_master ORDER BY ID ASC", cn, adOpenStatic, adLockReadOnly

    Set DGArchive.DataSource = rs

    Dim col As Integer
    For col = 0 To 12
        DGArchive.Columns(col).Width = 1500
    Next col
    Cleanup
End Sub

'Log Codes
Private Sub cmdReport_Click()
    ShowFrame Frame1
    Clear
    Highlight cmdReport, Frame1
End Sub

Private Sub cmdPatientLog_Click()
    ShowFrame fraPatientLog
    Clear
    Highlight cmdPatientLog, fraPatientLog
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM patient_master ORDER BY ID ASC", cn, adOpenStatic, adLockReadOnly

    Set DGData.DataSource = rs

    Dim col As Integer
    For col = 0 To 12
        DGData.Columns(col).Width = 1500
    Next col
End Sub

Private Sub cmdEdit_Click()
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master", cn, adOpenStatic, adLockReadOnly

    If rs.EOF Then
        MsgBox "No records found.", vbInformation
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    frmEditData.OpenPrescription = False
    frmEditData.InitializeForm
    frmEditData.Show vbModal

    rs.Close
    Set rs = Nothing

End Sub
Private Sub cmdRefresh_Click()
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM patient_master ORDER BY ID ASC", cn, adOpenStatic, adLockReadOnly

    Set DGData.DataSource = rs

    Dim col As Integer
    For col = 0 To 12
        DGData.Columns(col).Width = 1500
    Next col
    Text5.Text = ""
End Sub

'Logout Codes
Private Sub cmdLogout_Click()
        ResetAll
    If MsgBox("Are you sure you want to log out?", vbYesNo + vbQuestion) = vbYes Then
        Unload Me
        frmLogin.Show
        MsgBox "Successfully Logged Out!"
    End If
End Sub

'Key Ascii Codes
Private Sub NumOnly(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
        If KeyAscii < 48 Or KeyAscii > 57 Then
            KeyAscii = 0
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

Private Sub txtContact_KeyPress(KeyAscii As Integer)
    If txtContact.SelStart < 3 And KeyAscii = vbKeyBack Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtPID_KeyPress(KeyAscii As Integer)
    NumOnly KeyAscii
End Sub

Private Sub txtRID_Keypress(KeyAscii As Integer)
    NumOnly KeyAscii
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, False, " .-'"
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, " ,.-#'"
End Sub

Private Sub txtCondition_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, False, " ,.-()"
End Sub

Private Sub txtID_Keypress(KeyAscii As Integer)
    NumOnly KeyAscii
End Sub

Private Sub txtID_Validate(Cancel As Boolean)
    If txtID.Text <> "" And Not IsNumeric(txtID.Text) Then
        MsgBox "Invalid Student ID", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtSymptom_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, False, ""
    
        If KeyAscii = vbKeyBack Then Exit Sub
        If Len(txtSymptom.Text) >= 15 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtComplain_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, " ,.-#'()"
End Sub

Private Sub txtDiagnosis_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, " ,.-#'():/"
End Sub

Private Sub txtTreatment_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, " ,.-#'()"
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    If txtName.Text = "" Then Exit Sub
    If txtName.Text Like "*[!A-Za-zÒ— .'-]*" Then
        MsgBox "Invalid characters in Name.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, ""
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
    If Text5.Text Like "*[!A-Za-z0-9Ò— .'-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtSymptom_Validate(Cancel As Boolean)
    If txtSymptom.Text = "" Then Exit Sub
    If txtSymptom.Text Like "*[!A-Za-zÒ—]*" Then
        MsgBox "Invalid characters in Symptom.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtAddress_Validate(Cancel As Boolean)
    If txtAddress.Text = "" Then Exit Sub
    If txtAddress.Text Like "*[!A-Za-z0-9 ,.\#'-]*" Then
        MsgBox "Invalid characters.", vbExclamation
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

Private Sub txtDiagnosis_Validate(Cancel As Boolean)
    If txtDiagnosis.Text = "" Then Exit Sub
    If txtDiagnosis.Text Like "*[!A-Za-z0-9 ,.\#'():/-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtTreatment_Validate(Cancel As Boolean)
    If txtTreatment.Text = "" Then Exit Sub
    If txtTreatment.Text Like "*[!A-Za-z0-9 ,.\#'()-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtPID_Validate(Cancel As Boolean)
    If txtPID.Text <> "" And Not IsNumeric(txtPID.Text) Then
        MsgBox "Invalid Student ID.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtRID_Validate(Cancel As Boolean)
    If txtRID.Text <> "" And Not IsNumeric(txtRID.Text) Then
        MsgBox "Invalid Student ID.", vbExclamation
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

Private Sub Highlight(cmd As CommandButton, fra As Frame)
    If trigger Is cmd Then
        cmd.BackColor = vbButtonFace
        Set trigger = Nothing
        Exit Sub
    End If
    
    If Not trigger Is Nothing Then
        trigger.BackColor = vbButtonFace
    End If
    
    cmd.BackColor = fra.BackColor
    Set trigger = cmd

End Sub

Private Sub ResetAll()
    If Not trigger Is Nothing Then
        trigger.BackColor = vbButtonFace
        Set trigger = Nothing
        Exit Sub
    End If
End Sub

'Record Load Codes
Private Sub cmdLog_Click()
    ReportType = "LOG"
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient


    rs.Open "SELECT ID, Name, Complain, Diagnosis, Treatment, Medicine, GivenMed " & _
                  "FROM patient_master ORDER BY ID ASC", _
                  cn, adOpenStatic, adLockReadOnly

    If rs.EOF Then
        MsgBox "No records found.", vbInformation
                Set rs = Nothing
        Exit Sub
    End If

    Set DGReport.DataSource = rs
        Dim col As Integer
    For col = 0 To 6
        DGReport.Columns(col).Width = 2500
    Next col
    
    MsgBox "Daily Consultation Log Loaded.", vbInformation
End Sub

Private Sub cmdInventory_Click()
    ReportType = "INVENTORY"
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
                Set rs = Nothing
        Exit Sub
    End If

    Set DGReport.DataSource = rs
        Dim col As Integer
    For col = 0 To 3
        DGReport.Columns(col).Width = 2500
    Next col
    
    MsgBox "Inventory Report Loaded.", vbInformation

End Sub

Private Sub cmdLowStock_Click()
    ReportType = "LOWSTOCK"
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
        Set rs = Nothing
        Exit Sub
    End If

    Set DGReport.DataSource = rs
        Dim col As Integer
    For col = 0 To 3
        DGReport.Columns(col).Width = 2500
    Next col
    
    MsgBox "Low Stock Report Loaded.", vbInformation

End Sub

'Print Codes
Private Sub cmdMedicalCert_Click()
    On Error GoTo PrintError
    Dim PID As String

    frmInputPID.PIDValue = ""
    frmInputPID.Show vbModal

    PID = frmInputPID.PIDValue
    Unload frmInputPID

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

Private Sub PrintMultiLine(ByVal TextValue As String, ByVal LeftMargin As Integer)

    Dim Words() As String
    Dim LineText As String
    Dim i As Integer
    Dim MaxWidth As Single

    MaxWidth = Printer.ScaleWidth - LeftMargin - 500

    Words = Split(TextValue, " ")

    LineText = ""

    For i = 0 To UBound(Words)

        If Printer.TextWidth(LineText & Words(i) & " ") > MaxWidth Then
            Printer.CurrentX = LeftMargin
            Printer.Print LineText
            LineText = Words(i) & " "
        Else
            LineText = LineText & Words(i) & " "
        End If

    Next i

    If LineText <> "" Then
        Printer.CurrentX = LeftMargin
        Printer.Print LineText
    End If

End Sub

Private Sub PrintSection(ByVal Title As String, ByVal TextValue As String)

    Printer.FontBold = True
    Printer.Print Title
    Printer.FontBold = False

    Call PrintWrappedText(TextValue, 600)

    Printer.Print ""

End Sub
Private Sub PrintWrappedText(ByVal TextValue As String, ByVal LeftMargin As Integer)

    Dim MaxWidth As Single
    Dim CurrentLine As String
    Dim i As Integer
    Dim c As String
    Dim TempLine As String

    MaxWidth = Printer.ScaleWidth - LeftMargin - 400

    CurrentLine = ""

    For i = 1 To Len(TextValue)
        c = Mid(TextValue, i, 1)
        TempLine = CurrentLine & c

        If Printer.TextWidth(TempLine) > MaxWidth Then
            Printer.CurrentX = LeftMargin
            Printer.Print CurrentLine
            CurrentLine = c
        Else
            CurrentLine = TempLine
        End If
    Next i


    If CurrentLine <> "" Then
        Printer.CurrentX = LeftMargin
        Printer.Print CurrentLine
    End If

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

    If Printers.Count = 0 Then
        MsgBox "No printer installed.", vbCritical
        Exit Sub
    End If

    Printer.FontName = "Courier New"
    Printer.FontSize = 10

    If ReportType = "LOG" Then

        frmInputPID.PIDValue = ""
        frmInputPID.Show vbModal

        If frmInputPID.PIDValue = "" Then
            MsgBox "Printing cancelled.", vbInformation
            Exit Sub
        End If

        Dim rsPrint As ADODB.Recordset
        Set rsPrint = New ADODB.Recordset

        rsPrint.Open "SELECT ID, Name, Complain, Diagnosis, Treatment, Medicine " & _
                     "FROM patient_master WHERE ID = " & frmInputPID.PIDValue, _
                     cn, adOpenStatic, adLockReadOnly

        If rsPrint.EOF Then
            MsgBox "No consultation record found.", vbExclamation
            rsPrint.Close
            Exit Sub
        End If

        Printer.FontSize = 14
        Printer.FontBold = True
        Printer.Print "CLINIC CONSULTATION RECORD"
        Printer.FontBold = False
        Printer.Print String(70, "-")

        Printer.FontSize = 10
        Printer.Print "Date Printed : " & Format(Now, "mmmm dd, yyyy hh:mm AM/PM")
        Printer.Print ""

        Printer.Print "Patient ID   : " & rsPrint!ID
        Printer.Print "Patient Name : " & rsPrint!Name

        Printer.Print ""
        Printer.Print String(70, "-")
        Printer.Print ""

        Call PrintSection("Patient Explanation:", rsPrint!Complain)
        Call PrintSection("Diagnosis:", rsPrint!Diagnosis)
        Call PrintSection("Treatment:", rsPrint!Treatment)
        Call PrintSection("Medicine Given:", rsPrint!Medicine)

        Printer.Print ""
        Printer.Print "____________________________"
        Printer.Print "Attending Nurse"

        rsPrint.Close
        Set rsPrint = Nothing

    Else

        Printer.FontSize = 14
        Printer.FontBold = True

        If ReportType = "LOWSTOCK" Then
            Printer.Print "LOW STOCK MEDICINE REPORT"
        Else
            Printer.Print "MEDICINE INVENTORY REPORT"
        End If

        Printer.FontBold = False
        Printer.Print "Generated: " & Format(Now, "mmmm dd, yyyy hh:mm AM/PM")
        Printer.Print String(70, "-")

        Printer.Print ""
        Printer.FontBold = True
        Printer.Print "ID"; Tab(6); "MEDICINE"; Tab(28); "MANUFACTURER"; Tab(55); "QTY"; Tab(63); "STATUS"
        Printer.FontBold = False
        Printer.Print String(70, "-")

        rs.MoveFirst

        Do While Not rs.EOF

            Printer.Print _
                rs!MedID; _
                Tab(6); rs!MedName; _
                Tab(28); rs!Manufacturer; _
                Tab(55); rs!StockQty; _
                Tab(63); rs!AlertStatus

            rs.MoveNext

        Loop

        Printer.Print String(70, "-")
        Printer.Print ""
        Printer.Print "End of Report"

    End If

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

    If Not (rs.EOF And rs.BOF) Then rs.MoveFirst
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
                rowText = rowText & rs.Fields(field).Value
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

Private Sub timDisplay_Timer()
Dim Today As Variant
Today = Now
lblTime.Caption = Format(Today, "h:mm:ss ampm")
End Sub



