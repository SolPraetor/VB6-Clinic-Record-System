VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUserDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Height          =   7335
      Left            =   2520
      TabIndex        =   15
      Top             =   0
      Width           =   13335
      Begin VB.Frame fraPrescription 
         Height          =   7335
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   13335
         Begin VB.TextBox txtGivenMed 
            Height          =   615
            Left            =   3120
            MaxLength       =   1
            TabIndex        =   64
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CommandButton cmdMedRefresh 
            Caption         =   "Refresh"
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
            Left            =   4920
            TabIndex        =   61
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cboMedicine 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1320
            Width           =   1695
         End
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
            Left            =   9600
            TabIndex        =   56
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdPEdit 
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
            Left            =   7800
            TabIndex        =   55
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdPConfirm 
            Caption         =   "Confirm"
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
            Left            =   6360
            TabIndex        =   49
            Top             =   240
            Width           =   1335
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
            Height          =   4215
            Left            =   6720
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   3000
            Width           =   6255
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
            Height          =   4215
            Left            =   360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   3000
            Width           =   6255
         End
         Begin VB.TextBox txtPID 
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
            MaxLength       =   4
            TabIndex        =   42
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblMedicine 
            Caption         =   "Select Medicine"
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
            Left            =   360
            TabIndex        =   65
            Top             =   1200
            Width           =   2895
         End
         Begin VB.Label lblMedicineG 
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
            Left            =   360
            TabIndex        =   48
            Top             =   1920
            Width           =   2535
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
            Left            =   6720
            TabIndex        =   47
            Top             =   2520
            Width           =   1815
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
            Left            =   360
            TabIndex        =   45
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label lblPID 
            Caption         =   "ID"
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
            Left            =   3240
            TabIndex        =   43
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblSelectedID 
            Caption         =   "Selected ID: None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   6360
            TabIndex        =   50
            Top             =   1200
            Width           =   6615
         End
         Begin VB.Label Label3 
            Caption         =   "Medicine Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame fraAddPatient 
         Height          =   7335
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   13335
         Begin VB.ComboBox cboSex 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   4680
            Width           =   2175
         End
         Begin VB.TextBox txtAge 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   2760
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker dtpDOB 
            Height          =   375
            Left            =   1920
            TabIndex        =   58
            Top             =   3840
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   """MM/dd/yyyy"""
            Format          =   142606339
            CurrentDate     =   46073
         End
         Begin VB.CommandButton cmdAddConfirm 
            Caption         =   "Confirm"
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
            Left            =   3720
            TabIndex        =   8
            Top             =   240
            Width           =   1335
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
            Left            =   5160
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtContact 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            Top             =   1920
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
            Height          =   3255
            Left            =   4200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   3600
            Width           =   8895
         End
         Begin VB.TextBox txtCondition 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   6000
            TabIndex        =   5
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtSymptom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1920
            TabIndex        =   4
            Top             =   6120
            Width           =   2175
         End
         Begin VB.TextBox txtAddress 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1920
            TabIndex        =   3
            Top             =   5280
            Width           =   2175
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
            Left            =   1920
            TabIndex        =   2
            Top             =   3840
            Width           =   2175
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1920
            TabIndex        =   1
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox txtID 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1080
            Width           =   2175
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
            TabIndex        =   31
            Top             =   2040
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
            TabIndex        =   30
            Top             =   2880
            Width           =   3375
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
            Left            =   4200
            TabIndex        =   29
            Top             =   1200
            Width           =   1575
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
            Left            =   240
            TabIndex        =   28
            Top             =   6240
            Width           =   1575
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
            TabIndex        =   27
            Top             =   4560
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
            TabIndex        =   26
            Top             =   5400
            Width           =   1455
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
            TabIndex        =   25
            Top             =   3720
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
            TabIndex        =   24
            Top             =   2880
            Width           =   1215
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
            TabIndex        =   23
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblID 
            Caption         =   "ID"
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
            TabIndex        =   21
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblAddPatient 
            Caption         =   "Add Patient Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Frame fraPatientLog 
         Height          =   7335
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   13335
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
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
            TabIndex        =   37
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdLoad 
            Caption         =   "Load Data"
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
            Left            =   3360
            TabIndex        =   36
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdEdit 
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
            Left            =   5040
            TabIndex        =   35
            Top             =   240
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid DGData 
            Height          =   2535
            Left            =   120
            TabIndex        =   38
            Top             =   1200
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   4471
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
         Begin VB.Label lblLogs 
            Caption         =   "Patient Data Logs"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame fraRemovePatient 
         Height          =   7335
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   13335
         Begin VB.CommandButton cmdLoadArchive 
            Caption         =   "Load Archive"
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
            Left            =   3840
            TabIndex        =   54
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton cmdRefreshArchive 
            Caption         =   "Refresh"
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
            Left            =   5520
            TabIndex        =   53
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox txtRID 
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
            MaxLength       =   4
            TabIndex        =   33
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton cmdRConfirm 
            Caption         =   "Confirm"
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
            Left            =   4200
            TabIndex        =   32
            Top             =   840
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid DGArchive 
            Height          =   4335
            Left            =   120
            TabIndex        =   51
            Top             =   2760
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   7646
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
         Begin VB.Label lblSelectedPatientRID 
            Caption         =   "Selected Patient: None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   5760
            TabIndex        =   62
            Top             =   960
            Width           =   7095
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Archive Logs"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   52
            Top             =   1920
            Width           =   3975
         End
         Begin VB.Label lblRID 
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
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Archive Patient Data"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   3975
         End
      End
   End
   Begin VB.Frame fraUserControlPanel 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton cmdReport 
         Caption         =   "Create Patient Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   63
         Top             =   5760
         Width           =   1575
      End
      Begin VB.CommandButton cmdPatientLog 
         Caption         =   "View Patient Log"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   13
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddPatient 
         Caption         =   "Add Patient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   10
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrescription 
         Caption         =   "Add Patient Prescription"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   11
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton cmdRemovePatient 
         Caption         =   "Archive Patient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   12
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton cmdLogout 
         Caption         =   "Logout"
         Height          =   735
         Left            =   480
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         Height          =   1095
         Left            =   600
         Shape           =   2  'Oval
         Top             =   960
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   960
         Shape           =   2  'Oval
         Top             =   360
         Width           =   615
      End
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

'Main Logic
Private Sub Form_Load()
    fraAddPatient.Visible = False
    fraRemovePatient.Visible = False
    fraPrescription.Visible = False
    fraPatientLog.Visible = False
    cmdPConfirm.Enabled = False

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
    
    cboSex.Clear
    cboSex.AddItem "Male"
    cboSex.AddItem "Female"
    cboSex.ListIndex = 0

    dtpDOB.MaxDate = Date
    dtpDOB.MinDate = DateAdd("yyyy", -120, Date)
    
End Sub

'Helper Codes
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
    txtCondition.Text = rs!Condition
    txtContact.Text = rs!Contact
    txtComplain.Text = rs!Complain
    txtTreatment.Text = IIf(IsNull(rs!Treatment), "", rs!Treatment)
    txtDiagnosis.Text = IIf(IsNull(rs!Diagnosis), "", rs!Diagnosis)
End Sub

Private Sub ShowFrame(fra As Frame) 'Hide Frames
    fraAddPatient.Visible = False
    fraRemovePatient.Visible = False
    fraPatientLog.Visible = False
    fraPrescription.Visible = False
    fra.Visible = True
End Sub

Private Sub Clear() 'Clear filled inputs and assign default inputs
    txtName.Text = ""
    txtAge.Text = ""
    dtpDOB.Value = Date
    txtAddress.Text = ""
    cboSex.ListIndex = -1
    txtSymptom.Text = ""
    txtCondition.Text = ""
    txtContact.Text = "09"
    txtComplain.Text = ""
    txtDiagnosis.Text = ""
    txtTreatment.Text = ""
    txtRID.Text = ""
    txtPID.Text = ""
    txtGivenMed = ""
    lblSelectedID.Caption = "Selected Patient: None"
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

    If Age < 0 Or Age > 121 Then
        Exit Sub
    End If

    txtAge.Text = Age

End Sub

'Function Codes
Private Function GetNextID() As Long 'Display next unused IDs
    Set rs = New ADODB.Recordset

    rs.Open "SELECT MAX(ID) AS MaxID FROM patient_master", cn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Or IsNull(rs!MaxID) Then
        GetNextID = 1
    Else
        GetNextID = rs!MaxID + 1
    End If
    rs.Close
    Set rs = Nothing
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
    txtID.Text = GetNextID
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

Private Sub cmdAddConfirm_Click()
    Set rs = New ADODB.Recordset
    rs.Open "patient_master", cn, adOpenDynamic, adLockOptimistic

    If Trim(txtName.Text) = "" Then
        MsgBox "Please enter the Patient's Name.", vbExclamation
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
    rs!Name = txtName.Text
    If IsNumeric(txtAge.Text) And Val(txtAge.Text) >= 0 Then rs!Age = CLng(txtAge.Text)
    rs!DOB = dtpDOB.Value
    rs!Address = txtAddress.Text
    rs!Sex = cboSex.Text
    rs!Symptom = txtSymptom.Text
    rs!Condition = txtCondition.Text
    rs!Contact = txtContact.Text
    rs!Complain = txtComplain.Text

    rs.Update
    MsgBox "Patient record saved successfully!"
    rs.MoveLast
    ShowRecord
    txtID.Text = GetNextID
    Clear
    ShowFrame fraAddPatient
End Sub

Private Sub cmdReturn_Click()
    fraAddPatient.Visible = False
End Sub

'Prescription Codes
Private Sub cmdPrescription_Click()
    ShowFrame fraPrescription
    Clear
End Sub

Private Sub cmdInv_Click()
    frmMedicineInventory.Show
End Sub

Private Sub cmdMedRefresh_Click()
    LoadCombo cboMedicine
End Sub

Private Sub txtPID_Change()
    cmdPConfirm.Enabled = False

    If Trim(txtPID.Text) = "" Then
        lblSelectedID.Caption = "Selected Patient: None"
        Exit Sub
    End If

    If Not IsNumeric(txtPID.Text) Then
        lblSelectedID.Caption = ""
        Exit Sub
    End If

    Dim patientID As Long
    patientID = CLng(txtPID.Text)

    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT Name, Medicine, Treatment, Diagnosis FROM patient_master WHERE ID = " & patientID, _
                cn, adOpenForwardOnly, adLockReadOnly

    If rs.EOF Then
        lblSelectedID.Caption = "No patient found"
        Exit Sub
    End If


    lblSelectedID.Caption = "Selected Patient: " & rs!Name
    

    If Not IsNull(rs!Medicine) Or _
       Not IsNull(rs!Treatment) Or _
       Not IsNull(rs!Diagnosis) Then
       
        lblSelectedID.Caption = lblSelectedID.Caption
        MsgBox "Patient Already Has Prescription"
        cmdPConfirm.Enabled = False
    Else
        cmdPConfirm.Enabled = True
        SelectID = patientID
    End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub cmdPEdit_Click()
    frmEditData.OpenPrescription = True
    frmEditData.Show
End Sub

Private Sub cmdPConfirm_Click()
    If cboMedicine.Text = "" Then
        MsgBox "Please select a medicine.", vbExclamation
        Exit Sub
    End If

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

    If MsgBox("Are you sure you want to save this patient's prescription?", _
              vbYesNo + vbQuestion) = vbNo Then
        rs.Close
        Exit Sub
    End If

    rs!StockQty = rs!StockQty - QtyToDeduct
    rs.Update
    rs.Close

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

    MsgBox "Prescription saved and stock updated successfully!", vbInformation

    ShowRecord
    Clear
    ShowFrame fraPrescription

End Sub

'Archive Codes
Private Sub cmdRemovePatient_Click()
    ShowFrame fraRemovePatient
    Clear
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
        Age = CStr(rs!Age)
    End If

    Archive = "INSERT INTO archive_master (ID, Name, Address, Age, DOB, Sex, Symptom, Condition, Contact, Complain, Treatment, Diagnosis, Medicine) VALUES (" & _
                 rs!ID & ", '" & Replace(rs!Name, "'", "''") & "', '" & Replace(rs!Address, "'", "''") & "', " & _
                 Age & ", " & DOB & ", '" & Replace(rs!Sex, "'", "''") & "', '" & Replace(rs!Symptom, "'", "''") & "', '" & _
                 Replace(rs!Condition, "'", "''") & "', '" & Replace(rs!Contact, "'", "''") & "', '" & _
                 Replace(rs!Complain, "'", "''") & "', " & Treatment & ", " & Diagnosis & ", " & Medicine & ")"

    cn.Execute Archive
    rs.Delete
    rs.Close
    Set rs = Nothing

    MsgBox "Record archived and deleted successfully!"
    Clear
End Sub

Private Sub cmdLoadArchive_Click()
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
End Sub

Private Sub cmdRefreshArchive_Click()
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    cmdLoadArchive_Click
End Sub

'Log Codes
Private Sub cmdReport_Click()
    frmReport.Show
End Sub

Private Sub cmdPatientLog_Click()
    ShowFrame fraPatientLog
    Clear
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
    frmEditData.Show

    rs.Close
    Set rs = Nothing

End Sub

Private Sub cmdLoad_Click()
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM patient_master ORDER BY ID ASC", cn, adOpenStatic, adLockReadOnly

    Set DGData.DataSource = rs

    Dim col As Integer
    For col = 0 To 12
        DGData.Columns(col).Width = 1500
    Next col
End Sub

Private Sub cmdRefresh_Click()
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    cmdLoad_Click
End Sub

'Logout Codes
Private Sub cmdLogout_Click()
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

Private Sub txtSymptom_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, False, " ,.-()"
End Sub

Private Sub txtComplain_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, " ,.-#()"
End Sub

Private Sub txtDiagnosis_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, " ,.-#()"
End Sub

Private Sub txtTreatment_KeyPress(KeyAscii As Integer)
    FilterInput KeyAscii, True, " ,.-#()"
End Sub

Private Sub txtName_LostFocus()
    txtName.Text = CapitalizeName(txtName.Text)
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
    If txtAddress.Text Like "*[!A-Za-z0-9 ,.\#'-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtCondition_Validate(Cancel As Boolean)
    If txtCondition.Text = "" Then Exit Sub
        If txtCondition.Text Like "*[!A-Za-z ,.#()-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtSymptom_Validate(Cancel As Boolean)
    If txtSymptom.Text = "" Then Exit Sub
        If txtCondition.Text Like "*[!A-Za-z ,.#()-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtComplain_Validate(Cancel As Boolean)
    If txtComplain.Text = "" Then Exit Sub
    If txtComplain.Text Like "*[!A-Za-z0-9 ,.\#()-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtDiagnosis_Validate(Cancel As Boolean)
    If txtDiagnosis.Text = "" Then Exit Sub
    If txtDiagnosis.Text Like "*[!A-Za-z0-9 ,.\#()-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtTreatment_Validate(Cancel As Boolean)
    If txtTreatment.Text = "" Then Exit Sub
    If txtTreatment.Text Like "*[!A-Za-z0-9 ,.\#()-]*" Then
        MsgBox "Invalid characters.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtPID_Validate(Cancel As Boolean)
    If txtID.Text <> "" And Not IsNumeric(txtID.Text) Then
        MsgBox "Numbers only.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtRID_Validate(Cancel As Boolean)
    If txtID.Text <> "" And Not IsNumeric(txtID.Text) Then
        MsgBox "Numbers only.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtGivenMed_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub
