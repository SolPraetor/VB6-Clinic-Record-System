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
      Begin VB.Frame fraPatientLog 
         Height          =   7335
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   13335
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
            Left            =   3240
            TabIndex        =   58
            Top             =   240
            Width           =   1575
         End
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
            Left            =   8160
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
            Left            =   4920
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
            Left            =   6600
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
            TabIndex        =   63
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
      Begin VB.Frame fraPrescription 
         Height          =   7335
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   13335
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
            Left            =   5160
            TabIndex        =   62
            Top             =   1800
            Width           =   1215
         End
         Begin VB.ComboBox cboMedicine 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   2040
            Width           =   2055
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
            TabIndex        =   42
            Top             =   240
            Width           =   2175
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
            Height          =   495
            Left            =   360
            TabIndex        =   50
            Top             =   1200
            Width           =   8055
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
            TabIndex        =   61
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
            TabIndex        =   60
            Top             =   2760
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker dtpDOB 
            Height          =   375
            Left            =   1920
            TabIndex        =   59
            Top             =   3840
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   """MM/dd/yyyy"""
            Format          =   144113667
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
   End
   Begin VB.Frame fraUserControlPanel 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
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
         Top             =   5280
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
         Top             =   2760
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
         Top             =   3600
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
         Top             =   4440
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
Option Explicit

' General
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim SelectID As Long
Public OpenLog As Boolean

'Necessary Codes
Public Sub LoadCombo(cbo As ComboBox)
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

Private Sub cmdMedRefresh_Click()
    LoadCombo cboMedicine
End Sub

Private Sub cmdReport_Click()
    frmReport.Show
End Sub

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

Private Sub ShowFrame(fra As Frame)
    fraAddPatient.Visible = False
    fraRemovePatient.Visible = False
    fraPatientLog.Visible = False
    fraPrescription.Visible = False
    fra.Visible = True
End Sub

Private Sub ShowRecord()
    If rs Is Nothing Then Exit Sub
    If rs.EOF Or rs.BOF Then Exit Sub

    txtName.Text = rs!Name
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
    
    txtAddress.Text = rs!Address
    txtSymptom.Text = rs!Symptom
    txtCondition.Text = rs!Condition
    txtContact.Text = rs!Contact
    txtComplain.Text = rs!Complain
    txtTreatment.Text = IIf(IsNull(rs!Treatment), "", rs!Treatment)
    txtDiagnosis.Text = IIf(IsNull(rs!Diagnosis), "", rs!Diagnosis)
End Sub

Private Sub Clear()
    txtName.Text = ""
    txtAge.Text = ""
    dtpDOB.Value = Date
    txtAddress.Text = ""
    cboSex.ListIndex = -1
    txtSymptom.Text = ""
    txtCondition.Text = ""
    txtContact.Text = "09"
    txtComplain.Text = ""
    lblSelectedID.Caption = "Selected Patient: None"
    lblSelectedPatientRID.Caption = "Selected Patient: None"
    SelectID = 0
    cmdPConfirm.Enabled = False
End Sub

Private Function GetNextID() As Long
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
'Command Codes
Private Sub cmdLogout_Click()
    If MsgBox("Are you sure you want to log out?", vbYesNo + vbQuestion) = vbYes Then
        Unload Me
        frmLogin.Show
        MsgBox "Successfully Logged Out!"
    End If
End Sub

Private Sub cmdAddPatient_Click()
    ShowFrame fraAddPatient
    Clear
    txtID.Text = GetNextID
End Sub

Private Sub cmdRemovePatient_Click()
    ShowFrame fraRemovePatient
    Clear
End Sub

Private Sub cmdPrescription_Click()
    ShowFrame fraPrescription
    Clear
End Sub

Private Sub cmdPatientLog_Click()
    ShowFrame fraPatientLog
    Clear
End Sub


Private Sub cmdInv_Click()
    frmMedicineInventory.Show
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
    If IsNumeric(txtAge.Text) And Val(txtAge.Text) >= 0 Then rs!age = CLng(txtAge.Text)
    rs!dob = dtpDOB.Value
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

Private Sub txtPID_Change()

    cmdPConfirm.Enabled = False   ' Disable by default

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
    frmUserDB.Hide
End Sub

Private Sub cmdPConfirm_Click()
    If SelectID = 0 Then
        MsgBox "Please select a patient first using the Select button.", vbExclamation
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master WHERE ID = " & SelectID, cn, adOpenDynamic, adLockOptimistic

    If rs.EOF Then
        MsgBox "Patient record no longer exists.", vbCritical
        Exit Sub
    End If

    If cboMedicine.Text = "" Then
        MsgBox "Please select a medicine.", vbExclamation
        Exit Sub
    End If

    rs!Medicine = cboMedicine.Text
    rs!Treatment = txtTreatment.Text
    rs!Diagnosis = txtDiagnosis.Text

    If MsgBox("Are you sure you want to save this patient's prescription?", vbYesNo + vbQuestion) = vbNo Then
        rs.CancelUpdate
        Exit Sub
    End If

    rs.Update


    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM medicine_master WHERE MedName='" & Replace(cboMedicine.Text, "'", "''") & "'", cn, adOpenDynamic, adLockOptimistic

    If Not rs.EOF Then
        If rs!StockQty <= 0 Then
            MsgBox "Medicine is OUT OF STOCK!", vbCritical
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        rs!StockQty = rs!StockQty - 1
        rs.Update
    End If

    rs.Close
    Set rs = Nothing

    ShowRecord
    Clear
    ShowFrame fraPrescription
End Sub

Private Sub cmdRConfirm_Click()
    If Trim(txtRID.Text) = "" Then
        MsgBox "Please enter a Patient ID.", vbExclamation
        Exit Sub
    End If
    If Not IsNumeric(txtRID.Text) Then
        MsgBox "ID must be a number.", vbExclamation
        Exit Sub
    End If

    Dim patientID As Long
    patientID = CLng(txtRID.Text)

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master WHERE ID = " & patientID, cn, adOpenDynamic, adLockOptimistic

    If rs.EOF Then
        MsgBox "No record found with that ID.", vbInformation
        Exit Sub
    End If

    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
        Set cn = New ADODB.Connection
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ClinicRecord.mdb"

        Dim Treatment As String, Diagnosis As String, Medicine As String
        Treatment = IIf(IsNull(rs!Treatment), "NULL", "'" & Replace(rs!Treatment, "'", "''") & "'")
        Diagnosis = IIf(IsNull(rs!Diagnosis), "NULL", "'" & Replace(rs!Diagnosis, "'", "''") & "'")
        Medicine = IIf(IsNull(rs!Medicine), "NULL", "'" & Replace(rs!Medicine, "'", "''") & "'")
        Dim dobVal As String, ageVal As String
        dobVal = IIf(IsNull(rs!dob), "NULL", "#" & Format(rs!dob, "mm/dd/yyyy") & "#")
        ageVal = IIf(IsNull(rs!age), "NULL", rs!age)

        Dim sqlArchive As String
        sqlArchive = "INSERT INTO archive_master (ID, Name, Address, Age, DOB, Sex, Symptom, Condition, ContactNo, Complain, Treatment, Diagnosis, Medicine) VALUES (" & _
                     rs!ID & ", '" & Replace(rs!Name, "'", "''") & "', '" & Replace(rs!Address, "'", "''") & "', " & _
                     ageVal & ", " & dobVal & ", '" & Replace(rs!Sex, "'", "''") & "', '" & Replace(rs!Symptom, "'", "''") & "', '" & _
                     Replace(rs!Condition, "'", "''") & "', '" & Replace(rs!Contact, "'", "''") & "', '" & _
                     Replace(rs!Complain, "'", "''") & "', " & Treatment & ", " & Diagnosis & ", " & Medicine & ")"

        cn.Execute sqlArchive
        cn.Close
        Set cn = Nothing

        rs.Delete
        MsgBox "Record archived and deleted successfully!"
    End If

    rs.Close
    Set rs = Nothing
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

    Unload Me

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

Private Sub txtContact_KeyPress(KeyAscii As Integer)
    If txtContact.SelStart < 3 And KeyAscii = vbKeyBack Then
        KeyAscii = 0
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
