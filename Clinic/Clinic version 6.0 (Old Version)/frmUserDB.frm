VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUserDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   15885
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMain 
      Height          =   7335
      Left            =   2520
      TabIndex        =   17
      Top             =   0
      Width           =   13335
      Begin VB.Frame fraPrescription 
         Height          =   7335
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   13335
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
            TabIndex        =   59
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
            Left            =   8040
            TabIndex        =   53
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "Select"
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
            TabIndex        =   52
            Top             =   240
            Width           =   1335
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
            Left            =   5760
            TabIndex        =   50
            Top             =   1200
            Width           =   2175
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
            TabIndex        =   48
            Top             =   2640
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
            TabIndex        =   46
            Top             =   2640
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
            TabIndex        =   44
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
            Left            =   3120
            TabIndex        =   51
            Top             =   1320
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
            TabIndex        =   49
            Top             =   2040
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
            TabIndex        =   47
            Top             =   2040
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
            TabIndex        =   45
            Top             =   360
            Width           =   1215
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
            TabIndex        =   43
            Top             =   360
            Width           =   3015
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
            Height          =   735
            Left            =   360
            TabIndex        =   54
            Top             =   1200
            Width           =   3255
         End
      End
      Begin VB.Frame fraAddPatient 
         Height          =   7335
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   13335
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
            TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   7
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtAllergy 
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
            TabIndex        =   6
            Top             =   6120
            Width           =   2175
         End
         Begin VB.TextBox txtSex 
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
            TabIndex        =   5
            Top             =   5280
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
            TabIndex        =   4
            Top             =   4440
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
            Height          =   735
            Left            =   1920
            TabIndex        =   3
            Top             =   3600
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
            TabIndex        =   2
            Top             =   2760
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
            TabIndex        =   24
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
            Top             =   1200
            Width           =   1575
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
            Left            =   360
            TabIndex        =   30
            Top             =   6240
            Width           =   1215
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
            Left            =   360
            TabIndex        =   29
            Top             =   5400
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
            Left            =   360
            TabIndex        =   28
            Top             =   4560
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
            Left            =   360
            TabIndex        =   27
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
            Left            =   360
            TabIndex        =   26
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
            Left            =   360
            TabIndex        =   25
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
            Left            =   360
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Frame fraPatientLog 
         Height          =   7335
         Left            =   0
         TabIndex        =   21
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
            TabIndex        =   39
            Top             =   360
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
            TabIndex        =   38
            Top             =   360
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
            Left            =   3360
            TabIndex        =   37
            Top             =   360
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid DGData 
            Height          =   2535
            Left            =   120
            TabIndex        =   40
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
            TabIndex        =   41
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame fraRemovePatient 
         Height          =   7335
         Left            =   0
         TabIndex        =   20
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
            TabIndex        =   58
            Top             =   1680
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
            TabIndex        =   57
            Top             =   1680
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
            TabIndex        =   35
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
            Left            =   4320
            TabIndex        =   34
            Top             =   840
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid DGArchive 
            Height          =   4335
            Left            =   120
            TabIndex        =   55
            Top             =   2520
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
            TabIndex        =   56
            Top             =   1800
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
            TabIndex        =   36
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
            TabIndex        =   42
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
         TabIndex        =   15
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton cmdLogout 
         Caption         =   "Logout"
         Height          =   735
         Left            =   480
         TabIndex        =   16
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
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cnArchive As ADODB.Connection
Dim rsArchive As ADODB.Recordset
Dim SelectedPatientID As Long

'Necessary Codes
Private Sub Form_Load()
    fraAddPatient.Visible = False
    fraRemovePatient.Visible = False
    fraPrescription.Visible = False
    fraPatientLog.Visible = False
    cmdPConfirm.Enabled = False

    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ClinicRecord.mdb"


    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master", cn, adOpenDynamic, adLockOptimistic


    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Call ShowRecord
    End If
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
    txtAge.Text = IIf(IsNull(rs!Age), "", rs!Age)
    txtDOB.Text = IIf(IsNull(rs!DOB), "", Format(rs!DOB, "mm/dd/yyyy"))
    txtAddress.Text = rs!Address
    txtSex.Text = rs!Sex
    txtAllergy.Text = rs!Allergy
    txtCondition.Text = rs!Condition
    txtContact.Text = rs!Contact
    txtComplain.Text = rs!Complain
    txtMed.Text = IIf(IsNull(rs!Medicine), "", rs!Medicine)
    txtTreatment.Text = IIf(IsNull(rs!Treatment), "", rs!Treatment)
    txtDiagnosis.Text = IIf(IsNull(rs!Diagnosis), "", rs!Diagnosis)
    
End Sub

Private Sub ClearFields()
    txtName.Text = ""
    txtAge.Text = ""
    txtDOB.Text = ""
    txtAddress.Text = ""
    txtSex.Text = ""
    txtAllergy.Text = ""
    txtCondition.Text = ""
    txtContact.Text = ""
    txtComplain.Text = ""
    txtMed.Text = ""
    txtTreatment.Text = ""
    txtDiagnosis.Text = ""
    lblSelectedID.Caption = "Selected ID: None"
    SelectedPatientID = 0
    cmdPConfirm.Enabled = False

End Sub

Private Function GetNextPatientID() As Long
Dim rsID As ADODB.Recordset
Set rsID = New ADODB.Recordset

rsID.Open "SELECT MAX(ID) AS MaxID FROM patient_master", cn, adOpenForwardOnly, adLockReadOnly
    If rsID.EOF Or IsNull(rsID!MaxID) Then
        GetNextPatientID = 1
    Else
        GetNextPatientID = rsID!MaxID + 1
    End If
    rsID.Close
    Set rsID = Nothing
End Function

'Control Panel Codes
Private Sub cmdLogout_Click()
    If MsgBox("Are you sure you want to log out?", vbYesNo + vbQuestion) = vbYes Then
        Unload Me
        frmLogin.Show
        MsgBox "Successfully Logged Out!"
    End If
End Sub

Private Sub cmdAddPatient_Click()
    ShowFrame fraAddPatient
    ClearFields
    
    txtID.Text = GetNextPatientID
    
End Sub

Private Sub cmdRemovePatient_Click()
    ShowFrame fraRemovePatient
    ClearFields
End Sub

Private Sub cmdPrescription_Click()
    ShowFrame fraPrescription
    ClearFields
End Sub

Private Sub cmdPatientLog_Click()
    ShowFrame fraPatientLog
    ClearFields
End Sub

'Add Patient Codes
Private Sub cmdAddConfirm_Click()
    Set rs = New ADODB.Recordset
    rs.Open "patient_master", cn, adOpenDynamic, adLockOptimistic

    rs.AddNew
    rs!Name = txtName.Text

    If IsNumeric(txtAge.Text) And Val(txtAge.Text) >= 0 Then
        rs!Age = CLng(txtAge.Text)
    Else
        rs!Age = Null
    End If

    If IsDate(txtDOB.Text) Then
        rs!DOB = CDate(txtDOB.Text)
    Else
        rs!DOB = Null
    End If

    rs!Address = txtAddress.Text
    rs!Sex = txtSex.Text
    rs!Allergy = txtAllergy.Text
    rs!Condition = txtCondition.Text
    rs!Contact = txtContact.Text
    rs!Complain = txtComplain.Text
    
    If MsgBox("Are you sure you want to add this patient record?", _
              vbYesNo + vbQuestion, "Confirm Add") = vbNo Then
        rs.CancelUpdate
        Exit Sub
    End If

    rs.Update
    MsgBox "Patient record saved successfully!"

    rs.MoveLast
    ShowRecord
    txtID.Text = GetNextPatientID
    ClearFields
    ShowFrame fraAddPatient

End Sub

Private Sub cmdReturn_Click()
    fraAddPatient.Visible = False
    
End Sub

'View Prescription Code
Private Sub cmdSelect_Click()
    If Trim(txtPID.Text) = "" Then
        MsgBox "Please enter a Patient ID to select.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtPID.Text) Then
        MsgBox "ID must be numeric.", vbExclamation
        Exit Sub
    End If

    SelectedPatientID = CLng(txtPID.Text)

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master WHERE ID = " & SelectedPatientID, _
            cn, adOpenDynamic, adLockOptimistic

    If rs.EOF Then
        MsgBox "No patient found with that ID.", vbInformation
        cmdPConfirm.Enabled = False
        Exit Sub
    End If

    ShowRecord

    lblSelectedID.Caption = "Selected ID: " & SelectedPatientID
    cmdPConfirm.Enabled = True
End Sub

Private Sub cmdPEdit_Click()
    frmEditData.OpenPrescriptionOnLoad = True
    frmEditData.Show
End Sub

Private Sub cmdPConfirm_Click()
    If SelectedPatientID = 0 Then
        MsgBox "Please select a patient first using the Select button.", vbExclamation
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master WHERE ID = " & SelectedPatientID, _
            cn, adOpenDynamic, adLockOptimistic

    If rs.EOF Then
        MsgBox "Patient record no longer exists.", vbCritical
        Exit Sub
    End If

    If Not IsNull(rs!Medicine) Or Not IsNull(rs!Treatment) Or Not IsNull(rs!Diagnosis) Then
        MsgBox "This patient already has a prescription. Editing is not allowed in this section.", vbExclamation
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    rs!Medicine = txtMed.Text
    rs!Treatment = txtTreatment.Text
    rs!Diagnosis = txtDiagnosis.Text

    If MsgBox("Are you sure you want to save this patient's prescription?", _
              vbYesNo + vbQuestion, "Confirm Save") = vbNo Then
        rs.CancelUpdate
        Exit Sub
    End If

    rs.Update
    MsgBox "Prescription saved successfully!"

    ShowRecord
    ClearFields
    ShowFrame fraPrescription
End Sub

'Archive Patient Data Codes
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
    rs.Open "SELECT * FROM patient_master WHERE ID = " & patientID, _
            cn, adOpenDynamic, adLockOptimistic
    
    If rs.EOF Then
        MsgBox "No record found with that ID.", vbInformation
        Exit Sub
    End If
    

    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
        

        Set cnArchive = New ADODB.Connection
        cnArchive.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ClinicRecord.mdb"
        

        Dim treatmentVal As String, diagnosisVal As String, medicineVal As String
        If IsNull(rs!Treatment) Then
            treatmentVal = "NULL"
        Else
            treatmentVal = "'" & Replace(rs!Treatment, "'", "''") & "'"
        End If
        
        If IsNull(rs!Diagnosis) Then
            diagnosisVal = "NULL"
        Else
            diagnosisVal = "'" & Replace(rs!Diagnosis, "'", "''") & "'"
        End If
        
        If IsNull(rs!Medicine) Then
            medicineVal = "NULL"
        Else
            medicineVal = "'" & Replace(rs!Medicine, "'", "''") & "'"
        End If
        

        Dim dobVal As String
        If IsNull(rs!DOB) Then
            dobVal = "NULL"
        Else
            dobVal = "#" & Format(rs!DOB, "mm/dd/yyyy") & "#"
        End If
        

        Dim ageVal As String
        If IsNull(rs!Age) Then
            ageVal = "NULL"
        Else
            ageVal = rs!Age
        End If
        

        Dim sqlArchive As String
        sqlArchive = "INSERT INTO archive_master " & _
                     "(ID, Name, Address, Age, DOB, Sex, Allergy, Condition, Contact, Complain, Treatment, Diagnosis, Medicine) VALUES (" & _
                     rs!ID & ", '" & Replace(rs!Name, "'", "''") & "', '" & _
                     Replace(rs!Address, "'", "''") & "', " & _
                     ageVal & ", " & dobVal & ", '" & _
                     rs!Sex & "', '" & _
                     Replace(rs!Allergy, "'", "''") & "', '" & _
                     Replace(rs!Condition, "'", "''") & "', '" & _
                     Replace(rs!Contact, "'", "''") & "', '" & _
                     Replace(rs!Complain, "'", "''") & "', " & _
                     treatmentVal & ", " & diagnosisVal & ", " & medicineVal & ")"
        

        cnArchive.Execute sqlArchive
        cnArchive.Close
        Set cnArchive = Nothing
        

        rs.Delete
        MsgBox "Record archived and deleted successfully!"
        
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdLoadArchive_Click()
    Set cnArchive = New ADODB.Connection
    cnArchive.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ClinicRecord.mdb"
    
    Set rsArchive = New ADODB.Recordset
    rsArchive.CursorLocation = adUseClient
    rsArchive.Open "SELECT * FROM archive_master ORDER BY ID ASC", cnArchive, adOpenStatic, adLockReadOnly
    
    Set DGArchive.DataSource = rsArchive
    
    DGArchive.Columns(0).Width = 500
    DGArchive.Columns(1).Width = 2000
    DGArchive.Columns(2).Width = 500
    DGArchive.Columns(3).Width = 1000
    DGArchive.Columns(4).Width = 2000
    DGArchive.Columns(5).Width = 500
    DGArchive.Columns(6).Width = 1000
    DGArchive.Columns(7).Width = 1000
    DGArchive.Columns(8).Width = 1500
    DGArchive.Columns(9).Width = 1500
    DGArchive.Columns(10).Width = 1500
    DGArchive.Columns(11).Width = 1500
    DGArchive.Columns(12).Width = 1500
    
End Sub

Private Sub cmdRefreshArchive_Click()
    If Not rsArchive Is Nothing Then
        rsArchive.Close
        Set rsArchive = Nothing
    End If
    cmdLoadArchive_Click
End Sub

'View Patient Log Codes
Private Sub cmdEdit_Click()
    Dim rsCheck As ADODB.Recordset
    Set rsCheck = New ADODB.Recordset

    rsCheck.Open "SELECT * FROM patient_master", cn, adOpenStatic, adLockReadOnly
    
    If rsCheck.EOF Then
        MsgBox "No records found.", vbInformation
        rsCheck.Close
        Set rsCheck = Nothing
        Exit Sub
    End If
    frmEditData.Show
    
    rsCheck.Close
    Set rsCheck = Nothing
End Sub

Private Sub cmdLoad_Click()
    
    Dim rsData As ADODB.Recordset

    Set rsData = New ADODB.Recordset
    rsData.CursorLocation = adUseClient
    rsData.Open "SELECT * FROM patient_master ORDER BY ID ASC", cn, adOpenStatic, adLockReadOnly

    Set DGData.DataSource = rsData

    DGData.Columns(0).Width = 500
    DGData.Columns(1).Width = 2000
    DGData.Columns(2).Width = 500
    DGData.Columns(3).Width = 1000
    DGData.Columns(4).Width = 2000
    DGData.Columns(5).Width = 500
    DGData.Columns(6).Width = 1000
    DGData.Columns(7).Width = 1000
    DGData.Columns(8).Width = 1500
    DGData.Columns(9).Width = 1500
    DGData.Columns(10).Width = 1500
    DGData.Columns(11).Width = 1500
    DGData.Columns(12).Width = 1500
End Sub

Private Sub cmdRefresh_Click()
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    cmdLoad_Click
End Sub

