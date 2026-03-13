VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21480
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   21480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPatientLog 
      BackColor       =   &H00C0C0C0&
      Height          =   7335
      Left            =   -6600
      TabIndex        =   58
      Top             =   1200
      Width           =   13335
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
         Height          =   615
         Left            =   1800
         TabIndex        =   63
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Data"
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
         Left            =   240
         TabIndex        =   62
         Top             =   960
         Width           =   1455
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
         Height          =   615
         Left            =   3360
         TabIndex        =   61
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   240
         MaxLength       =   4
         TabIndex        =   60
         Top             =   1680
         Width           =   3375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Search"
         Height          =   495
         Left            =   3840
         TabIndex        =   59
         Top             =   1680
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DGData 
         Height          =   4695
         Left            =   240
         TabIndex        =   64
         Top             =   2280
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   8281
         _Version        =   393216
         BackColor       =   14737632
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
            Name            =   "Tahoma"
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
      Begin VB.Shape Shape3 
         BackColor       =   &H00800000&
         BorderColor     =   &H00800000&
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   0
         Top             =   120
         Width           =   13335
      End
      Begin VB.Label lblLogs 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
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
         Left            =   0
         TabIndex        =   65
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   7215
      Left            =   3240
      TabIndex        =   51
      Top             =   3240
      Visible         =   0   'False
      Width           =   13335
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print Log"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6360
         TabIndex        =   56
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdMedicalCert 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4800
         TabIndex        =   55
         ToolTipText     =   "Generate Medical Certificate"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdLowStock 
         Caption         =   "Low Stocks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         TabIndex        =   54
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdInventory 
         Caption         =   "Full Inventory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         TabIndex        =   53
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdLog 
         Caption         =   "Consultation Log"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   52
         Top             =   1080
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DGReport 
         Height          =   5055
         Left            =   120
         TabIndex        =   57
         Top             =   1920
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   8916
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
      Begin VB.Shape Shape14 
         BackColor       =   &H00800000&
         BorderColor     =   &H00800000&
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   0
         Top             =   120
         Width           =   13335
      End
   End
   Begin VB.Frame fraAddPatient 
      BackColor       =   &H00C0C0C0&
      Height          =   7335
      Left            =   -7320
      TabIndex        =   28
      Top             =   2880
      Width           =   13335
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
         TabIndex        =   40
         Text            =   "Don't Remove, Put under Date Dropdown"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddSymptom 
         Caption         =   "Add New Symptom"
         Height          =   255
         Left            =   11520
         TabIndex        =   38
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
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
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
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
         TabIndex        =   36
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
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
         TabIndex        =   35
         Top             =   2520
         Width           =   7095
      End
      Begin VB.TextBox txtComplain 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   4080
         Width           =   12855
      End
      Begin VB.TextBox txtContact 
         Appearance      =   0  'Flat
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
         Left            =   10560
         MaxLength       =   11
         TabIndex        =   33
         Top             =   2520
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
         Left            =   11760
         TabIndex        =   32
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
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
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox cboSex 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   10560
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox txtSymptom 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   10560
         List            =   "Form1.frx":0002
         TabIndex        =   29
         Text            =   "txtSymptom"
         Top             =   1800
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   495
         Left            =   6720
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   """MM/dd/yyyy"""
         Format          =   9175043
         CurrentDate     =   46073
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H0000C000&
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   240
         Top             =   3360
         Width           =   9855
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00800000&
         BorderColor     =   &H00800000&
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   0
         Top             =   120
         Width           =   13335
      End
      Begin VB.Label lblAddPatient 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
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
         Left            =   0
         TabIndex        =   50
         Top             =   240
         Width           =   4215
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
         TabIndex        =   49
         Top             =   1200
         Width           =   1215
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
         TabIndex        =   48
         Top             =   1920
         Width           =   1215
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
         Height          =   615
         Left            =   5160
         TabIndex        =   47
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblDOB 
         BackStyle       =   0  'Transparent
         Caption         =   "BirthDate:"
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
         Left            =   5160
         TabIndex        =   46
         Top             =   1200
         Width           =   1575
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
         TabIndex        =   45
         Top             =   2640
         Width           =   1455
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
         TabIndex        =   44
         Top             =   1200
         Width           =   1215
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
         TabIndex        =   43
         Top             =   1800
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
         TabIndex        =   42
         Top             =   3480
         Width           =   3855
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
         TabIndex        =   41
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H0000C000&
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   10320
         Top             =   3360
         Width           =   375
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H0000C000&
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   10920
         Top             =   3360
         Width           =   2055
      End
   End
   Begin VB.Frame fraRemovePatient 
      BackColor       =   &H00C0C0C0&
      Height          =   7335
      Left            =   5760
      TabIndex        =   18
      Top             =   -120
      Width           =   13335
      Begin VB.CommandButton cmdLoadArchive 
         Caption         =   "Load Archive"
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
         Left            =   240
         TabIndex        =   22
         Top             =   1560
         Width           =   1575
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
         Height          =   615
         Left            =   1920
         TabIndex        =   21
         Top             =   1560
         Width           =   1455
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
         MaxLength       =   4
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
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
         Left            =   3480
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DGArchive 
         Height          =   4215
         Left            =   240
         TabIndex        =   23
         Top             =   2880
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   7435
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
            Name            =   "Tahoma"
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
      Begin VB.Shape Shape7 
         BackColor       =   &H00800000&
         BorderColor     =   &H00800000&
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   0
         Top             =   120
         Width           =   13335
      End
      Begin VB.Label lblSelectedPatientRID 
         BackColor       =   &H00C0C0C0&
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
         Left            =   3600
         TabIndex        =   27
         Top             =   1680
         Width           =   7095
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
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   2280
         Width           =   4335
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
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
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
         Left            =   0
         TabIndex        =   24
         Top             =   240
         Width           =   4455
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H0000C000&
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   4800
         Top             =   2280
         Width           =   375
      End
      Begin VB.Shape Shape16 
         BackColor       =   &H0000C000&
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   5400
         Top             =   2280
         Width           =   7695
      End
   End
   Begin VB.Frame fraPrescription 
      BackColor       =   &H00C0C0C0&
      Height          =   7335
      Left            =   -4200
      TabIndex        =   0
      Top             =   -840
      Width           =   13335
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
         Left            =   4680
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
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
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
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
         Height          =   735
         Left            =   11400
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
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
         Height          =   735
         Left            =   11400
         TabIndex        =   6
         Top             =   960
         Width           =   1695
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
         Left            =   4680
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtTreatment 
         Appearance      =   0  'Flat
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
         Left            =   6840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3840
         Width           =   6255
      End
      Begin VB.TextBox txtDiagnosis 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   3840
         Width           =   6255
      End
      Begin VB.TextBox txtPID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtGivenMed 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   1
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00800000&
         BorderColor     =   &H00800000&
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   0
         Top             =   120
         Width           =   13335
      End
      Begin VB.Label lblSelectedID 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Selected ID: None"
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
         Height          =   1095
         Left            =   6120
         TabIndex        =   17
         Top             =   1080
         Width           =   5055
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
         TabIndex        =   16
         Top             =   1680
         Width           =   2655
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
         TabIndex        =   15
         Top             =   2280
         Width           =   2535
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
         Left            =   6840
         TabIndex        =   14
         Top             =   3240
         Width           =   6255
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
         Left            =   240
         TabIndex        =   13
         Top             =   3240
         Width           =   6255
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
         TabIndex        =   12
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
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
         TabIndex        =   11
         Top             =   240
         Width           =   3735
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0000C000&
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   6000
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblSelectedSymptom 
         BackColor       =   &H00C0C0C0&
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
         Left            =   6120
         TabIndex        =   10
         Top             =   2280
         Width           =   5055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
