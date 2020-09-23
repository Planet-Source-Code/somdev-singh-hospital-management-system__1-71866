VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form TestTransaction 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Transaction"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9315
   Begin VB.TextBox txtReceive 
      Height          =   285
      Left            =   7800
      TabIndex        =   42
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      Picture         =   "TestTransaction.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      Picture         =   "TestTransaction.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3480
      Top             =   7320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from AInt"
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3480
      Top             =   6960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from AInt"
      Caption         =   "c3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   360
      Top             =   7320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from TestTransaction"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   6960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TestTransaction"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00FFC0FF&
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      Picture         =   "TestTransaction.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   3375
      Left            =   5280
      TabIndex        =   29
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1440
         Picture         =   "TestTransaction.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtTNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   31
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   30
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblTNo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TreatNo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblId 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PatientID "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Which record you want to update?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   3375
      Left            =   5280
      TabIndex        =   24
      Top             =   3480
      Width           =   3855
      Begin VB.TextBox txtBal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         TabIndex        =   25
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtPay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Received"
         Height          =   285
         Left            =   2520
         TabIndex        =   33
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdNSave 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   720
         Picture         =   "TestTransaction.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdNCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         Picture         =   "TestTransaction.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblReceive 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Your received amount is"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblBal 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Your balance is"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblPay 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Now you are paying"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      Picture         =   "TestTransaction.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      Picture         =   "TestTransaction.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C0C0C0&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "TestTransaction.frx":2210
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   5895
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtTest 
         Appearance      =   0  'Flat
         DataField       =   "Treatment"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1560
         TabIndex        =   41
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         DataField       =   "Date"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1560
         TabIndex        =   39
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         DataField       =   "Received"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtTreatNo 
         Appearance      =   0  'Flat
         DataField       =   "TreatNo"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtRefDoctor 
         Appearance      =   0  'Flat
         DataField       =   "RefDoc"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtBalance 
         Appearance      =   0  'Flat
         DataField       =   "Balance"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox txtCharge 
         Appearance      =   0  'Flat
         DataField       =   "Charge"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtSpecification 
         Appearance      =   0  'Flat
         DataField       =   "Cause"
         DataSource      =   "Adodc1"
         Height          =   615
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         DataField       =   "PatientID"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblTreatNo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TreatNo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblRefDoctor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ref Doctor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblBalance 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label lblReceived 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Received"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label lblTest 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Treatment/Test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblCharge 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Charge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label lblSpecification 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Specification"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblPId 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Patient ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Test Transaction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "******************************************************************************************************************"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.Line Line1 
      X1              =   4920
      X2              =   4920
      Y1              =   240
      Y2              =   8760
   End
End
Attribute VB_Name = "TestTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim doc
Dim flag As Boolean
Dim query As String
Private Sub cboTest_Change()

Dim i As Integer
Dim test As String

' Adodc2.Recordset.MoveFirst
 'test = cboTest.Text
 'test = "Treatment = '" & test & "'"
 'With Adodc2.Recordset
     '.FindFirst test

'End With

'If Adodc2.Recordset.NoMatch = False Then
' txtCharge.Text = (Adodc2.Recordset.Fields(1).Value)
 
'For i = 0 To Adodc2.Recordset.RecordCount + 1
' test = cboTest.Text
' test = "Treatment= '" & test & "'"
'
'With Adodc2.Recordset
 '    .FindNext test
'End With

'If Adodc2.Recordset.NoMatch = False Then
' txtCharge.Text = (Adodc2.Recordset.Fields(1).Value)
'End If

      
'Next
'Exit Sub
'End If

End Sub

Private Sub cboTest_Click()

Dim i
Dim test As String
 
 Adodc2.Recordset.MoveFirst
 test = cboTest.Text
 test = "Treatment = '" & test & "'"
 
With Adodc2.Recordset
     '.FindFirst test

End With

If Adodc2.Recordset.NoMatch = False Then
   txtCharge.Text = (Adodc2.Recordset.Fields(1).Value)

For i = 0 To Adodc2.Recordset.RecordCount + 1
   test = cboTest.Text
   test = "Treatment = '" & test & "'"

With Adodc2.Recordset
     .FindNext test
End With

If Adodc2.Recordset.NoMatch = False Then
   txtCharge.Text = (Adodc2.Recordset.Fields(1).Value)
End If
      
Next

Exit Sub

End If
End Sub

Private Sub cmdNew_Click()

' cboTest.Locked = False
 'txtDate.Locked = False
 txtPId.Locked = False
 txtSpecification.Locked = False
 
 txtRefDoctor.Locked = False
 'txtTreatNo.Locked = False
 txtReceived.Locked = False
 txtBalance.Locked = False

  
 txtDate.Text = ""
 txtTreatNo.Text = ""
 txtPId.Text = ""
 txtSpecification.Text = ""
 txtCharge.Text = ""
 txtReceived.Locked = False
 txtReceived.Text = ""
 txtBalance.Text = ""
 txtRefDoctor.Text = ""
 'cboTest.Text = ""
 
 txtDate.Text = Date
 txtPId.Enabled = True
 txtPId.SetFocus
 
 cmdCancel.Enabled = True
 cmdNew.Enabled = False
 cmdSave.Enabled = True
 CmdExit.Enabled = False
 cmdUpdate.Enabled = False
 
End Sub

Private Sub cmdSave_Click()

'DischargeInt.Hide
'Form2.Hide


If txtPId.Text = "" Or txtSpecification.Text = "" Or txtRefDoctor.Text = "" Or txtTest.Text = "" Then
 MsgBox "Please fill all required information"
ElseIf txtReceived.Text = "" Then txtReceived.Text = "0"

Else
'If txtReceived.Text <> "0" Then
'Form3oo.txtDate.Text = txtDate.Text
'Form3oo.txtPId.Text = txtPId.Text
'Form3oo.txtSpecification.Text = txtSpecification.Text
'Form3oo.txtCharge.Text = txtCharge.Text
'Form3oo.txtReceived.Text = txtReceived.Text
'Form3oo.txtBalance.Text = txtBalance.Text
'Form3oo.txtRefDoctor.Text = txtRefDoctor.Text
'Form3oo.txtTreatNo.Text = txtTreatNo.Text
'Form3oo.txtReceive.Text = cboTest.Text

'End If
cmdSave.Enabled = False
 cmdCancel.Enabled = False
 cmdNew.Enabled = True
 CmdExit.Enabled = True

 
 Adodc1.Recordset.Update
 
 cmdUpdate.Enabled = True
 
 txtReceived.Locked = True
 txtTest.Locked = True
 txtDate.Locked = True
 txtPId.Locked = True
 txtPId.Enabled = False
 txtSpecification.Locked = True
 txtCharge.Locked = True
 'txtCharge.Enabled = False
 txtReceived.Locked = True
 txtRefDoctor.Locked = True
 txtTreatNo.Locked = True
 txtBalance.Locked = True

 
' cmdNew.SetFocus
' Me.Hide
 'Form3oo.Show
End If

End Sub

Private Sub cmdexit_Click()
 TestTransaction.Hide
End Sub

Private Sub cmdcancel_Click()
 txtDate.Enabled = True
 txtPId.Enabled = True

If txtPId.Text = "" Then
 txtSpecification.Text = ""
 txtCharge.Text = ""
 txtReceived.Text = ""
 txtBalance.Text = ""
 txtRefDoctor.Text = ""
 'cboTest.Text = ""
 txtTreatNo.Text = ""

 cmdSave.Enabled = False
 cmdCancel.Enabled = False
 cmdNew.Enabled = True
 CmdExit.Enabled = True
 cmdUpdate.Enabled = True
Else
 Adodc1.Recordset.CancelUpdate
 txtSpecification.Text = ""
 txtCharge.Text = ""
 txtReceived.Text = ""
 txtBalance.Text = ""
 txtRefDoctor.Text = ""
 'cboTest.Text = ""
 txtTreatNo.Text = ""
 txtReceived.Locked = True
' cboTest.Locked = True
 txtDate.Locked = True
 txtPId.Locked = True
 txtPId.Enabled = False
 txtSpecification.Locked = True
 txtCharge.Locked = True
 txtRefDoctor.Locked = True
 txtTreatNo.Locked = True
 txtBalance.Locked = True

 cmdSave.Enabled = False
 cmdCancel.Enabled = False
 cmdNew.Enabled = True
 CmdExit.Enabled = True
 cmdUpdate.Enabled = True
 cmdNew.SetFocus
End If

 txtDate.Text = ""
 txtPId.Text = ""

End Sub

Private Sub cmdUpdate_Click()
 Width = 9435
 cmdNew.Enabled = False
 CmdExit.Enabled = False
 cmdUpdate.Enabled = False
 txtId.Text = ""
 txtTNo.Text = ""
 txtId.SetFocus
 Frame2.Visible = True
End Sub

Private Sub cmdFind_Click()
'txtPay.Enabled = False
'txtBal.Enabled = False
If txtId.Text = "" Or txtTNo.Text = "" Then
 MsgBox "PatientID and TreatmentNo is neccessary for search"
 txtId.SetFocus

Else
 Adodc2.RecordSource = "Select * from TestTransaction where PatientID= " & Val(txtId.Text) & " and TreatNo=" & Val(txtTNo.Text) & ""
 Adodc2.Refresh
 
 txtPay = Adodc2.Recordset.Fields!received
 txtBal = Adodc2.Recordset.Fields!balance
cmdNSave.Enabled = True

If Adodc2.Recordset.RecordCount = 0 Then
 MsgBox " No record found, Please check your entry"
 txtTNo.Text = ""
 txtId.Text = ""
 txtId.SetFocus
 'txtReceive = Adodc2.Recordset.Fields!
 Adodc2.RecordSource = "select * from TestTransaction"
 Adodc2.Refresh


ElseIf txtBalance.Text = "0" Then
 MsgBox "Patient has paid all the charge"
 Adodc2.RecordSource = "select * from TestTransaction"
 Adodc2.Refresh
 Width = 5325
 cmdNew.Enabled = True
 CmdExit.Enabled = True
 cmdUpdate.Enabled = True

Else
' Frame2.Visible = False
 'txtPay.SetFocus
 'txtDate.Locked = True
 'txtPId.Locked = True
 'txtSpecification.Locked = True
 'txtCharge.Locked = True
 'txtReceived.Locked = True
 'txtTreatNo.Locked = True
 'txtRefDoctor.Locked = True
 'txtBalance.Locked = True
'txtBal.Enabled = True
 'txtReceive.Text = txtReceived.Text
 'txtBal.Text = txtBalance.Text
 'Adodc2.Recordset.Update
 'txtReceived.Enabled = True
End If

End If
'txtBal.Enabled = True
'txtReceive.Enabled = True
End Sub

Private Sub cmdNSave_Click()
 If txtPay.Text = "" Or txtBal.Text = "" Or txtReceive.Text = "" Then
MsgBox "Please fill all fields"
Else
Adodc2.RecordSource = "Select * from TestTransaction where PatientID=" & Val(txtId.Text) & " and TreatNo=" & Val(txtTNo.Text) & ""
 Adodc2.Refresh
 
   Adodc2.Recordset.Fields!received = Val(txtPay) + Val(txtReceive)
   Adodc2.Recordset.Fields!balance = 0
 
 Adodc2.Recordset.Update
 
 Adodc1.Recordset.Update
 txtReceived.Enabled = True
 txtDate.Text = Date
 txtBalance.Text = txtBal.Text
 txtReceived.Text = txtReceive.Text
 Adodc1.Recordset.Update
 
 Adodc1.Recordset(5) = "0"
'Adodc1.Recordset.MoveLast
' Adodc1.Recordset.Update
 
 Frame2.Visible = True
 Width = 5160
 cmdNCancel.Enabled = False
 cmdNew.Enabled = True
 CmdExit.Enabled = True
 cmdUpdate.Enabled = True
 cmdNSave.Enabled = False
 cmdNew.SetFocus
 'Me.Hide
 
 'Form3oo.Show
 'Data1.RecordSource = "select * from TestTransaction"
 'Data1.Refresh
 End If
End Sub

Private Sub cmdNCancel_Click()
 txtPay.SetFocus
 txtReceive.Text = txtReceived.Text
 txtBal.Text = txtBalance.Text
 txtPay.Text = ""
 cmdNSave.Enabled = False
 cmdNCancel.Enabled = False
End Sub

Private Sub cmdEnd_Click()
 Frame2.Visible = True
 Width = 5160
 cmdNCancel.Enabled = False
 cmdNew.Enabled = True
 CmdExit.Enabled = True
 cmdUpdate.Enabled = True
 cmdNSave.Enabled = False
 cmdNew.SetFocus
 Adodc1.RecordSource = "select * from TestTransaction"
 'Adodc1.Refresh
End Sub

Private Sub Form_Load()

 'Width = 5160
 'Height = 8460
txtDate.Text = ""
txtPId.Text = ""
txtTreatNo.Text = ""
txtRefDoctor.Text = ""
txtTest.Text = ""
txtSpecification.Text = ""
txtCharge.Text = ""
txtReceived.Text = ""
txtBalance.Text = ""
txtId.Text = ""
txtTNo.Text = ""
'txtPay.Text = ""
'txtBal.Text = ""
'txtReceive.Text = ""
End Sub








Private Sub txtPay_Change()
'txtBal.Enabled = True
'txtReceive.Enabled = True

End Sub

Private Sub txtReceive_Click()
txtBal.Enabled = True
txtReceive.Enabled = True
End Sub

Private Sub txtTNo_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If
End Sub

Private Sub txtId_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If
End Sub

Private Sub txtId_LostFocus()

If txtId.Text <> "" Then
 Adodc3.RecordSource = "select* from AInt where PatientID=" & (txtId.Text) & ""
 Adodc3.Refresh

If Adodc3.Recordset.Fields(4) = lblTest.Caption Then
 MsgBox "Patient with this ID is already discharged."
 txtId.Text = ""
 txtId.SetFocus

Else
 txtTNo.SetFocus
End If

End If

End Sub

Private Sub txtBal_KeyPress(KeyAscii As Integer)
'If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 'KeyAscii = 0
'End If
End Sub

Private Sub txtPay_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub txtPay_LostFocus()

If Val(txtPay.Text) > Val(txtBal.Text) Then
' MsgBox "Amount more than balance is not acceptable"
 'txtPay.Text = ""
 txtPay.SetFocus

ElseIf txtPay.Text <> "" Then
 'txtReceive.Text = Val(txtReceive.Text) + Val(txtPay.Text)
 txtBal.Text = Val(txtBal.Text) - Val(txtPay.Text)
' cmdNSave.Enabled = True
' cmdNCancel.Enabled = True
' cmdNSave.SetFocus
End If

End Sub

Private Sub txtPId_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub txtPId_LostFocus()
 Adodc4.RecordSource = "select * from AInt"
 query = "select * from AInt"
 Adodc4.RecordSource = query
 Adodc4.CommandType = adCmdText
 Adodc4.Refresh
 Adodc4.Recordset.MoveFirst
 
    query = "select*from TestTransaction where PatientID=" & txtPId.Text & ""
    Adodc2.RecordSource = query
    Adodc2.CommandType = adCmdText
    Adodc2.Refresh
 
 Dim temp
 temp = txtPId.Text

 If Adodc4.Recordset.RecordCount > 0 Then
    If txtPId.Text <> "" Then
    
    query = "select * from AInt where PatientID=" & txtPId.Text & ""
    Adodc3.RecordSource = query
    Adodc3.CommandType = adCmdText
    Adodc3.Refresh

    Do While Not Adodc3.Recordset.EOF
        If Adodc3.Recordset(0) = Val(txtPId.Text) Then
              
           If Adodc3.Recordset(3) = "Discharge" Then
                MsgBox "this patient is already Discharged"
                GoTo p:
           
           Else
            txtTreatNo.Text = "1"
            Do While Not Adodc2.Recordset.EOF
            
            If Adodc2.Recordset(0) = Val(txtPId.Text) Then
                Adodc2.Refresh
                Adodc2.Recordset.AddNew
                txtPId.Text = temp
                txtPId.Enabled = False
                txtTreatNo.Text = Adodc2.Recordset.RecordCount ' + 1
                txtDate.Text = Date
                GoTo p:
            Else
                txtPId.Text = temp
'                txtRefDoctor.Text = doc
                txtPId.Enabled = False
                txtTreatNo.Text = "1"
                txtDate.Text = Date
            End If
            
            Adodc2.Recordset.MoveNext
            
            Loop
            
           End If
            
        End If
        Adodc3.Recordset.MoveNext
    Loop
    Else
    MsgBox "Please enter PId"
    End If
    Else
    MsgBox "Record is Not Available"
 End If
p:
End Sub

Private Sub txtSpecification_LostFocus()
If txtCharge.Text = "--" Then
txtCharge.Locked = False
txtCharge.Text = ""
txtCharge.SetFocus
End If
End Sub

Private Sub txtCharge_GotFocus()
'If txtCharge.Text = "--" Then
'txtCharge.Locked = False
'txtCharge.Text = ""
'txtCharge.SetFocus
'End If
End Sub

Private Sub txtCharge_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub txtCharge_LostFocus()
If txtCharge.Text = "" Then
MsgBox "What's the charge?"
'txtCharge.SetFocus
'Else
'txtReceived.SetFocus
End If
End Sub

Private Sub txtReceived_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub txtRefDoctor_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Then
 KeyAscii = 0
End If
End Sub

Private Sub txtRefDoctor_LostFocus()
 doc = (txtRefDoctor.Text)
End Sub

Private Sub txtTreatNo_Change()
 txtDate.Text = Date
End Sub

Private Sub txtReceived_LostFocus()

If (txtReceived.Text = "" And Val(txtCharge.Text) > 5000) Or (Val(txtCharge.Text) > 5000 And Val(txtReceived.Text) < (Val(txtCharge.Text) * 0.5)) Then
'Val(txtReceived.Text) < (Val(txtCharge.Text) * 0.5) Then
 MsgBox "Please pay at least 50% of charge."
 txtReceived.Text = ""
 txtReceived.SetFocus

'ElseIf Val(txtReceived.Text) > Val(txtCharge.Text) Then
 'MsgBox "Amount more than charge is not acceptable"
' txtReceived.Text = ""
' txtReceived.SetFocus

Else
 txtBalance.Text = Val(txtCharge.Text) - Val(txtReceived.Text)
 cmdSave.SetFocus
End If

End Sub

Private Sub txtTreatNo_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If
End Sub

Private Sub txtReceive_KeyPress(KeyAscii As Integer)
'If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 'KeyAscii = 0
'End If
End Sub
