VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form AdmissionInt 
   BackColor       =   &H00C0C0C0&
   Caption         =   "AdmissionInt"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "AdmissionInt.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc_1 
      Height          =   330
      Left            =   3120
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      RecordSource    =   "select * from availability"
      Caption         =   ""
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
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   375
      Left            =   5520
      Top             =   6000
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      RecordSource    =   "select * from availability"
      Caption         =   "Adodc3"
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
      Left            =   9600
      Top             =   5880
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
      RecordSource    =   "Availability"
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
      Left            =   8160
      Top             =   5880
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
      RecordSource    =   "AInt"
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
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "AdmissionInt.frx":0442
      Height          =   1935
      Left            =   2400
      OleObjectBlob   =   "AdmissionInt.frx":0456
      TabIndex        =   42
      Top             =   6480
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   1680
      TabIndex        =   13
      Top             =   0
      Width           =   6855
      Begin MSDataListLib.DataList txtRate 
         Height          =   255
         Left            =   1440
         TabIndex        =   47
         Top             =   4800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
      End
      Begin VB.ComboBox cboReserved 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Reserved"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "AdmissionInt.frx":0E29
         Left            =   1440
         List            =   "AdmissionInt.frx":0E33
         TabIndex        =   44
         Top             =   5400
         Width           =   1215
      End
      Begin MSDataListLib.DataList lstVBedNo 
         Bindings        =   "AdmissionInt.frx":0E44
         Height          =   1035
         Left            =   5520
         TabIndex        =   43
         Top             =   3360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1826
         _Version        =   393216
      End
      Begin VB.ComboBox cboGender 
         DataField       =   "Gender"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AdmissionInt.frx":0E58
         Left            =   5280
         List            =   "AdmissionInt.frx":0E62
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Name"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Address"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   735
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox txtPhNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PhoneNo"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Age"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   5760
         TabIndex        =   4
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PatientID"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox cboBedCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "BedCode"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   315
         ItemData        =   "AdmissionInt.frx":0E73
         Left            =   1560
         List            =   "AdmissionInt.frx":0E80
         TabIndex        =   7
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "AdmissionDate"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtBedNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "BedNo"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3960
         Width           =   495
      End
      Begin VB.TextBox txtEntry 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "EntryType"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Admit"
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblRate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rate"
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
         Left            =   480
         TabIndex        =   46
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label lblReserved 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reserved"
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
         Left            =   480
         TabIndex        =   45
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label lblGender 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gender"
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
         Left            =   4440
         TabIndex        =   28
         Top             =   1440
         Width           =   735
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
         Left            =   480
         TabIndex        =   27
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Name"
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
         Left            =   480
         TabIndex        =   26
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Address"
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
         Left            =   480
         TabIndex        =   25
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblPhNo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Phone No"
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
         Left            =   480
         TabIndex        =   24
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Age"
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
         Left            =   5280
         TabIndex        =   23
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblEntry 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status Of Entry"
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
         Left            =   5280
         TabIndex        =   22
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblBedCode 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bed Code"
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
         Left            =   480
         TabIndex        =   21
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label lblBedNo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bed No"
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
         Left            =   2760
         TabIndex        =   20
         Top             =   3960
         Width           =   735
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
         Left            =   480
         TabIndex        =   19
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblAInt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Admission Intimation"
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
         Left            =   1680
         TabIndex        =   18
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "********************************************************************************************************************"
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
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   6615
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00808080&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
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
      Height          =   495
      Left            =   4680
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnd 
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
      Left            =   10440
      Picture         =   "AdmissionInt.frx":0E9D
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   1095
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
      Left            =   8880
      Picture         =   "AdmissionInt.frx":12DF
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
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
      Left            =   10440
      Picture         =   "AdmissionInt.frx":1721
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add"
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
      Left            =   8880
      Picture         =   "AdmissionInt.frx":1B63
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   600
      TabIndex        =   41
      Top             =   4680
      Width           =   180
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   600
      TabIndex        =   40
      Top             =   4320
      Width           =   195
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   600
      TabIndex        =   39
      Top             =   3960
      Width           =   195
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   645
      TabIndex        =   38
      Top             =   3600
      Width           =   105
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   600
      TabIndex        =   37
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   615
      TabIndex        =   36
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   585
      TabIndex        =   35
      Top             =   2520
      Width           =   225
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   585
      TabIndex        =   34
      Top             =   2160
      Width           =   210
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   600
      TabIndex        =   33
      Top             =   1560
      Width           =   195
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   600
      TabIndex        =   32
      Top             =   1200
      Width           =   180
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   600
      TabIndex        =   31
      Top             =   840
      Width           =   195
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   585
      TabIndex        =   30
      Top             =   480
      Width           =   225
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   585
      TabIndex        =   29
      Top             =   120
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "AdmissionInt.frx":1FA5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1410
   End
End
Attribute VB_Name = "AdmissionInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i
Dim query, find As String
Dim a, b As String
Function f1() As Boolean
cmdCancel.SetFocus
End Function

Private Sub cboBedCode_Click()
find = cboBedCode.List(cboBedCode.ListIndex)

query = "select BedNo from Availability where BedCode='" & find & "' and Reserved='FALSE'"
Adodc.RecordSource = query
Adodc.Refresh

Set lstVBedNo.RowSource = Adodc
lstVBedNo.ListField = "BedNo"
lstVBedNo.Refresh
End Sub

Private Sub cboBedCode_LostFocus()

If Val(txtAge.Text) > 15 And (cboBedCode.Text = "ACA" Or cboBedCode.Text = "ACB" Or cboBedCode.Text = "ACC" Or cboBedCode.Text = "BCA" Or cboBedCode.Text = "BCB" Or cboBedCode.Text = "BCC" Or cboBedCode.Text = "CCA" Or cboBedCode.Text = "CCB" Or cboBedCode.Text = "CCC" Or cboBedCode.Text = "DCA" Or cboBedCode.Text = "DCB" Or cboBedCode.Text = "DCC") Then
 MsgBox "You can't select children's ward"
 cboBedCode.Text = ""
 cboBedCode.SetFocus
 
ElseIf Val(txtAge.Text) < 15 And (cboBedCode.Text = "AGA" Or cboBedCode.Text = "AGB" Or cboBedCode.Text = "AGC" Or cboBedCode.Text = "BGA" Or cboBedCode.Text = "BGB" Or cboBedCode.Text = "BGC" Or cboBedCode.Text = "CGA" Or cboBedCode.Text = "CGB" Or cboBedCode.Text = "CGC" Or cboBedCode.Text = "DGA" Or cboBedCode.Text = "DGB" Or cboBedCode.Text = "DGC" Or cboBedCode.Text = "ALA" Or cboBedCode.Text = "ALB" Or cboBedCode.Text = "ALC" Or cboBedCode.Text = "BLA" Or cboBedCode.Text = "BLB" Or cboBedCode.Text = "BLC" Or cboBedCode.Text = "CLA" Or cboBedCode.Text = "CLB" Or cboBedCode.Text = "CLC" Or cboBedCode.Text = "DLA" Or cboBedCode.Text = "DLB" Or cboBedCode.Text = "DLC") Then
 MsgBox "Please select children's ward"
 cboBedCode.Text = ""
 cboBedCode.SetFocus
 
 
 If (cboGender.Text = "Male" And Val(txtAge.Text) > 15) And (cboBedCode.Text <> "AGA" Or cboBedCode.Text <> "AGB" Or cboBedCode.Text <> "AGC" Or cboBedCode.Text <> "BGA" Or cboBedCode.Text <> "BGB" Or cboBedCode.Text <> "BGC" Or cboBedCode.Text <> "CGA" Or cboBedCode.Text <> "CGB" Or cboBedCode.Text <> "CGC") Then
  
MsgBox "Please select Gent's ward"
 cboBedCode.Text = ""
 cboBedCode.SetFocus
 
 ElseIf (cboGender.Text = "Female" And Val(txtAge.Text) > 15) And (cboBedCode.Text <> "ALA" Or cboBedCode.Text <> "ALB" Or cboBedCode.Text <> "ALC" Or cboBedCode.Text <> "BLA" Or cboBedCode.Text <> "BLB" Or cboBedCode.Text <> "BLC" Or cboBedCode.Text <> "CLA" Or cboBedCode.Text <> "CLB" Or cboBedCode.Text <> "CLC" Or cboBedCode.Text <> "DLA" Or cboBedCode.Text <> "DLB" Or cboBedCode.Text <> "DLC") Then
 MsgBox "Please select Ladies's ward"
 cboBedCode.Text = ""
 cboBedCode.SetFocus
End If
End If
End Sub

Private Sub cboReserved_Click()
a = MsgBox("Are you sure to reserved Bed", vbYesNo, "Noble Hospital")
If a = vbYes Then
b = "TRUE"
Else
If a = vbNo Then
b = "FALSE"
End If
End If
End Sub

Private Sub cmdAdd_Click()
query = "select BedNo from Availability where Reserved='FALSE'"
Adodc.RecordSource = query
Adodc.Refresh

Set lstVBedNo.RowSource = Adodc
lstVBedNo.ListField = "BedNo"
lstVBedNo.Refresh
Adodc1.Refresh
 Adodc1.Recordset.AddNew
 'lstVBedNo.Enabled = True
 cboBedCode.Locked = False
 txtPId.Enabled = True
 txtDate.Enabled = True
 'Text2.Enabled = True

 txtDate.Text = Date
 txtPId.Text = Adodc1.Recordset.RecordCount + 2
 cmdCancel.Enabled = True
 cmdEnd.Enabled = False
 Command4.Enabled = False
 cmdAdd.Enabled = False
 cmdSave.Enabled = True
 cboBedCode.Enabled = True
 cboGender.Enabled = True
 Adodc1.Enabled = False

 txtName.Enabled = True
 txtName.SetFocus
' lstVBedNo.Enabled = True

 txtAddress.Enabled = True
 txtPhNo.Enabled = True
 txtAge.Enabled = True
 txtBedNo.Enabled = True
 txtEntry.Enabled = True
 txtEntry.Text = "Admit"

End Sub

Private Sub cmdSave_Click()
'MsgBox cboReserved.Text

If txtDate.Text = " " Or txtPId.Text = " " Or txtName.Text = " " Or txtAddress.Text = " " Or txtPhNo.Text = " " Or cboBedCode.Text = " " Or txtBedNo.Text = " " Or txtRate.Text = " " Or cboReserved.Text = " " Or cboGender.Text = " " Or txtAge.Text = " " Then
MsgBox "Please fill all information."
End If

If cboReserved.Text = "TRUE" Or cboReserved.Text = "FALSE" Then
Else
MsgBox "Please Fill reservation True or False"
End If

If cboReserved.Text = "TRUE" Then

If (txtEntry.Text = "" Or txtPId.Text = "" Or txtAddress.Text = "" Or txtName.Text = "" Or txtDate.Text = "" Or txtPhNo.Text = "" Or txtAge.Text = "" Or txtBedNo.Text = "" Or cboBedCode.Text = "" Or cboGender.Text = "") Then
 MsgBox "Please fill all the information"
 
Else
  
 cboBedCode.Locked = True

 cmdSave.Enabled = True


query = "select * from Availability "

Adodc_1.RecordSource = query

Adodc_1.CommandType = adCmdText

Adodc_1.Refresh

Do While Not Adodc_1.Recordset.EOF
If Adodc_1.Recordset(2) = txtBedNo.Text Then
Adodc_1.Recordset(3) = "TRUE"
Adodc_1.Recordset.Update
GoTo p:
End If
Adodc_1.Recordset.MoveNext
Loop

p:

Adodc1.Recordset.Update

 txtName.Enabled = False
 txtAddress.Enabled = False
 txtPhNo.Enabled = False
 txtAge.Enabled = False
 txtBedNo.Enabled = False
 txtDate.Enabled = False
 'Text2.Locked = True
 txtDate.Enabled = False
' Text2.Enabled = False
 txtEntry.Enabled = False
 txtPId.Enabled = False

 cboBedCode.Enabled = False
 cboGender.Enabled = False
 cmdSave.Enabled = False
 cmdAdd.Enabled = True
 Command4.Enabled = True
 cmdCancel.Enabled = False
 cmdEnd.Enabled = True
 Adodc1.Enabled = True
End If
End If
End Sub

Private Sub cmdEnd_Click()
 AdmissionInt.Hide
End Sub


Private Sub Command4_Click()
Dim reply
 reply = MsgBox("Do you wish to update the selected record?", vbYesNo)

If reply = vbNo Then

 MsgBox "Select the record, You wish to update"

ElseIf txtEntry.Text = "Discharge" Then

 MsgBox "You can't update the record of discharged patient"

Else

 cmdAdd.Enabled = False
 cmdSave.Enabled = True
 Command4.Enabled = False
 cmdEnd.Enabled = False
' cmdCancel.Enabled = True
 Command6.Visible = True
 Adodc1.Enabled = False

 txtAddress.Enabled = True
 txtName.Enabled = True
 txtPhNo.Enabled = True
 txtAge.Enabled = True
 'Text2.Enabled = True
 'Adodc1.Recordset.Edit
End If

End Sub

Private Sub cmdCancel_Click()
lstVBedNo.ListField = ""
Dim ans
 ans = MsgBox("Do you want to cancel your admission?", vbYesNo)

If ans = vbYes Then

 cboBedCode.Locked = True

If (txtBedNo.Text = "") Then

 Adodc1.Recordset.CancelUpdate

 txtName.Enabled = False
 txtAddress.Enabled = False
 txtPhNo.Enabled = False
 txtAge.Enabled = False
 'Text2.Locked = True
 txtDate.Enabled = False
 'Text2.Enabled = False
 txtPId.Enabled = False
 txtBedNo.Enabled = False
 txtEntry.Enabled = False
 'lstVBedNo.Enabled = False

 cboBedCode.Enabled = False
 cboGender.Enabled = False
 cmdSave.Enabled = False
 cmdAdd.Enabled = True
 Command4.Enabled = True
 cmdCancel.Enabled = False
 cmdEnd.Enabled = True
 Adodc1.Enabled = True

Else

 MsgBox "Please cancel your reservation before canceling admission"
' Reservation2.Show
' Reservation2.Adodc1.RecordSource = "SELECT * FROM Availability WHERE BedCode = '" & AdmissionInt.cboBedcode.Text & "' AND BedNo='" & AdmissionInt.txtBedNo.Text & "' "
' Reservation2.Adodc1.Refresh
 'Text2.Text = ""
 cboBedCode.Text = ""
 txtBedNo.Text = ""
 cboGender.Text = ""
' Adodc1.Recordset.CancelUpdate

 txtName.Enabled = False
 txtAddress.Enabled = False
 txtPhNo.Enabled = False
 txtAge.Enabled = False
 'Text2.Locked = True
 txtDate.Enabled = False
 'Text2.Enabled = False
 txtPId.Enabled = False
 txtBedNo.Enabled = False
 txtEntry.Enabled = False
' lstVBedNo.Enabled = False

 cboBedCode.Enabled = False
 cboGender.Enabled = False
 cmdSave.Enabled = False
 cmdAdd.Enabled = True
 Command4.Enabled = True
 cmdCancel.Enabled = False
 cmdEnd.Enabled = True
 Adodc1.Enabled = True


End If
End If

End Sub

Private Sub Command6_Click()
'Form1.Show
 Adodc1.Recordset.Update

 txtPId.Enabled = False
 txtBedNo.Enabled = False
 txtName.Enabled = False
 txtAddress.Enabled = False
 txtPhNo.Enabled = False
 txtAge.Enabled = False
 txtBedNo.Enabled = False
 txtDate.Enabled = False
 'Text2.Locked = True
 txtDate.Enabled = False
 'Text2.Enabled = False
 txtEntry.Enabled = False
 lstVBedNo.Enabled = False

 cboBedCode.Enabled = False
 cmdSave.Enabled = False
 cmdAdd.Enabled = True
 Command4.Enabled = True
 cmdCancel.Enabled = False
 cmdEnd.Enabled = True
 Command6.Visible = False
 Adodc1.Enabled = True

End Sub



Private Sub DataList1_Click()

End Sub

Private Sub Form_Load()

txtDate.Text = " "
txtPId.Text = " "
txtName.Text = " "
txtAddress.Text = " "
txtPhNo.Text = " "
cboBedCode.Text = " "
txtBedNo.Text = " "
txtRate.Text = " "
cboReserved.Text = ""
cboGender.Text = " "
txtAge.Text = " "

'query = "select BedNo from Availability where Reserved='FALSE'"
'Adodc.RecordSource = query
'Adodc.Refresh

'Set DataList1.RowSource = Adodc
'DataList1.ListField = "BedNo"
'DataList1.Refresh

' Width = 7830
'Height = 9000
' lstVBedNo.Clear

End Sub





'Private Sub lstVBedNo_Click()
 'txtBedNo.Text = lstVBedNo.List(lstVBedNo.ListIndex)
'End Sub
'
'Private Sub lstVBedNo_GotFocus()

'If cboBedCode.Text = "" Then

 'MsgBox "Please select bed"
 'cboBedCode.SetFocus

'ElseIf lstVBedNo.List(0) = "" Then
 
 'MsgBox "Please select bed from another ward ,Because there is no vacancy in this ward"
 'cboBedCode.SetFocus

'End If

'End Sub

'Private Sub lstVBedNo_LostFocus()
 
'If cmdCancel.TabIndex = lstVBedNo.TabIndex + 1 Then

 'MsgBox ""

'Else

'If lstVBedNo.List(lstVBedNo.ListIndex) <> "" Then

'Dim reply

 'reply = MsgBox("Are you sure you want to be on this bed ?", vbYesNo)

'If reply = vbNo Then
 'txtBedNo.Text = ""
 'cboBedCode.SetFocus

'Else
 'cboBedCode.Locked = True
 'Text2.Locked = False

'Dim code As String
 'Adodc2.Recordset.MoveFirst
 'code = cboBedCode.Text
 'code = "BedCode = '" & code & "'"
 
'With Adodc2.Recordset
 '    .FindFirst code

'End With

'If Adodc2.Recordset.NoMatch = False Then
  ' Text2.Text = 10 * (Adodc2.Recordset.Fields(4).Value)
 
'For i = 0 To Adodc2.Recordset.RecordCount + 1
 '  code = cboBedCode.Text
  ' code = "BedCode = '" & code & "'"

'With Adodc2.Recordset
 '    .FindNext code

'End With

'If Adodc2.Recordset.NoMatch = False Then
   'Text2.Text = 10 * (Adodc2.Recordset.Fields(4).Value)
  
'End If
 
 'Next

'End If

'If txtBedNo.Text = "" Then
 'MsgBox "Please select the bed"
' cboBedCode.SetFocus
 'txtBedNo.Text = lstVBedNo.List(lstVBedNo.ListIndex)

'Else

 'Reservation.Show
 'Reservation.Adodc1.RecordSource = "SELECT * FROM Availability WHERE BedCode = '" & AdmissionInt.cboBedCode.Text & "' AND BedNo='" & AdmissionInt.txtBedNo.Text & "' "
 'Reservation.Adodc1.Refresh
 'lstVBedNo.Enabled = False

'End If
'End If
'End If
'End If

'End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub lstVBedNo_Click()
txtBedNo = lstVBedNo.BoundText
End Sub

Private Sub txtBedNo_Change()
If cboBedCode.Text <> "" And txtBedNo <> "" Then

query = "select Charge from Availability where BedCode='" & cboBedCode.Text & "' and BedNo= " & Val(txtBedNo.Text) & ""
Adodc_1.RecordSource = query
Adodc_1.Refresh

Set txtRate.RowSource = Adodc_1
txtRate.ListField = "Charge"
txtRate.Refresh
Else
'MsgBox ""
End If
End Sub

Private Sub txtPId_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub


Private Sub txtPId_LostFocus()
cboGender.SetFocus
End Sub



Private Sub txtAddress_LostFocus()

If IsNumeric(txtAddress.Text) Then
 txtAddress.SetFocus
 MsgBox "Please write the address"
 txtAddress.Text = ""
End If

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Then
 KeyAscii = 0
End If
End Sub

Private Sub txtName_LostFocus()

If IsNumeric(txtName.Text) Then
 txtName.SetFocus
 MsgBox "Please write patient's name"
 txtName.Text = ""
End If

End Sub

Private Sub txtPhNo_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub txtAge_LostFocus()
Dim reply

If Val(txtAge.Text) > 150 Then
 reply = MsgBox("Please check the Age of patient's you entred,Is it right?", vbYesNo)

If reply = vbNo Then
 txtAge.Text = ""
 txtAge.SetFocus
End If

End If

End Sub



