VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form TransferInt 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Intimation"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   705
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   8310
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   960
      TabIndex        =   30
      Top             =   6360
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc_3 
      Height          =   330
      Left            =   840
      Top             =   6960
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
      RecordSource    =   "select * from Availability"
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
   Begin VB.TextBox txtCharge 
      DataField       =   "Charge"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4440
      TabIndex        =   29
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txtDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4440
      TabIndex        =   28
      Top             =   5280
      Width           =   735
   End
   Begin MSDataListLib.DataList lstVBedNo1 
      Bindings        =   "TransferInt.frx":0000
      Height          =   645
      Left            =   6120
      TabIndex        =   27
      Top             =   5160
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1138
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc_2 
      Height          =   330
      Left            =   2640
      Top             =   7440
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
      RecordSource    =   "select * from Availability"
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
   Begin MSAdodcLib.Adodc Adodc_1 
      Height          =   330
      Left            =   840
      Top             =   7440
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
      RecordSource    =   "select * from Availability"
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
   Begin MSDataListLib.DataCombo cboBed_Code 
      DataField       =   "ToBedCode"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   4440
      TabIndex        =   26
      Top             =   4800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   7320
      Top             =   4920
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
      RecordSource    =   "select * from Availability"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "TransferInt.frx":0014
      Height          =   1695
      Left            =   480
      TabIndex        =   25
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2990
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   6360
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
      Left            =   6360
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
      Left            =   6360
      Top             =   6600
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
      RecordSource    =   "TransferIntimation"
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
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "TransferInt.frx":0029
      Height          =   1695
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2990
      _Version        =   393216
      BackColorBkg    =   16761024
      Enabled         =   -1  'True
      Appearance      =   0
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
      Left            =   5040
      Picture         =   "TransferInt.frx":003D
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6240
      Width           =   975
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
      Left            =   2280
      Picture         =   "TransferInt.frx":047F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0C0&
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
      Height          =   855
      Left            =   3720
      Picture         =   "TransferInt.frx":08C1
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   3255
      Left            =   1200
      TabIndex        =   10
      Top             =   2880
      Width           =   5775
      Begin VB.TextBox txtBedNo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "ToBedNo"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtBedNo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "FromBedNo"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtBedCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "FromBedCode"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Date"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PatientID"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtRefDoctor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "RefDoctor"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Days at Previous Place"
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
         TabIndex        =   23
         Top             =   2400
         Width           =   2895
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
         Left            =   2280
         TabIndex        =   21
         Top             =   2880
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5880
         Y1              =   1320
         Y2              =   1320
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
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblBedNo2 
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
         Left            =   4080
         TabIndex        =   18
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblBedCode2 
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
         Left            =   2280
         TabIndex        =   17
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "To Bed"
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
         TabIndex        =   16
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Transfering From Bed "
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
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblBedNo1 
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
         Left            =   4080
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblBedCode1 
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
         Left            =   2280
         TabIndex        =   13
         Top             =   1440
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
         TabIndex        =   12
         Top             =   840
         Width           =   975
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
         Left            =   3720
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label lblFind 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Transfer  Intimation"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "TransferInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i
Dim tempsorce As Recordset
Dim query As String
 
Private Sub cboBed_Code_Change()
Dim f As String

f = cboBed_Code.BoundText
'f = 1
query = "select BedNo from Availability where BedCode='" & f & "' and Reserved='FALSE'"
Adodc_2.RecordSource = query
Adodc_2.Refresh

Set lstVBedNo1.RowSource = Adodc_2
lstVBedNo1.ListField = "BedNo"
lstVBedNo1.Refresh
End Sub

Private Sub cboBedCode_Change()
Dim find As Integer
find = cboBedCode.List(cboBedCode.ListIndex)

query = "select BedNo from Availability where BedCode='" & find & "' and Reserved='FALSE'"
Adodc.RecordSource = query
Adodc.Refresh

Set cboBedCode.RowSource = Adodc
cboBedCode.ListField = "BedNo"
cboBedCode.Refresh
End Sub

Private Sub cboBedCode_Click()

Dim a As String
Dim code As String

 lstVBedNo.Clear
 Adodc2.Recordset.MoveFirst
 code = cboBedCode.Text
 code = "BedCode = '" & code & "' And Reserved='" & False & "'"
 
With Adodc2.Recordset
     .FindFirst code
End With
  
If Adodc2.Recordset.NoMatch = False Then
 lstVBedNo.AddItem Adodc2.Recordset.Fields(2).Value
   
For i = 0 To Adodc2.Recordset.RecordCount + 1
 code = cboBedCode.Text
 code = "BedCode = '" & code & "' And Reserved='" & False & "'"
 
With Adodc2.Recordset
    .FindNext code
End With
  
If Adodc2.Recordset.NoMatch = False Then
 lstVBedNo.AddItem Adodc2.Recordset.Fields(2).Value
End If
   
Next

Exit Sub

End If

End Sub

Private Sub cmdAdd_Click()
Adodc1.Refresh
Adodc1.Recordset.MoveLast
 Adodc1.Recordset.AddNew
 txtPId.Locked = False
 txtPId.Text = txtFind.Text
 txtRefDoctor.Locked = False
 txtBedCode.Locked = False
 txtBedNo1.Locked = False
 cboBed_Code.Locked = False
 txtDays.Locked = False

 cmdSave.Enabled = True
 cmdAdd.Enabled = False
 txtDate.Enabled = True
 txtDate.Text = Date
 txtDays.Text = ""
 cmdEnd.Enabled = False
lstVBedNo1.Visible = True
' txtRefDoctor.SetFocus

End Sub

Private Sub cmdCancel_Click()
Adodc1.Refresh
Adodc1.Recordset.CancelUpdate
Adodc1.Refresh

End Sub

Private Sub cmdSave_Click()

If (txtBedCode.Text = "" Or txtBedNo1.Text = "" Or txtCharge.Text = "" Or txtDays.Text = "") Then
 MsgBox "Please fill all the information"

Else
 Adodc1.Recordset.Update
 Adodc1.Refresh
 txtDate.Enabled = False
 txtPId.Locked = True
 txtRefDoctor.Locked = True
 txtBedCode.Locked = True
 txtBedNo1.Locked = True
 'cboBedCode.Locked = True
 txtDays.Locked = True
 txtBedNo2.Locked = True
 'lstVBedNo.Visible = False
 cmdAdd.Enabled = True
 cmdSave.Enabled = False
 cmdEnd.Enabled = True
 
End If

End Sub

Private Sub cmdEnd_Click()
 txtFind.Visible = True
 cmdAdd.Visible = True

 txtBedCode.Enabled = True
 txtBedNo1.Enabled = True
 txtCharge.Enabled = True
 txtDays.Enabled = True
 txtBedNo2.Enabled = True
 'cboBedCode.Enabled = True
 TransferInt.Hide
End Sub

Private Sub DataList1_Click()
 txtBedNo2.Text = lstVBedNo.List(lstVBedNo.ListIndex)
End Sub

Private Sub DataList1_GotFocus()
If cboBedCode.Text = "" Then
 cboBedCode.SetFocus
 MsgBox "Please select bed of your choice"

ElseIf lstVBedNo.List(0) = "" Then
 cboBedCode.SetFocus
 MsgBox "Please select bed form another ward ,Because there is no vacancy in this ward"
End If
End Sub

Private Sub DataList1_LostFocus()
If txtBedNo2.Text <> "" Then

Dim reply
 reply = MsgBox("Are you sure you want to be on this bed only?", vbYesNo)

If reply = vbNo Then
 txtBedNo2.Text = ""
 lstVBedNo.SetFocus

ElseIf reply = vbYes Then
 Reservation.Show
 Reservation.Adodc1.RecordSource = "SELECT * FROM Availability WHERE BedCode = '" & TransferInt.cboBedCode.Text & "' AND BedNo='" & TransferInt.txtBedNo2.Text & "' "
 Reservation.Adodc1.Refresh
End If

End If
End Sub

Private Sub Form_Load()
Adodc1.Refresh
' Width = 8430
 'Height = 8370
 txtDate.Text = Date

query = "select BedNo from Availability where Reserved='FALSE'"
Adodc.RecordSource = query
Adodc.Refresh

Set lstVBedNo1.RowSource = Adodc
lstVBedNo1.ListField = "BedNo"
lstVBedNo1.Refresh


'Dim find As String
'find = DataCombo1.List(DataCombo1.ListIndex)

query = "select distinct BedCode from Availability where Reserved='FALSE'"
Adodc_1.RecordSource = query
Adodc_1.Refresh

Set cboBed_Code.RowSource = Adodc_1
cboBed_Code.ListField = "BedCode"
cboBed_Code.Refresh
End Sub






Private Sub lstVBedNo_Click()
End Sub



Private Sub lstVBedNo1_Click()
txtBedNo2 = lstVBedNo1.BoundText
End Sub

Private Sub txtDays_LostFocus()
Dim a As String
query = "select Charge from Availability where BedCode= '" & cboBed_Code.BoundText & "'"
Adodc_3.RecordSource = query
Adodc_3.Refresh

a = Adodc_3.Recordset.Fields!charge
txtCharge = Val(a) * Val(txtDays)
End Sub

Private Sub txtPId_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub txtFind_GotFocus()
 txtFind.Text = ""
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub txtFind_LostFocus()

If txtFind.Text <> "" Then
 'Adodc1.RecordSource = "SELECT * FROM TransferIntimation WHERE TransferIntimation.PatientID = '" & (txtFind.Text) & "'"  ' , AInt.EntryType='" & (14.Caption) & "',AInt.PatientID='" & (txtFind.Text) & "'"
 'Adodc1.Refresh
 
 Adodc3.RecordSource = "select * from AInt"
 query = "select * from AInt"
 Adodc3.RecordSource = query
 Adodc3.CommandType = adCmdText
 Adodc3.Refresh
 
 Adodc3.Recordset.MoveFirst
 
Do While Not Adodc3.Recordset.EOF
If txtFind = Adodc3.Recordset(0) Then

 txtPId = Adodc3.Recordset(0)
 'txtDate = Adodc3.Recordset(9)
 txtBedCode = Adodc3.Recordset(1)
 txtBedNo1 = Adodc3.Recordset(2)
 
 GoTo p:
 End If
 
 Adodc3.Recordset.MoveNext
Loop
p:
End If
'If Adodc1.Recordset.RecordCount = 0 Then
' TransferInt.Frame1.Enabled = False
' TransferInt.cmdAdd.Visible = False
' TransferInt.cmdSave.Visible = False
' TransferInt.cmdEnd.Visible = False
' MsgBox "There is no patient with this ID"
' txtFind.SetFocus
'Else
' Adodc1.Recordset.MoveLast
'
'If Adodc1.Recordset.Fields(6) = "--" Then
' TransferInt.Frame1.Enabled = False
' TransferInt.cmdAdd.Visible = False
' TransferInt.cmdSave.Visible = False
' TransferInt.cmdEnd.Visible = False
' MsgBox "This patient is already discharged, You can't make any transaction for him"
'
'Else
' TransferInt.Frame1.Enabled = True
' TransferInt.cmdAdd.Visible = True
' TransferInt.cmdSave.Visible = True
' TransferInt.cmdEnd.Visible = True
'End If

'End If

'End If

End Sub

Private Sub txtRefDoctor_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Then
 KeyAscii = 0
End If
End Sub

Private Sub txtRefDoctor_LostFocus()

If txtBedCode.Text = "" Then
 txtBedCode.Text = MSFlexGrid1.TextMatrix(Adodc1.Recordset.RecordCount, 6)
 txtBedNo1.Text = MSFlexGrid1.TextMatrix(Adodc1.Recordset.RecordCount, 7)

If txtBedNo1.Enabled = True Then
 txtBedNo1.SetFocus
End If

End If

End Sub

Private Sub txtBedCode_LostFocus()

'If Not (txtBedCode.Text = "AGA" Or txtBedCode.Text = "AGB" Or txtBedCode.Text = "AGC" Or txtBedCode.Text = "ALA" Or txtBedCode.Text = "ALB" Or txtBedCode.Text = "ALC" Or txtBedCode.Text = "ACA" Or txtBedCode.Text = "ACB" Or txtBedCode.Text = "ACC" Or txtBedCode.Text = "BGA" Or txtBedCode.Text = "BGB" Or txtBedCode.Text = "BGC" Or txtBedCode.Text = "BLA" Or txtBedCode.Text = "BLB" Or txtBedCode.Text = "BLC" Or txtBedCode.Text = "BCA" Or txtBedCode.Text = "BCB" Or txtBedCode.Text = "BCC" Or txtBedCode.Text = "CGA" Or txtBedCode.Text = "CGB" Or txtBedCode.Text = "CGC" Or txtBedCode.Text = "CLA" Or txtBedCode.Text = "CLB" Or txtBedCode.Text = "CLC" Or txtBedCode.Text = "CCA" Or txtBedCode.Text = "CCB" Or txtBedCode.Text = "CCC" Or txtBedCode.Text = "DGA" Or txtBedCode.Text = "DGB" Or txtBedCode.Text = "DGC" Or txtBedCode.Text = "DLA" Or txtBedCode.Text = "DLB" Or txtBedCode.Text = "DLC" Or txtBedCode.Text = "DCA" Or txtBedCode.Text = "DCB" Or txtBedCode.Text = "DCC") Then
' txtBedCode.Text = ""
' MsgBox "Entered bed code was wrong"
'End If

End Sub

Private Sub txtBedNo1_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub txtBedNo1_LostFocus()

'If txtBedNo1.Text <> "" Then
' Reservation2.Show
' Reservation2.Adodc1.RecordSource = "SELECT * FROM Availability WHERE BedCode = '" & TransferInt.txtBedCode.Text & "' AND BedNo='" & TransferInt.txtBedNo1.Text & "' "
' Reservation2.Adodc1.Refresh
'End If

End Sub

Private Sub txtCharge_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

