VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Availability 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Availability"
   ClientHeight    =   8490
   ClientLeft      =   -60
   ClientTop       =   600
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboBedcode 
      DataField       =   "BedCode"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Availability.frx":0000
      Left            =   10200
      List            =   "Availability.frx":0058
      TabIndex        =   38
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Last"
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
      Picture         =   "Availability.frx":00E8
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Previous"
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
      Left            =   9120
      Picture         =   "Availability.frx":052A
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Next"
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
      Picture         =   "Availability.frx":096C
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C0C0&
      Caption         =   "First"
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
      Left            =   9120
      Picture         =   "Availability.frx":0DAE
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6720
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   375
      Left            =   360
      Top             =   5760
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Availability.frx":11F0
      Height          =   4215
      Left            =   2280
      TabIndex        =   33
      Top             =   1080
      Width           =   6615
      _ExtentX        =   11668
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9480
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   3495
      Left            =   0
      TabIndex        =   26
      Top             =   960
      Width           =   2175
      Begin VB.CommandButton cmdCWard 
         BackColor       =   &H00C0C0C0&
         Caption         =   """C"" Ward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdBWard 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   """B"" Ward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWard 
         BackColor       =   &H00C0C0C0&
         Caption         =   """A"" Ward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblA 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Availability of Beds in"
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
         TabIndex        =   32
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Availability of Beds in"
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
         TabIndex        =   31
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblC 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Availability of Beds in"
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
         TabIndex        =   30
         Top             =   2520
         Width           =   1935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Availability.frx":1205
      Height          =   4215
      Left            =   2280
      TabIndex        =   25
      Top             =   1080
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7435
      _Version        =   393216
      BackColor       =   16777215
      BackColorBkg    =   16761024
      Appearance      =   0
   End
   Begin VB.CommandButton cmdVacancy 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Vacancy"
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
      Left            =   360
      Picture         =   "Availability.frx":1219
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4680
      Width           =   1575
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
      Left            =   10440
      Picture         =   "Availability.frx":165B
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00808080&
      Caption         =   "Modify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "BedNo"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10200
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
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
      Left            =   9120
      Picture         =   "Availability.frx":1A9D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
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
      Left            =   9120
      Picture         =   "Availability.frx":1EDF
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0C0C0&
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
      Left            =   10440
      Picture         =   "Availability.frx":2321
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Charge"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10200
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtReserved 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Reserved"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10200
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ComboBox cboGender 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Gender"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Availability.frx":2763
      Left            =   10200
      List            =   "Availability.frx":2770
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboWard 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Wards"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Availability.frx":278D
      Left            =   10200
      List            =   "Availability.frx":279A
      TabIndex        =   4
      Top             =   840
      Width           =   1215
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
      Height          =   975
      Left            =   7680
      Picture         =   "Availability.frx":27A7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "A"
      Height          =   135
      Left            =   3120
      TabIndex        =   23
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label16 
      Caption         =   "D"
      Height          =   135
      Left            =   4800
      TabIndex        =   22
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label15 
      Caption         =   "C"
      Height          =   135
      Left            =   3960
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "B"
      Height          =   135
      Left            =   3480
      TabIndex        =   20
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"Availability.frx":2BE9
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   19
      Top             =   480
      Width           =   11535
   End
   Begin VB.Label lblCode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BedCode"
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
      Left            =   9240
      TabIndex        =   17
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblNumber 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BedNo"
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
      Left            =   9240
      TabIndex        =   16
      Top             =   2640
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   9000
      X2              =   9000
      Y1              =   480
      Y2              =   8520
   End
   Begin VB.Label lblWardDetails 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ward Details"
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
      Left            =   3480
      TabIndex        =   15
      Top             =   120
      Width           =   1695
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
      Left            =   9240
      TabIndex        =   14
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblWard 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ward"
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
      Left            =   9240
      TabIndex        =   13
      Top             =   840
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
      Left            =   9240
      TabIndex        =   11
      Top             =   3240
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
      Left            =   9240
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "Availability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim query As String

'Private Sub Adodc1_FieldChangeComplete(ByVal cFields As Long, Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'Dim reply, response

'If txtCode.AdodcChanged Or txtNumber.AdodcChanged Or txtRate.AdodcChanged Or cboWard.AdodcChanged Or txtReserved.AdodcChanged Or cboGender.AdodcChanged Then
 '  reply = MsgBox("The record has been change.save?", vbYesNo)

'If reply = vbNo Then
 'Save = Fals

'If reply = vbYes Then
 'Save = True

'End If

'End If

'End If
'End Sub


Private Sub cmdCancel_Click()
Adodc1.Refresh
 Adodc1.Recordset.CancelUpdate
 cboWard.Enabled = False
 cboGender.Enabled = False
 'txtCode.Enabled = False
 txtNumber.Enabled = False
 txtRate.Enabled = False
 txtReserved.Text = "FALSE"

 cmdCancel.Enabled = False
 cmdSave.Enabled = False
 cmdUpdate.Enabled = True
 cmdOk.Enabled = True
 cmdAdd.Enabled = True
 Command6.Enabled = True
 cmdEnd.Enabled = True
 cmdAWard.Enabled = True
 cmdBWard.Enabled = True
 cmdCWard.Enabled = True
 'cmdCancel1.Enabled = True
 cmdVacancy.Enabled = True

 Adodc1.Refresh

End Sub

Private Sub cmdCWard_Click()
 lblNumber.Caption = "C"
 query = "select * from availability where Wards='C' and Reserved='FALSE'"
Adodc.RecordSource = query
Adodc.Refresh

Set DataGrid1.DataSource = Adodc
DataGrid1.Refresh
 'Adodc1.RecordSource = "SELECT * FROM Availability WHERE Wards = '" & lblNumber.Caption & " '"
 'Adodc1.Refresh
End Sub

Private Sub cmdCancel1_Click()
 lblNumber.Caption = "D"
 Adodc1.RecordSource = "SELECT * FROM Availability WHERE Wards = '" & lblNumber.Caption & " '"
 Adodc1.Refresh
End Sub

Private Sub cmdFirst_Click()
cboGender.Enabled = True
 'txtCode.Enabled = True
 txtNumber.Enabled = True
 txtReserved.Enabled = True
 txtRate.Enabled = True
 cboWard.Enabled = True
 cboWard.SetFocus
 
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
cboGender.Enabled = True
 'txtCode.Enabled = True
 txtNumber.Enabled = True
 txtReserved.Enabled = True
 txtRate.Enabled = True
 cboWard.Enabled = True
 cboWard.SetFocus
Adodc1.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
cboGender.Enabled = True
 'txtCode.Enabled = True
 txtNumber.Enabled = True
 txtReserved.Enabled = True
 txtRate.Enabled = True
 cboWard.Enabled = True
 cboWard.SetFocus
If Adodc1.Recordset.EOF = True Then
MsgBox "No More Record"
Adodc1.Recordset.MoveFirst
Else
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub cmdPrevious_Click()
cboGender.Enabled = True
 'txtCode.Enabled = True
 txtNumber.Enabled = True
 txtReserved.Enabled = True
 txtRate.Enabled = True
 cboWard.Enabled = True
 cboWard.SetFocus
 If Adodc1.Recordset.BOF = True Then
MsgBox "No More Record"
Adodc1.Recordset.MoveLast
Else
Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub cmdVacancy_Click()
 Vacancy.Show
End Sub

Private Sub cmdEnd_Click()
 Width = 9075
 Availability.Hide
 
End Sub

Private Sub cmdAWard_Click()
 lblNumber.Caption = "A"
query = "select * from availability where Wards='A' and Reserved='FALSE'"
Adodc.RecordSource = query
Adodc.Refresh

Set DataGrid1.DataSource = Adodc
DataGrid1.Refresh

'Set DataList1.RowSource = Adodc
'DataList1.ListField = "BedNo"
'DataList1.Refresh

 'Adodc1.RecordSource = "SELECT * FROM Availability WHERE Wards = '" & lblNumber.Caption & " '"
 'Adodc1.Refresh
End Sub

Private Sub cmdBWard_Click()
 lblNumber.Caption = "B"
query = "select * from availability where Wards='B' and Reserved='FALSE'"
Adodc.RecordSource = query
Adodc.Refresh

Set DataGrid1.DataSource = Adodc
DataGrid1.Refresh

' Adodc1.RecordSource = "SELECT * FROM Availability WHERE Wards = '" & lblNumber.Caption & " '"
' Adodc1.Refresh
End Sub

Private Sub cmdAdd_Click()
 cboGender.Enabled = True
 'txtCode.Enabled = True
 txtNumber.Enabled = True
 txtReserved.Enabled = True
 txtRate.Enabled = True
 cboWard.Enabled = True
 cboWard.SetFocus
 'txtCode.Locked = False

 cmdCancel.Enabled = True
 cmdSave.Enabled = True
 cmdUpdate.Enabled = False
 Command6.Enabled = False
 cmdEnd.Enabled = False
 cmdOk.Enabled = False
 cmdAdd.Enabled = False
 cmdAWard.Enabled = False
 cmdBWard.Enabled = False
 cmdCWard.Enabled = False
' cmdCancel1.Enabled = False
 cmdVacancy.Enabled = False

 'Adodc1.RecordSource = "SELECT * FROM Availability" ' WHERE Wards ='" & Label17.Caption & "' AND Wards='" & Label6.Caption & "'AND Wards='" & Label15.Caption & "'AND Wards='" & Label16.Caption & "'"
 'Adodc1.Refresh
' Adodc1.Recordset.MoveLast
 Adodc1.Refresh
 Adodc1.Recordset.AddNew
 'txtNumber.Text = Adodc1.Recordset.RecordCount + 1
 'txtReserved.Text = "FALSE"

End Sub

Private Sub cmdSave_Click()

If cboGender.Text = "" Or cboWard.Text = "" Or cboBedCode.Text = "" Or txtNumber.Text = "" Or txtReserved.Text = "" Or txtRate.Text = "" Then
 MsgBox "Please Fill all the fields"
 cboWard.SetFocus
Else

On Error Resume Next

 Adodc1.Recordset.Update
 cboWard.Enabled = False
 cboGender.Enabled = False
 'txtCode.Enabled = False
 txtNumber.Enabled = False
 txtRate.Enabled = False
 txtReserved.Text = "FALSE"

 cmdCancel.Enabled = False
 cmdSave.Enabled = False
 cmdUpdate.Enabled = True
 cmdOk.Enabled = True
 cmdAdd.Enabled = True
 Command6.Enabled = True
 cmdEnd.Enabled = True
 cmdAWard.Enabled = True
 cmdBWard.Enabled = True
 cmdCWard.Enabled = True
 'cmdCancel1.Enabled = True
 cmdVacancy.Enabled = True

 Adodc1.Refresh
End If

End Sub

Private Sub cmdUpdate_Click()

If cboGender.Text = "" Or cboWard.Text = "" Or cboBedCode.Text = "" Or txtNumber.Text = "" Or txtReserved.Text = "" Then
 MsgBox "Please select the bed you want to update"
Else
 cboGender.Enabled = True
 'txtCode.Enabled = True
 txtNumber.Enabled = True
 txtRate.Enabled = True
 cboWard.Enabled = True
 cboWard.SetFocus

 cmdSave.Enabled = True
 cmdUpdate.Enabled = False
 cmdCancel.Enabled = True
 cmdSave.Enabled = True
 cmdAdd.Enabled = False
 Command6.Enabled = False
 cmdEnd.Enabled = False
 cmdOk.Enabled = False
 cmdAdd.Enabled = False
 cmdAWard.Enabled = False
 cmdBWard.Enabled = False
 cmdCWard.Enabled = False
 'cmdCancel1.Enabled = False
 cmdVacancy.Enabled = False
 cmdOk.Enabled = False
 Adodc1.Recordset.Update
End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command6_Click()
Width = 11715
End Sub

Private Sub cmdok_Click()
'Width = 9105
Adodc1.Refresh
End Sub



Private Sub Form_Load()
 
 cboWard.Text = ""
 cboGender.Text = ""
 cboBedCode.Text = ""
 txtNumber.Text = ""
 txtReserved.Text = ""
 txtRate.Text = ""
 
 'Width = 9075
 'Height = 7185
 'Adodc1.Refresh
End Sub



Private Sub txtCode_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = 0
End If
End Sub
