VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Contact 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Contact"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin MSAdodcLib.Adodc Adodc_1 
      Height          =   330
      Left            =   5040
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
      RecordSource    =   "select * from Contact"
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
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   8280
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
      RecordSource    =   "select * from Contact"
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
      Left            =   3960
      Picture         =   "Contact.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6120
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
      Left            =   7320
      Picture         =   "Contact.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6120
      Width           =   1095
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
      Left            =   5640
      Picture         =   "Contact.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6120
      Width           =   1095
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
      Left            =   9000
      Picture         =   "Contact.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6120
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Contact.frx":1108
      Height          =   1095
      Left            =   2640
      TabIndex        =   30
      Top             =   7320
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1931
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
      Left            =   6720
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
      RecordSource    =   "Contact"
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
      Height          =   5655
      Left            =   1680
      TabIndex        =   8
      Top             =   240
      Width           =   8895
      Begin VB.TextBox txtCategory1 
         Appearance      =   0  'Flat
         DataField       =   "CATAGORY"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1320
         TabIndex        =   37
         Top             =   2520
         Width           =   3015
      End
      Begin MSDataListLib.DataCombo cboCategory2 
         Height          =   315
         Left            =   7080
         TabIndex        =   36
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataList lstNames 
         Height          =   1620
         Left            =   6600
         TabIndex        =   35
         Top             =   3480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2858
         _Version        =   393216
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "NAME"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "ADDRESS"
         DataSource      =   "Adodc1"
         Height          =   885
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox txtPhNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PHONE NUMBER"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"Contact.frx":111D
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
         TabIndex        =   16
         Top             =   960
         Width           =   8655
      End
      Begin VB.Label lblContactDiary 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contact Diary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NAME"
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
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblCategory1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CATEGORY"
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
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ADDRESS"
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
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label lblPhNo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PHONE NO"
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
         TabIndex        =   11
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   5280
         X2              =   5280
         Y1              =   1080
         Y2              =   5760
      End
      Begin VB.Label lblCategory2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "CATEGORY"
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
         Left            =   5520
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblNames 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "LIST OF NAMES"
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
         TabIndex        =   9
         Top             =   3240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      Picture         =   "Contact.frx":11C9
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SAVE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      Picture         =   "Contact.frx":160B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      Picture         =   "Contact.frx":1A4D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      Picture         =   "Contact.frx":1E8F
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Left            =   10680
      Picture         =   "Contact.frx":22D1
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   1095
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
      Left            =   345
      TabIndex        =   29
      Top             =   480
      Width           =   210
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
      Left            =   345
      TabIndex        =   28
      Top             =   840
      Width           =   225
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
      Left            =   360
      TabIndex        =   27
      Top             =   1200
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
      Left            =   360
      TabIndex        =   26
      Top             =   1560
      Width           =   180
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
      Left            =   360
      TabIndex        =   25
      Top             =   1920
      Width           =   195
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
      Left            =   345
      TabIndex        =   24
      Top             =   2880
      Width           =   210
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
      Left            =   345
      TabIndex        =   23
      Top             =   3240
      Width           =   225
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
      Left            =   375
      TabIndex        =   22
      Top             =   3600
      Width           =   165
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
      Left            =   360
      TabIndex        =   21
      Top             =   3960
      Width           =   180
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
      Left            =   405
      TabIndex        =   20
      Top             =   4320
      Width           =   105
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
      Left            =   360
      TabIndex        =   19
      Top             =   4680
      Width           =   195
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
      Left            =   360
      TabIndex        =   18
      Top             =   5040
      Width           =   195
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
      Left            =   360
      TabIndex        =   17
      Top             =   5400
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "Contact.frx":2713
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1530
   End
End
Attribute VB_Name = "Contact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i
Dim query As String


'Private Sub Adodc1_FieldChangeComplete(ByVal cFields As Long, Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'Dim reply, response

'If txtName.DataChanged Or cboCategory1.DataChanged Or txtAddress.DataChanged Or txtPhNo.DataChanged Then
 '  reply = MsgBox("Are you sure you want to save?", vbYesNo)

'If reply = vbNo Then
 '  Save = False

'If reply = vbYes Then
 '  Save = True

'End If

'End If

'End If
'End Sub


'Private Sub cboCategory1_LostFocus()

'If cboCategory1.Text = "" Then
' MsgBox "Please fill the category"
' cboCategory1.SetFocus
'End If

'End Sub

Public Sub cboCategory2_Click(Area As Integer)
query = "select distinct CATAGORY from Contact"
Adodc.RecordSource = query
Adodc.Refresh

Set cboCategory2.RowSource = Adodc
cboCategory2.ListField = "CATAGORY"
cboCategory2.Refresh


query = "select NAME from Contact where CATAGORY='" & cboCategory2.BoundText & "'"
Adodc_1.RecordSource = query
Adodc_1.Refresh

Set lstNames.RowSource = Adodc_1
lstNames.ListField = "NAME"
lstNames.Refresh
End Sub

'Private Sub cboCategory2_Change()
' lstNames.Clear

'Dim A As String

'Dim CATAGORY As String
 'Adodc1.Recordset.MoveFirst
 'CATAGORY = cboCategory2.Text
 'CATAGORY = "CATAGORY = '" & CATAGORY & "'"
 
'With Adodc1.Recordset
 '    .FindFirst CATAGORY

'End With

'If Adodc1.Recordset.NoMatch = False Then
 '  lstNames.AddItem Adodc1.Recordset.Fields(4).Value
 
'For i = 0 To Adodc1.Recordset.RecordCount + 1
 '  CATAGORY = cboCategory2.Text
  ' CATAGORY = "CATAGORY = '" & CATAGORY & "'"'
'With Adodc1.Recordset
 '    .FindNext CATAGORY

'End With

'If Adodc1.Recordset.NoMatch = False Then
 '  lstNames.AddItem Adodc1.Recordset.Fields(4).Value
  
'End If
      
'Next

'Exit Sub
'End If
'End Sub

'Public Sub cboCategory2_Click()
 'lstNames.Clear

'Dim A As String
'Dim CATAGORY As String

 'Adodc1.Recordset.MoveFirst
 'CATAGORY = cboCategory2.Text
 'CATAGORY = "CATAGORY = '" & CATAGORY & "'"
 
'With Adodc1.Recordset
 '    .FindFirst CATAGORY

'End With
  
'If Adodc1.Recordset.NoMatch = False Then
 '  lstNames.AddItem Adodc1.Recordset.Fields(4).Value
   
'For i = 0 To Adodc1.Recordset.RecordCount + 1
 '   CATAGORY = cboCategory2.Text
  '  CATAGORY = "CATAGORY = '" & CATAGORY & "'"
 
'With Adodc1.Recordset
 '    .FindNext CATAGORY

'End With
  
'If Adodc1.Recordset.NoMatch = False Then
 '  lstNames.AddItem Adodc1.Recordset.Fields(4).Value
  
'End If
   
'Next

'Exit Sub

'End If

'End Sub

Private Sub cmdAdd_Click()
 Adodc1.Recordset.AddNew

 txtName.Locked = False
 txtAddress.Locked = False
 txtPhNo.Locked = False
 txtCategory1.Locked = False

 txtName.SetFocus

 cmdSave.Enabled = True
 cmdAdd.Enabled = False
 cmdUpdate.Enabled = False
 cmdDelete.Enabled = False
 cmdEnd.Enabled = False
End Sub

Private Sub cmdSave_Click()

If txtName.Text = "" Or txtAddress.Text = "" Or txtPhNo.Text = "" Or txtCategory1.Text = "" Then
 MsgBox "Please fill all the fields"
 txtName.SetFocus

Else
 Adodc1.Recordset.Update
 txtName.Locked = True
 txtAddress.Locked = True
 txtPhNo.Locked = True
 txtCategory1.Locked = True

 cmdSave.Enabled = False
 cmdAdd.Enabled = True
 cmdUpdate.Enabled = True
 cmdDelete.Enabled = True
 cmdEnd.Enabled = True
End If

End Sub

Private Sub cmdUpdate_Click()
 MsgBox "Please make your changes here"
 Adodc1.Recordset.Update
 txtName.Locked = False
 txtAddress.Locked = False
 txtPhNo.Locked = False
 txtCategory1.Locked = False

 txtName.SetFocus
 cmdSave.Enabled = True
 cmdAdd.Enabled = False
 cmdUpdate.Enabled = False
 cmdDelete.Enabled = False
 cmdEnd.Enabled = False
End Sub

Private Sub cmdDelete_Click()

On Error Resume Next

Dim reply

 reply = MsgBox("Do you wish to delete the record", vbYesNo)

If reply = vbYes Then
   Adodc1.Recordset.Delete
   MsgBox "your record is deleted"

If Not Adodc1.Recordset.EOF Then
   Adodc1.Recordset.MoveNext

ElseIf Not Adodc1.Recordset.BOF Then
   Adodc1.Recordset.MovePrevious

Else
   MsgBox "This was the only record in the table"

End If

End If

End Sub

Private Sub cmdEnd_Click()
 Contact.Hide
End Sub


Private Sub Form_Load()
txtName.Text = ""
txtCategory1.Text = ""
txtAddress.Text = ""
txtPhNo.Text = ""
cboCategory2.Text = ""
lstNames.Text = ""
 
 Width = 12000
 Height = 9000

query = "select distinct CATAGORY from Contact"
Adodc.RecordSource = query
Adodc.Refresh

Set cboCategory2.RowSource = Adodc
cboCategory2.ListField = "CATAGORY"
cboCategory2.Refresh
End Sub

'Private Sub lstNames_Click()

'Dim name As String

 'name = "NAME='" & lstNames.List(lstNames.ListIndex) & "'" & "And CATAGORY ='" & cboCategory2.List(cboCategory2.ListIndex) & "'"
 'Adodc1.Recordset.FindFirst name

'End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Then
 KeyAscii = 0
End If
End Sub

Private Sub txtName_LostFocus()

If txtName.Text = "" Then
 txtName.SetFocus
 MsgBox "Please Fill The Name"

ElseIf IsNumeric(txtName.Text) Then
 MsgBox "Please Enter Your Name"
 txtName.Text = ""
 txtName.SetFocus

End If

End Sub

Private Sub txtAddress_LostFocus()

If txtAddress.Text = "" Then
 txtAddress.SetFocus
 MsgBox "Please Fill The Patient's Address"

ElseIf IsNumeric(txtAddress.Text) Then
 MsgBox "Please Enter Patient's Address"
 txtAddress.Text = ""
 txtAddress.SetFocus
End If

End Sub

Private Sub txtPhNo_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub txtPhNo_LostFocus()

If txtPhNo.Text = "" Then
 txtPhNo.SetFocus
 MsgBox "PLEASE FILL THE PHONE NUMBER"
End If

End Sub
Private Sub cmdFirst_Click()
'cboGender.Enabled = True
 'txtCode.Enabled = True
 'txtNumber.Enabled = True
 'txtReserved.Enabled = True
 'txtRate.Enabled = True
 'cboWard.Enabled = True
 'cboWard.SetFocus
 
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
'cboGender.Enabled = True
 'txtCode.Enabled = True
 'txtNumber.Enabled = True
 'txtReserved.Enabled = True
 'txtRate.Enabled = True
 'cboWard.Enabled = True
 'cboWard.SetFocus
Adodc1.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
'cboGender.Enabled = True
 'txtCode.Enabled = True
 'txtNumber.Enabled = True
 'txtReserved.Enabled = True
 'txtRate.Enabled = True
 'cboWard.Enabled = True
 'cboWard.SetFocus
If Adodc1.Recordset.EOF = True Then
MsgBox "No More Record"
Adodc1.Recordset.MoveFirst
Else
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub cmdPrevious_Click()
'cboGender.Enabled = True
 'txtCode.Enabled = True
 'txtNumber.Enabled = True
 'txtReserved.Enabled = True
 'txtRate.Enabled = True
 'cboWard.Enabled = True
 'cboWard.SetFocus
 If Adodc1.Recordset.BOF = True Then
MsgBox "No More Record"
Adodc1.Recordset.MoveLast
Else
Adodc1.Recordset.MovePrevious
End If
End Sub

