VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Reservation 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make your bed reserved here"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   3990
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1920
      Top             =   4200
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
   Begin VB.ComboBox cboReserved 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Reserved"
      DataSource      =   "Data1"
      Height          =   315
      ItemData        =   "Reservation.frx":0000
      Left            =   1680
      List            =   "Reservation.frx":000A
      TabIndex        =   0
      Text            =   "Reservation"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00FFC0FF&
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
      Picture         =   "Reservation.frx":001B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "BedNo"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "BedCode"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Charge"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Gender"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Availability"
      Connect         =   "Access"
      DatabaseName    =   "..\Hospital Management System\Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
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
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Availability"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "******************************************************************"
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
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label lblReserved 
      BackColor       =   &H00C0E0FF&
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
      Left            =   720
      TabIndex        =   11
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Please Make reservation True for selected Bed"
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
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblNumber 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Number"
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
      Left            =   720
      TabIndex        =   9
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblCode 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Code"
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
      Left            =   720
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblRate 
      BackColor       =   &H00C0E0FF&
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
      Left            =   720
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblType 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Type"
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
      Left            =   720
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "Reservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboReserved_LostFocus()

If cboReserved.Text <> "TRUE" Then
 MsgBox "Please Reserve the Bed"
 cboReserved.SetFocus

Else
 Adodc1.Recordset.Edit
 Adodc1.Recordset.Update
End If

End Sub

Private Sub cmdEnd_Click()
If cboReserved.Text = "FALSE" Then
 MsgBox "Please Reserve the bed"
 Reservation.cboReserved.SetFocus
Else
 Reservation.Hide
 End If
End Sub

Private Sub Form_Load()
' Width = 3900
 'Height = 6240
 Adodc1.Refresh
End Sub

Private Sub Form_LostFocus()

If cboReserved.Text = "FALSE" Then
 MsgBox "Please Reserve the bed"
 Reservation.cboReserved.SetFocus
End If

End Sub

