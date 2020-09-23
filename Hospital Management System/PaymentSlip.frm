VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form PaymentSlip 
   Caption         =   "Payment Slip"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command12 
      Caption         =   "LAST"
      Height          =   495
      Left            =   4320
      TabIndex        =   20
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   1680
      TabIndex        =   18
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "FIRST"
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee Payment Slip"
      Height          =   5055
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5175
      Begin VB.TextBox Text1 
         DataField       =   "EmpID"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3000
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         DataField       =   "EmpName"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3000
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         DataField       =   "PPM"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         DataField       =   "TPD"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         DataField       =   "TAD"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         DataField       =   "TPG"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         DataField       =   "Other"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3000
         TabIndex        =   2
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         DataField       =   "Date"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emp ID"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emp Name"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Payment Per Month"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Present Days "
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Absent Days"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   3960
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total Payment Given"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   4560
         Width           =   1485
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Other Payment"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2760
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   345
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   6240
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\Hms.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\Hms.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Emp_PDetails"
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
End
Attribute VB_Name = "PaymentSlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command10_Click()
If Adodc1.Recordset.BOF Then
   MsgBox "Sorry!no previous record is there", vbInformation, "HMS"
Else
Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command11_Click()
If Adodc1.Recordset.EOF Then
MsgBox "Sorry!no more record is there", vbInformation, "HMS"
Else
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub Command12_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command9_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Form_Load()
MsgBox "YOU CAN'T MAKE CHANGES HERE", vbInformation, "HMS"
End Sub
