VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Emp_PaymentDetails 
   Caption         =   "Emp_PaymentDetails"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Caption         =   "Employee Payment Details"
      Height          =   6255
      Left            =   480
      TabIndex        =   11
      Top             =   240
      Width           =   4935
      Begin VB.TextBox Text10 
         DataField       =   "Date"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         TabIndex        =   30
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         DataField       =   "Other"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         TabIndex        =   29
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         DataField       =   "TPG"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         TabIndex        =   19
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         DataField       =   "TAD"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         TabIndex        =   18
         Top             =   5040
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         DataField       =   "TPD"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         TabIndex        =   17
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         DataField       =   "PPD"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         TabIndex        =   16
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         DataField       =   "WDM"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         TabIndex        =   15
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         DataField       =   "PPM"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         TabIndex        =   14
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         DataField       =   "EmpName"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         TabIndex        =   13
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         DataField       =   "EmpID"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Other Payment"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   2760
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total Payment Given"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   5760
         Width           =   1485
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Absent Days"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   5160
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Present Days "
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Payment Per Day "
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   3960
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Working Days In Month"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   3360
         Width           =   1680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Payment Per Month"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emp Name"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emp ID"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   525
      End
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   3240
      Top             =   7800
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\Hms.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\Hms.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Emp_Details"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   495
      Left            =   7200
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ADD"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "FIRST"
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "LAST"
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1680
      Top             =   7800
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
Attribute VB_Name = "Emp_PaymentDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim query As String

Private Sub Command1_Click()
Text8 = Val(Text6.Text) * Val(Text5.Text)
End Sub

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

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()

Adodc1.Recordset.AddNew
Text10.Text = Date
Text1.SetFocus
'If Text1.Text = "" Then
'MsgBox "Please Enter Correct Employee ID"
'Text1.SetFocus
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command8_Click()
End
End Sub

Private Sub Command9_Click()
Adodc1.Recordset.MoveFirst
End Sub








Private Sub Form_Load()
'Text10.Text = Date
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub Text1_LostFocus()
If Text1.Text <> "" Then
'query = "select * from Emp_PDetails where EmpID='" & Text1.Text & "'"
'Adodc1.RecordSource = query
'Adodc1.CommandType = adCmdText
'Adodc1.Refresh
'If Adodc1.Recordset.EOF Then

'Adodc1.Recordset(1) = Text2.Text
'Val(Text2.Text) = Adodc1.Recordset(1)
'End If
'End If
'Adodc1.Refresh

Adodc.RecordSource = "Select * from Emp_Details where RecordNo= " & Text1.Text '& "'"" "
Adodc.CommandType = adCmdText
Adodc.Refresh
Text2.Text = Adodc.Recordset.Fields(1)
Text3.Text = Adodc.Recordset.Fields(8)
End If
End Sub



Private Sub Text4_LostFocus()
Text5 = Val(Text3.Text) / Val(Text4.Text)
End Sub





Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub Text7_LostFocus()
Text8 = Val(Text6.Text) * Val(Text5.Text) + Val(Text9.Text)
End Sub

