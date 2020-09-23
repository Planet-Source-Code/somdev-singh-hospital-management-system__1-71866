VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form DischargeInt 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Discharge Intimation"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   8910
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   7440
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.MDB;Persist Security Info=False"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   7440
      Top             =   3960
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
      CommandType     =   8
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
   Begin MSAdodcLib.Adodc Adodc_1 
      Height          =   330
      Left            =   7320
      Top             =   4560
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   7440
      Top             =   6360
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
      RecordSource    =   "select * from TestTransaction"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7440
      Top             =   5400
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
      Left            =   5760
      Picture         =   "DischargeInt.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDischargeEntry 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Discharge Entry"
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
      Left            =   240
      Picture         =   "DischargeInt.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   1935
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
      Left            =   4200
      Picture         =   "DischargeInt.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
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
      Left            =   2640
      Picture         =   "DischargeInt.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   6615
      Begin VB.ComboBox cboReserved 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Reserved"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "DischargeInt.frx":1108
         Left            =   5040
         List            =   "DischargeInt.frx":1112
         TabIndex        =   26
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtBedNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtBedCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Discharge"
         Top             =   5400
         Width           =   1095
      End
      Begin VB.TextBox txtPhNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "DischargeDate"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtPStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   5400
         Width           =   3135
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label lblReserved 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Discharge"
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
         TabIndex        =   27
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "************************************************************************************************************************"
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
         TabIndex        =   25
         Top             =   840
         Width           =   6375
      End
      Begin VB.Label lblDInt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Discharge Intimation"
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
         Left            =   840
         TabIndex        =   24
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lblBed 
         BackColor       =   &H00C0C0C0&
         Caption         =   "From Bed"
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
         TabIndex        =   21
         Top             =   4560
         Width           =   975
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
         Left            =   4080
         TabIndex        =   19
         Top             =   5040
         Width           =   1335
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
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   855
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
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   735
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
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblPStatus 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status Of Patient At Discharge Time"
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
         TabIndex        =   11
         Top             =   5040
         Width           =   3255
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
         Left            =   240
         TabIndex        =   10
         Top             =   2160
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
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "On Date"
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
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
   End
End
Attribute VB_Name = "DischargeInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim query, query1 As String
Dim a, b As String

Private Sub cboReserved_Click()
a = MsgBox("Are you sure to leave Bed", vbYesNo, "Noble Hospital")
If a = vbYes Then
b = "FALSE"
Else
If a = vbNo Then
b = "TRUE"
End If
End If
End Sub

Private Sub cmdEnd_Click()
 DischargeInt.Hide
 'MsgBox " Pay all your bed charge "
 'DischargeBill.Show
End Sub

Private Sub cmdSave_Click()

If txtDate.Text = "" Or txtPId.Text = "" Or txtName.Text = "" Or txtAge.Text = "" Or txtAddress.Text = "" Or txtPhNo.Text = "" Or txtPStatus.Text = "" Or txtEntry.Text = "" Or txtBedCode.Text = "" Or txtBedNo.Text = "" Then
 MsgBox "Please fill all the fields"
Else

If cboReserved.Text = "TRUE" Then

query = "select * from AInt where PatientID=" & Val(txtPId) & ""
Adodc4.RecordSource = query
'Adodc4.CommandType = adCmdText
Adodc4.Refresh

'Adodc2.Recordset.Requery
query1 = "select * from Availability where BedNo=" & Val(txtBedNo) & ""
Adodc2.RecordSource = query1
Adodc2.CommandType = adCmdText
Adodc2.Refresh


Do While Not Adodc4.Recordset.EOF
If Adodc4.Recordset("PatientID") = Val(txtPId.Text) Then
'Adodc_1.Recordset.Requery
Adodc4.Recordset("EntryType") = "Discharge"
Adodc4.Recordset.Update

Adodc2.Recordset(3) = "FALSE"
Adodc2.Recordset.Update

GoTo p:
End If
Adodc4.Recordset.MoveNext
Loop

p:
MsgBox "save successfully"
'Adodc1.Recordset.Update

End If
 

 cmdSave.Enabled = True
 cmdDischargeEntry.Enabled = True
 cmdSave.Enabled = False
 cmdCancel.Enabled = False
End If

End Sub

Private Sub cmdCancel_Click()
Dim temp
temp = MsgBox("you really don't want to take discharge?", vbYesNo)
If temp = vbYes Then
If txtBedCode.Text <> "" And txtBedNo.Text <> "" Then
Reservation.Show
Reservation.Data1.RecordSource = "select * from Availability where BedCode='" & txtBedCode.Text & "'"
'Reservation.Data1.Refresh

' txtBedCode.Text = ""
 'txtBedNo.Text = ""
 txtEntry.Text = "Admit"
 txtDate.Text = ""
 txtPStatus.Text = ""

 cmdDischargeEntry.Enabled = True
 cmdSave.Enabled = False
 cmdCancel.Enabled = False
 Else
 'txtBedCode.Text = ""
' txtBedNo.Text = ""
 txtEntry.Text = "Admit"
 txtDate.Text = ""
 txtPStatus.Text = ""

 cmdDischargeEntry.Enabled = True
 cmdSave.Enabled = False
 cmdCancel.Enabled = False
 End If
 End If
End Sub

Private Sub cmdDischargeEntry_Click()
 txtPId.SetFocus
 txtDate.Text = Date
 txtPId.Text = ""
 txtName.Text = ""
 'txtAge .Text = ""
 txtAddress.Text = ""
 txtPhNo.Text = ""
 txtPStatus.Text = ""
 txtBedCode.Text = ""
 txtBedNo.Text = ""
 txtEntry.Text = "Discharge"

 txtPStatus.Locked = False
 txtPId.Locked = False

 cmdSave.Enabled = True
 cmdDischargeEntry.Enabled = False
 cmdSave.Enabled = True
 cmdCancel.Enabled = True
End Sub

Private Sub Form_Load()

txtDate.Text = ""
txtPId.Text = ""
txtName.Text = ""
txtAge.Text = ""
txtAddress.Text = ""
txtPhNo.Text = ""
txtBedCode.Text = ""
txtBedNo.Text = ""
txtPStatus.Text = ""

'txtDate = Date
 'Width = 9000
 'Height = 8235
 txtEntry.Text = "Dischage"
End Sub





Private Sub txtBedNo_LostFocus()
 Reservation2.Show
 Reservation2.Data1.RecordSource = "SELECT* FROM Availability where BedCode='" & (txtBedCode.Text) & "' AND BedNo='" & (txtBedNo.Text) & "'"
 Reservation2.Data1.Refresh
End Sub

Private Sub txtPId_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub

Private Sub txtPId_LostFocus()
Dim query2 As String

query2 = "select EntryType from AInt where PatientID=" & Val(txtPId.Text) & ""
Adodc2.RecordSource = query2
Adodc2.CommandType = adCmdText
Adodc2.Refresh

If Adodc2.Recordset("EntryType") = "Discharge" Then
MsgBox "This patient is already discharged", vbInformation, "HMS"
Else
If txtPId.Text <> "" Then
query = "select * from TestTransaction where PatientID=" & Val(txtPId.Text) & ""
Adodc3.RecordSource = query
Adodc3.CommandType = adCmdText
Adodc3.Refresh
Do While Not Adodc3.Recordset.EOF

If Adodc3.Recordset("PatientID") = Val(txtPId.Text) Then
'Adodc3.Recordset(3) = "TRUE"
'Adodc3.Recordset.Update
    If Adodc3.Recordset("Balance") <> 0 Then
        MsgBox "Please pay all the bill first"
        TreatmentBill.Show
        GoTo p:
    End If
    
    query1 = "select * from AInt where PatientID=" & Val(txtPId.Text) & ""
    Adodc1.RecordSource = query1
    Adodc1.CommandType = adCmdText
    Adodc1.Refresh
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset("PatientID") = Val(txtPId.Text) Then
            txtName = Adodc1.Recordset(4)
            txtAge = Adodc1.Recordset(5)
            txtAddress = Adodc1.Recordset(6)
            txtPhNo = Adodc1.Recordset(7)
            txtBedCode = Adodc1.Recordset(1)
            txtBedNo = Adodc1.Recordset(2)
            'txtPStatus = Adodc1.Recordset(8)
        GoTo p:
        End If
    Adodc1.Recordset.MoveNext
    Loop
    GoTo p:
Else
MsgBox "Please Pay first Bill"
End If
Adodc3.Recordset.MoveNext
Loop
End If
End If
p:

'If Adodc3.Recordset.RecordCount = 0 Then
'
' MsgBox "Please pay all the treatment bill first"
 
' txtBedCode.Text = ""
' txtBedNo.Text = ""
' txtBedNo.Enabled = False
' txtEntry.Text = "Admit"
' txtDate.Text = ""
' txtPStatus.Text = ""

' cmdDischargeEntry.Enabled = True
' cmdSave.Enabled = False
' cmdCancel.Enabled = False
  'cmdEnd.TabIndex = 2
 'cmdEnd.SetFocus
 'TreatmentBill.Show
 
'Else

'Adodc2.RecordSource = "SELECT * From TransferIntimation WHERE TransferIntimation.PatientID='" & (txtPId.Text) & " '"
' Adodc2.Refresh

' Adodc2.Recordset.MoveLast
 
' Adodc1.RecordSource = "SELECT * FROM AInt WHERE AInt.PatientID = '" & (txtPId.Text) & " '"
' Adodc1.Refresh
' txtDate.Text = Date

 'If Adodc1.Recordset.RecordCount = 0 Then
 'MsgBox "There is No Patient with this ID"
'  txtBedNo.Enabled = False
  'cmdEnd.TabIndex = 2
 'cmdEnd.SetFocus
 'End If

 
'If Data1.Recordset.NoMatch = True Then
 'MsgBox "There is No Patient with this ID"
  'txtBedNo.Enabled = False
  'cmdEnd.TabIndex = 2
 'cmdEnd.SetFocus
 

'ElseIf txtEntry.Text = "Discharge" Then
' MsgBox "Patient is already discharged"
' cmdCancel.Enabled = False
' cmdSave.Enabled = False
' cmdDischargeEntry.Enabled = True
' cmdEnd.Enabled = True
' txtBedNo.Enabled = True
' cmdEnd.TabIndex = 2
' cmdEnd.SetFocus
 
 
'Else
' txtEntry.Text = "Discharge"
'End If

'End If
'End If
'
End Sub

