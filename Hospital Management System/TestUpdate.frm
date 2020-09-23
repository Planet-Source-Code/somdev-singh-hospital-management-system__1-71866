VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form TreatmentBill 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TreatmentBill"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   5520
      Top             =   7920
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
      RecordSource    =   "select * from final_bill"
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
   Begin VB.CommandButton cmdSaveAll 
      Caption         =   "Save"
      Height          =   375
      Left            =   6720
      TabIndex        =   39
      Top             =   6960
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   5775
      Left            =   6000
      TabIndex        =   26
      Top             =   960
      Width           =   4815
      Begin VB.TextBox txtMedicine 
         Height          =   285
         Left            =   2880
         TabIndex        =   32
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtOperation 
         Height          =   285
         Left            =   2880
         TabIndex        =   31
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtdoctorvisit 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2880
         TabIndex        =   30
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtNursingcharge 
         Height          =   285
         Left            =   2880
         TabIndex        =   29
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtBedcharge 
         Height          =   285
         Left            =   2880
         TabIndex        =   28
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "TotalCharge"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2880
         TabIndex        =   27
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Medicine"
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
         Left            =   1320
         TabIndex        =   38
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Operation / Anasthesia"
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
         TabIndex        =   37
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nursing Charge"
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
         Left            =   840
         TabIndex        =   36
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Doctor visit"
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
         Left            =   1080
         TabIndex        =   35
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bed Charge"
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
         Left            =   1080
         TabIndex        =   34
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total"
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
         Left            =   1440
         TabIndex        =   33
         Top             =   4440
         Width           =   855
      End
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6960
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   5055
      Begin MSDataListLib.DataList lstBalance 
         Height          =   2010
         Left            =   4080
         TabIndex        =   25
         Top             =   2760
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   3545
         _Version        =   393216
      End
      Begin MSDataListLib.DataList lstReceived 
         Height          =   2010
         Left            =   3120
         TabIndex        =   24
         Top             =   2760
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   3545
         _Version        =   393216
      End
      Begin MSDataListLib.DataList lstCharge 
         Height          =   2010
         Left            =   2160
         TabIndex        =   23
         Top             =   2760
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   3545
         _Version        =   393216
      End
      Begin MSDataListLib.DataList lstTreatment 
         Height          =   2010
         Left            =   240
         TabIndex        =   22
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   3545
         _Version        =   393216
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Date"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PatientID"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Charge"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Received"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtBalance 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "balance"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   4080
         TabIndex        =   5
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtPaid 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Final Treatment Bill"
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
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "*********************************************************************************"
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
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label lblPId 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PatientID"
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
         Top             =   1800
         Width           =   855
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
         TabIndex        =   14
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblTreatment 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Treatment"
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
         Top             =   2280
         Width           =   855
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
         Left            =   2160
         TabIndex        =   12
         Top             =   2280
         Width           =   615
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
         Left            =   3120
         TabIndex        =   11
         Top             =   2280
         Width           =   855
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
         Left            =   4080
         TabIndex        =   10
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblPaid 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Paid"
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
         Left            =   360
         TabIndex        =   9
         Top             =   5160
         Width           =   855
      End
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5400
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
      Left            =   5400
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
      RecordSource    =   "select * from TestTransaction"
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   600
      TabIndex        =   20
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "NOBLE HOSPITAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   21
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "TreatmentBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim id, tempo
Dim i
Dim totalcharge
Dim totalr
Dim totalb
Dim query, query1, query2, query3 As String
Dim ans

Public Sub saveprocedure()
Dim bal As Integer

bal = Val(txtTotal) - Val(txtBalance)

Dim query1 As String
query1 = "select * from TestTransaction where PatientID=" & Val(txtPId.Text) & " and Balance > 0"
Adodc2.RecordSource = query1
Adodc2.Refresh
If Adodc2.Recordset.EOF = True And Adodc2.Recordset.BOF = True Then
Adodc2.Recordset.AddNew
'Adodc2.Recordset.MoveLast
Else
Adodc2.Recordset("Received") = Adodc2.Recordset("Charge")
Adodc2.Recordset("balance") = "0"
Adodc2.Recordset.Update
End If

Dim query As String
query = "select * from final_bill"
Adodc.RecordSource = query
Adodc.Refresh

Adodc.Recordset.AddNew
'Adodc.Recordset.Fields(0) = txtBillNo
'Adodc.Recordset.Fields(1) = txtCharge
'Adodc.Recordset.Fields(2) = txtReceived
'Adodc.Recordset.Fields!TreatmentBalance = txtBalance
Adodc.Recordset.Fields!PatientID = txtPId
Adodc.Recordset.Fields!Date = txtDate
Adodc.Recordset.Fields!Bedcharge = txtBedcharge
Adodc.Recordset.Fields!nursingcharge = txtNursingcharge
Adodc.Recordset.Fields!Doctorvisit = txtdoctorvisit
Adodc.Recordset.Fields!Operation = txtOperation
Adodc.Recordset.Fields!MedicineCharge = txtMedicine
Adodc.Recordset.Update
MsgBox "Save sucessfully"
End Sub

Private Sub cmdNew_Click()

 Adodc2.Recordset.AddNew
 txtPId.Locked = False
 txtDate.Text = Date
 txtPId.SetFocus
 cmdNew.Enabled = False
 cmdSave.Enabled = True
 cmdCancel.Enabled = True
 cmdEnd.Enabled = False
End Sub

Private Sub cmdSave_Click()
 
On Error GoTo errdiscription
Call saveprocedure

' ans = MsgBox("Are you sure, you are paying all the bill?", vbYesNo)
'If ans = vbNo Then
' Adodc2.Recordset.CancelUpdate
' cmdSave.Enabled = False
' cmdNew.Enabled = True
' cmdCancel.Enabled = False
' cmdEnd.Enabled = True
'Else
'
'Form4oo.Show
'Form4oo.txtDate.Text = txtDate.Text
'Form4oo.txtPId.Text = txtPId.Text
'Form4oo.txtCharge.Text = txtCharge.Text
'Form4oo.txtReceived.Text = txtReceived.Text
'Form4oo.txtBalance.Text = txtBalance.Text
'Form4oo.txtPaid.Text = txtPaid.Text

'For i = 0 To lstTreatment.ListCount
'Form4oo.lstTreatment.AddItem (lstTreatment.List(i))
'Next

'For i = 0 To lstCharge.ListCount
'Form4oo.lstCharge.AddItem (lstCharge.List(i))
'Next

'For i = 0 To lstReceived.ListCount
'Form4oo.lstReceived.AddItem (lstReceived.List(i))
'Next

'For i = 0 To lstBalance.ListCount
'Form4oo.lstBalance.AddItem (lstBalance.List(i))

'
'Next

' Adodc2.Recordset.Update
 
 
 cmdNew.Enabled = True
 cmdSave.Enabled = False
 txtPId.Locked = True
cmdCancel.Enabled = False
cmdEnd.Enabled = True

 
'Exit Sub

errdiscription:
' MsgBox "This patient has paid all treatment bill & got reciept"
' Adodc2.Recordset.CancelUpdate
' cmdNew.Enabled = True
' cmdSave.Enabled = False
' cmdEnd.Enabled = True
 
'Form4oo.Hide

 '//changes are not checked
 'Form4oo.Hide
'End If

End Sub

Private Sub cmdcancel_Click()

 Adodc2.Recordset.CancelUpdate
 cmdSave.Enabled = False
 cmdNew.Enabled = True
 cmdCancel.Enabled = False
 cmdEnd.Enabled = False
End Sub

Private Sub cmdEnd_Click()
Unload Me
'MsgBox " Take Intimation of your discharge "
'DischargeInt.Show

End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmdSaveAll_Click()
Call saveprocedure
End Sub

Private Sub Form_Load()
txtDate = Date
'Height = 8055
'Width = 6330
Frame2.Visible = False
End Sub





Private Sub txtBalance_Change()
If txtBalance <> "" Then
Frame2.Visible = True
Else
Frame2.Visible = False
End If
End Sub

Private Sub txtBedcharge_Change()
txtTotal = Val(txtBedcharge) + Val(txtNursingcharge) + Val(txtdoctorvisit) + Val(txtOperation) + Val(txtMedicine) + Val(txtBalance)
End Sub

Private Sub txtdoctorvisit_Change()
txtTotal = Val(txtBedcharge) + Val(txtNursingcharge) + Val(txtdoctorvisit) + Val(txtOperation) + Val(txtMedicine) + Val(txtBalance)
End Sub

Private Sub txtMedicine_Change()
txtTotal = Val(txtBedcharge) + Val(txtNursingcharge) + Val(txtdoctorvisit) + Val(txtOperation) + Val(txtMedicine) + Val(txtBalance)
End Sub

Private Sub txtNursingcharge_Change()
txtTotal = Val(txtBedcharge) + Val(txtNursingcharge) + Val(txtdoctorvisit) + Val(txtOperation) + Val(txtMedicine) + Val(txtBalance)
End Sub

Private Sub txtOperation_Change()
txtTotal = Val(txtBedcharge) + Val(txtNursingcharge) + Val(txtdoctorvisit) + Val(txtOperation) + Val(txtMedicine) + Val(txtBalance)
End Sub

Private Sub txtPId_KeyPress(KeyAscii As Integer)


If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
KeyAscii = 0

End If
End Sub

Private Sub txtPId_LostFocus()
If txtPId.Text <> "" Then
query = "select * from TestTransaction where PatientID=" & txtPId.Text & ""
Adodc1.RecordSource = query
Adodc1.CommandType = adCmdText
Adodc1.Refresh

Do While Not Adodc1.Recordset.EOF

If Adodc1.Recordset(0) = Val(txtPId.Text) Then

Set lstTreatment.RowSource = Adodc1
lstTreatment.ListField = "Treatment"
lstTreatment.Refresh

Set lstCharge.RowSource = Adodc1
lstCharge.ListField = "Charge"
lstCharge.Refresh

Set lstReceived.RowSource = Adodc1
lstReceived.ListField = "Received"
lstReceived.Refresh

Set lstBalance.RowSource = Adodc1
lstBalance.ListField = "Balance"
lstBalance.Refresh

GoTo p:
End If

Adodc1.Recordset.MoveNext
Loop
End If
p:

query = "select * from TestTransaction where PatientID=" & Val(txtPId.Text) & ""
Adodc1.RecordSource = query
Adodc1.CommandType = adCmdText
Adodc1.Refresh

Dim mtot As Double
mtot = 0
With Adodc1.Recordset
Do Until .EOF
If Not IsNull(.Fields!charge) Then mtot = mtot + .Fields!charge
.MoveNext
Loop
End With
txtCharge = mtot

query = "select * from TestTransaction where PatientID=" & Val(txtPId.Text) & ""
Adodc1.RecordSource = query
Adodc1.CommandType = adCmdText
Adodc1.Refresh

Dim mtot1 As Double
mtot1 = 0
With Adodc1.Recordset
Do Until .EOF
If Not IsNull(.Fields!received) Then mtot1 = mtot1 + .Fields!received
.MoveNext
Loop
txtReceived = mtot1
End With

query = "select * from TestTransaction where PatientID=" & Val(txtPId.Text) & ""
Adodc2.RecordSource = query
Adodc2.CommandType = adCmdText
Adodc2.Refresh

Dim mtot2 As Double
mtot2 = 0
With Adodc2.Recordset
Do Until .EOF
If Not IsNull(.Fields!balance) Then mtot2 = mtot2 + .Fields!balance
.MoveNext
Loop
End With
txtBalance = mtot2

txtPaid = txtReceived
'If txtPId.Text <> "" Then

' totalcharge = 0
' totalr = 0
' totalb = 0
' lstTreatment.Clear
' lstCharge.Clear
' lstReceived.Clear
' lstBalance.Clear

' Adodc1.RecordSource = "select * from TestTransaction where PatientID='" & txtPId.Text & "'"
' Adodc1.Refresh
 
'If Adodc1.Recordset.RecordCount <> 0 Then

'Dim temp

 'Adodc1.Recordset.MoveFirst
 'temp = txtPId.Text
 'temp = "PatientId='" & temp & "'"
'
'With Adodc1.Recordset
 '    .FindFirst temp
'    End With
     
'If Adodc1.Recordset.NoMatch = False Then
' lstTreatment.AddItem Adodc1.Recordset.Fields(2).Value
' lstCharge.AddItem Adodc1.Recordset.Fields(4).Value
' lstReceived.AddItem Adodc1.Recordset.Fields(5).Value
' lstBalance.AddItem Adodc1.Recordset.Fields(6).Value
 
'For i = 0 To Adodc1.Recordset.RecordCount + 10
' temp = txtPId.Text
' temp = "PatientId='" & temp & "'"
 
'With Adodc1.Recordset
      '.FindNext temp
      'End With
  
'If Adodc1.Recordset.NoMatch = False Then
' lstTreatment.AddItem Adodc1.Recordset.Fields(2).Value
' lstCharge.AddItem Adodc1.Recordset.Fields(4).Value
' lstReceived.AddItem Adodc1.Recordset.Fields(5).Value
' lstBalance.AddItem Adodc1.Recordset.Fields(6).Value
' End If
 
' Next
 
' End If
 
'For i = 0 To lstCharge.ListCount
' totalcharge = totalcharge + (Val(lstCharge.List(i)))

'Next
' txtCharge.Text = totalcharge

'For i = 0 To lstReceived.ListCount
' totalr = totalr + (Val(lstReceived.List(i)))

'Next
' txtReceived.Text = totalr

'For i = 0 To lstBalance.ListCount
' totalb = totalb + (Val(lstBalance.List(i)))

'Next
' txtBalance.Text = totalb
'txtPaid.Text = txtBalance.Text

'ElseIf Adodc1.Recordset.RecordCount = 0 Then
' MsgBox "Patient with this ID is not found. Please check your entry."
' txtPId.Text = ""
' Adodc2.Recordset.CancelUpdate
' cmdNew.Enabled = True
' cmdSave.Enabled = False
' cmdCancel.Enabled = False
'End If
'End If
End Sub

Private Sub txtPaid_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If

End Sub


Private Sub txtTotal_Change()
cmdSave.Visible = True
End Sub
