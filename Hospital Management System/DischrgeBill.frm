VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form DischargeBill 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discharge Bill"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7935
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   6720
      Top             =   6360
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
      Left            =   5400
      Top             =   6360
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
      Left            =   6720
      Top             =   5760
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
      Top             =   5760
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
      RecordSource    =   "FinalBill"
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4575
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         ItemData        =   "DischrgeBill.frx":0000
         Left            =   240
         List            =   "DischrgeBill.frx":0002
         TabIndex        =   18
         Top             =   2760
         Width           =   975
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         ItemData        =   "DischrgeBill.frx":0004
         Left            =   1320
         List            =   "DischrgeBill.frx":0006
         TabIndex        =   17
         Top             =   2760
         Width           =   735
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         ItemData        =   "DischrgeBill.frx":0008
         Left            =   2160
         List            =   "DischrgeBill.frx":000A
         TabIndex        =   16
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtRecieved 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Balance"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "TotalCharge"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PatientID"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtBillNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "BillNo"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Date"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblTransaction 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Transaction"
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
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblRecieved 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Recieved"
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
         Left            =   2040
         TabIndex        =   15
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H00C0E0FF&
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
         Left            =   240
         TabIndex        =   14
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label lblDischargeBill 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Final Discharge Bill"
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
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblPId 
         BackColor       =   &H00C0E0FF&
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
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   11
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblBillNo 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Bill No"
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
         Left            =   2400
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "**************************************************************************************************************"
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
         TabIndex        =   9
         Top             =   720
         Width           =   4335
      End
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "..\Hospital Management System\Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "AInt"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "..\Hospital Management System\Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TestTransaction"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "..\Hospital Management System\Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TransferIntimation"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0FF&
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
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0FF&
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
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "..\Hospital Management System\Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FinalBill"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0FF&
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
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   3825
      Left            =   4800
      Picture         =   "DischrgeBill.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3120
   End
End
Attribute VB_Name = "DischargeBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim totalcharge, received, balance, charge

Private Sub cmdAdd_Click()
 Adodc1.Recordset.AddNew
 txtBillNo.Text = Adodc1.Recordset.RecordCount
 txtDate.Text = Date
 txtPId.SetFocus
 cmdAdd.Enabled = False
 cmdCancel.Enabled = True
 cmdSave.Enabled = True
End Sub

Private Sub cmdCancel_Click()
txtTotal.Text = ""
 Adodc1.Recordset.CancelUpdate
 cmdAdd.Enabled = True
 cmdCancel.Enabled = False
 cmdSave.Enabled = False
End Sub

Private Sub cmdSave_Click()
If Val(txtBillNo1.Text) < Val(txtTotal.Text) Then
MsgBox "Please pay all bill"
txtBillNo1.Text = ""
txtBillNo1.SetFocus
ElseIf Val(txtBillNo1.Text) > Val(txtTotal.Text) Then
MsgBox "Please pay exact bill"
txtBillNo1.Text = ""
txtBillNo1.SetFocus
Else

Form5oo.txtBillNo.Text = txtBillNo.Text
Form5oo.txtDate.Text = txtDate.Text
Form5oo.txtPId.Text = txtPId.Text
Form5oo.txtTotal.Text = txtTotal.Text
Form5oo.txtBillNo1.Text = txtBillNo1.Text

For i = 0 To List1.ListCount
Form5oo.List1.AddItem (List1.List(i))
Next

For i = 0 To List2.ListCount
Form5oo.List2.AddItem (List2.List(i))
Next

For i = 0 To List3.ListCount
Form5oo.List3.AddItem (List3.List(i))
Next

Form5oo.Show

 Adodc1.Recordset.Update
 cmdAdd.Enabled = True
 cmdCancel.Enabled = False
 cmdSave.Enabled = False
 End If
 
End Sub

Private Sub cmdEnd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Width = 8025
hight = 6000
End Sub



Private Sub txtBillNo1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If
End Sub

Private Sub txtPId_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
 KeyAscii = 0
End If
End Sub

Private Sub txtPId_LostFocus()
If txtDate.Text <> "" Then
 Adodc3.RecordSource = "select * from TreatBill where TreatBill.PatientID='" & (txtPId.Text) & "'"
 Adodc3.Refresh
 
 Adodc4.RecordSource = "select*from AInt where AInt.PatientID='" & (txtPId.Text) & "'"
 Adodc4.Refresh

If Adodc3.Recordset.RecordCount = 0 Then
 MsgBox "Please pay all the treatment bill first"
 
 txtTotal.Text = ""
 Adodc1.Recordset.CancelUpdate
 cmdAdd.Enabled = True
 cmdCancel.Enabled = False
 cmdSave.Enabled = False
 
 
 ElseIf Adodc4.Recordset.Fields(4).Value <> "Discharge" Then
 MsgBox "Please release the bed"

 
 txtTotal.Text = ""
 Adodc1.Recordset.CancelUpdate
 cmdAdd.Enabled = True
 cmdCancel.Enabled = False
 cmdSave.Enabled = False
 DischargeInt.Show
 
Else


Dim i

Dim id As String

 List1.Clear
 List2.Clear
 List3.Clear

 Adodc2.Recordset.MoveFirst
 id = txtPId.Text
 id = "PatientID = '" & id & "'"
 
With Adodc2.Recordset
     .FindFirst id
End With

If Adodc2.Recordset.NoMatch = False Then
   List1.AddItem Adodc2.Recordset.Fields(0).Value
   List2.AddItem Adodc2.Recordset.Fields(5).Value
   List3.AddItem Adodc2.Recordset.Fields(7).Value
 
For i = 0 To Adodc2.Recordset.RecordCount + 1
    id = txtPId.Text
    id = "PatientID= '" & id & "'"

With Adodc2.Recordset
     .FindNext id
End With

If Adodc2.Recordset.NoMatch = False Then
   List1.AddItem Adodc2.Recordset.Fields(0).Value
   List2.AddItem Adodc2.Recordset.Fields(5).Value
   List3.AddItem Adodc2.Recordset.Fields(7).Value
End If
      
Next
End If

For i = 0 To List3.ListCount
    charge = charge + (Val(List3.List(i)))
Next

 Adodc4.Recordset.MoveFirst
 id = txtPId.Text
 id = "PatientID= '" & id & "'"

With Adodc4.Recordset
.FindFirst id
End With

For i = 0 To List3.ListCount
 totalcharge = totalcharge + (Val(List3.List(i)))
Next

txtTotal.Text = totalcharge
'txtBillNo1.Text = Val(txtTotal.Text) - Val(Text9.Text)
Exit Sub
End If
End If
End Sub

'Private Sub txtTotal_GotFocus()'

'For i = 0 To List5.ListCount
 '   totalcharge = totalcharge + (Val(List5.List(i)))

'Next
' totalcharge = totalcharge + charge
' txtTotal.Text = totalcharge

'For i = 0 To List6.ListCount
 'received = received + (Val(List6.List(i)))
'Next
 'txtBillNo0.Text = received
 'txtBillNo1.Text = Val(txtTotal.Text) - (Val(Text9.Text) + Val(txtBillNo0.Text))
'Exit Sub

'End Sub


