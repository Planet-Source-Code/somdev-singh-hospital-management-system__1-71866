VERSION 5.00
Begin VB.Form DischargeInt 
   BackColor       =   &H00808080&
   Caption         =   "Discharge Intimation"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   7125
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\practicals\vb files\project\HOSPITAL.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FinalBill"
      Top             =   7440
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\practicals\vb files\project\HOSPITAL.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TreatBill"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "End"
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
      TabIndex        =   23
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
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
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\practicals\vb files\project\HOSPITAL.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TransferIntimation"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
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
      Left            =   3840
      TabIndex        =   5
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
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
      Left            =   2520
      TabIndex        =   4
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Recod"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\practicals\vb files\project\HOSPITAL.mdb"
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
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "AInt"
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   6375
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   6615
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "ToBedNo"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "ToBedCode"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "EntryType"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Discharge"
         Top             =   5400
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PhoneNo"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Address"
         DataSource      =   "Data1"
         Height          =   735
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Age"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PatientID"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text1 
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
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "DischargeStatus"
         DataSource      =   "Data1"
         Height          =   765
         Left            =   600
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   5400
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Name"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label11 
         BackColor       =   &H00808080&
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
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
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
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
      Begin VB.Label Label10 
         BackColor       =   &H00808080&
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
         Left            =   3840
         TabIndex        =   19
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00808080&
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
      Begin VB.Label Label8 
         BackColor       =   &H00808080&
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
      Begin VB.Label Label7 
         BackColor       =   &H00808080&
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
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
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
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
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
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
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
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
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

Private Sub Command1_Click()
 DischargeInt.Hide
End Sub

Private Sub Command2_Click()



If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text10.Text = "" Or Text11.Text = "" Then
 MsgBox "Please fill all the fields"
Else
 TransferInt.Show
 TransferInt.Data1.RecordSource = "Select * from TransferIntimation Where PatientID='" & (DischargeInt.Text2.Text) & " '"
 TransferInt.Data1.Refresh
 TransferInt.List1.Visible = False
 TransferInt.Command1.Visible = False
 TransferInt.Text3.Visible = False
 TransferInt.Label13.Visible = False
 TransferInt.Command6.Enabled = False
 Command2.Enabled = True
 TransferInt.Data1.Recordset.AddNew
 TransferInt.Text4.Locked = False
 TransferInt.Command2.Enabled = True
 TransferInt.Command1.Enabled = False
 TransferInt.Text2.Text = Date
 TransferInt.Command6.Enabled = False
 TransferInt.Text1.Text = DischargeInt.Text2.Text
 TransferInt.Text5.Text = DischargeInt.Text10.Text
 TransferInt.Text6.Text = DischargeInt.Text11.Text
 TransferInt.Combo1.Text = "--"
 TransferInt.Text9.Text = "--"
 TransferInt.Text4.Enabled = True
 TransferInt.Text4.Locked = False
 TransferInt.Text6.Locked = True
 TransferInt.Text4.SetFocus
 TransferInt.Text5.Enabled = False
 TransferInt.Text6.Enabled = False
 TransferInt.Text7.Enabled = False
 TransferInt.Text8.Enabled = True
 TransferInt.Text8.Locked = False
 TransferInt.Text9.Enabled = False
 TransferInt.Combo1.Enabled = False
 
 Form2oo.Show

 Data1.Recordset.Edit
 Data1.Recordset.Update

 Command4.Enabled = True
 Command2.Enabled = False
 Command3.Enabled = False
End If

End Sub

Private Sub Command3_Click()
Dim temp
temp = MsgBox("you really don't want to take discharge?", vbYesNo)
If temp = vbYes Then
If Text10.Text <> "" And Text11.Text <> "" Then
Reservation.Show
Reservation.Data1.RecordSource = "select * from Availability where BedCode='" & Text10.Text & "'"
'Reservation.Data1.Refresh

' Text10.Text = ""
 'Text11.Text = ""
 Text8.Text = "Admit"
 Text1.Text = ""
 Text7.Text = ""

 Command4.Enabled = True
 Command2.Enabled = False
 Command3.Enabled = False
 Else
 'Text10.Text = ""
' Text11.Text = ""
 Text8.Text = "Admit"
 Text1.Text = ""
 Text7.Text = ""

 Command4.Enabled = True
 Command2.Enabled = False
 Command3.Enabled = False
 End If
 End If
End Sub

Private Sub Command4_Click()
 Text2.SetFocus
 Text1.Text = Date
 Text2.Text = ""
 Text3.Text = ""
 Text4.Text = ""
 Text5.Text = ""
 Text6.Text = ""
 Text7.Text = ""
 Text10.Text = ""
 Text11.Text = ""
 Text8.Text = "Discharge"

 Text7.Locked = False
 Text2.Locked = False

 Command2.Enabled = True
 Command4.Enabled = False
 Command2.Enabled = True
 Command3.Enabled = True
End Sub

Private Sub Form_Load()
 Width = 7305
 Height = 8235
 Text8.Text = "Dischage"
End Sub

Private Sub Text11_LostFocus()
 Reservation2.Show
 Reservation2.Data1.RecordSource = "SELECT* FROM Availability where BedCode='" & (Text10.Text) & "' AND BedNo='" & (Text11.Text) & "'"
 Reservation2.Data1.Refresh
End Sub

Private Sub Text2_LostFocus()
If Text2.Text <> "" Then
 Data4.RecordSource = "select * from finalBill where finalBill.PatientID='" & (Text2.Text) & "'"
 Data4.Refresh

If Data4.Recordset.RecordCount = 0 Then
 MsgBox "Please pay all the treatment bill first"

 Text10.Text = ""
 Text11.Text = ""
 Text8.Text = "Admit"
 Text1.Text = ""
 Text7.Text = ""

 Command4.Enabled = True
 Command2.Enabled = False
 Command3.Enabled = False

Else
 Data1.RecordSource = "SELECT * FROM AInt WHERE AInt.PatientID = '" & (Text2.Text) & " '"
 Data1.Refresh

 Data2.RecordSource = "SELECT * From TransferIntimation WHERE TransferIntimation.PatientID='" & (Text2.Text) & " '"
 Data2.Refresh

 Data2.Recordset.MoveLast
 Text1.Text = Date

If Data1.Recordset.NoMatch = True Then
 MsgBox "There is No Patient with this ID"

ElseIf Text8.Text = "Discharge" Then
 MsgBox "Patient is already discharged"
 Command3.Enabled = False
 Command2.Enabled = False
 Command4.Enabled = True
 Command1.Enabled = True

Else
 Text8.Text = "Discharge"
End If

End If
End If

End Sub

