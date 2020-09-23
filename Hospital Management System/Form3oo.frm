VERSION 5.00
Begin VB.Form Form3oo 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6570
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form3oo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   5055
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Date"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Received"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtTreatNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "TreatNo"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtRefDoctor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "RefDoc"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtBalance 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Balance"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Charge"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtSpecification 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Cause"
         DataSource      =   "Data1"
         Height          =   615
         Left            =   1560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "PatientID"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label lblTreatNo 
         BackColor       =   &H00C0E0FF&
         Caption         =   "TreatNo"
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
         Left            =   3360
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblRefDoctor 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   19
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblBalance 
         BackColor       =   &H00C0E0FF&
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
         Left            =   240
         TabIndex        =   18
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label lblReceived 
         BackColor       =   &H00C0E0FF&
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
         Left            =   240
         TabIndex        =   17
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lblTest 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Treatment/Test"
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
         TabIndex        =   16
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblCharge 
         BackColor       =   &H00C0E0FF&
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
         Left            =   240
         TabIndex        =   15
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lblSpecification 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Specification"
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
         Top             =   2880
         Width           =   1095
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
         TabIndex        =   13
         Top             =   1080
         Width           =   495
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
         Left            =   2040
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblTBill 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Test Transaction Bill"
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
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "******************************************************************************************************************"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Form3oo.frx":000C
      Height          =   615
      Left            =   360
      TabIndex        =   22
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
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
      Left            =   360
      TabIndex        =   21
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form3oo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub Form_Click()
'On Error GoTo Err

'Form3oo.Hide
'Unload Me
'Exit Sub
'Err:
'MsgBox "Printer Error"
'Unload Me
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'On Error GoTo Err

Form3oo.Hide
Unload Me
TestTransaction.Command1.SetFocus
'Exit Sub
'Err:
'MsgBox "Printer Error"
'Unload Me
End Sub

Private Sub Form_Load()
txtDate.Text = TestTransaction.txtDate.Text
txtPId.Text = TestTransaction.txtPId.Text
txtSpecification.Text = TestTransaction.txtSpecification.Text
txtCharge.Text = TestTransaction.txtCharge.Text
txtReceived.Text = TestTransaction.txtReceived.Text
txtBalance.Text = TestTransaction.txtBalance.Text
txtRefDoctor.Text = TestTransaction.txtRefDoctor.Text
txtTreatNo.Text = TestTransaction.txtTreatNo.Text
txtTest.Text = TestTransaction.cboTest.Text

End Sub


