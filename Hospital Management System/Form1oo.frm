VERSION 5.00
Begin VB.Form Form1oo 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6660
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1oo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   6120
      Width           =   6855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   6855
      Begin VB.TextBox txtGender 
         BorderStyle     =   0  'None
         DataField       =   "Gender"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Index           =   0
         Left            =   5760
         TabIndex        =   24
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtEntry 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "EntryType"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Admit"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtBedNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "BedNo"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3960
         Width           =   495
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "AdmissionDate"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "PatientID"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Age"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   5760
         TabIndex        =   7
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtPhNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "PhoneNo"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Address"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   735
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Name"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox txtBedCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lblGender 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   4560
         TabIndex        =   23
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "********************************************************************************************************************"
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
         TabIndex        =   22
         Top             =   720
         Width           =   6615
      End
      Begin VB.Label lblAInt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Admission Intimation"
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
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   3615
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
         Left            =   480
         TabIndex        =   20
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Bed No"
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
         Left            =   2760
         TabIndex        =   19
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label lblBedCode 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Bed Code"
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
         Left            =   480
         TabIndex        =   18
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label lblEntry 
         BackColor       =   &H00C0E0FF&
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
         Left            =   5280
         TabIndex        =   17
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00C0E0FF&
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
         Left            =   5280
         TabIndex        =   16
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblPhNo 
         BackColor       =   &H00C0E0FF&
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
         Left            =   480
         TabIndex        =   15
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0E0FF&
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
         Left            =   480
         TabIndex        =   14
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
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
         Left            =   480
         TabIndex        =   13
         Top             =   2040
         Width           =   615
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
         Left            =   480
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
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
      TabIndex        =   1
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Form1oo.frx":000C
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   6615
   End
End
Attribute VB_Name = "Form1oo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Form_Click()
'On Error GoTo Err

Private Sub cmdOK_Click()
Unload Me
End Sub

'Form1oo.Hide
'Unload Me
'Exit Sub
'Err:
'MsgBox "Printer Error"
'Unload Me

'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo Err

Form1oo.Hide
Unload Me
Exit Sub
Err:
MsgBox "Printer Error"
Unload Me

End Sub

Private Sub Form_Load()
txtEntry.Text = AdmissionInt.txtEntry.Text
txtBedCode.Text = AdmissionInt.cboBedCode.Text
txtPId.Text = AdmissionInt.txtPId.Text
txtAddress.Text = AdmissionInt.txtAddress.Text
txtName.Text = AdmissionInt.txtName.Text
txtDate.Text = AdmissionInt.txtDate.Text
txtPhNo.Text = AdmissionInt.txtPhNo.Text
txtAge.Text = AdmissionInt.txtAge.Text
txtBedNo.Text = AdmissionInt.txtBedNo.Text
txtGender(0).Text = AdmissionInt.cboGender.Text
End Sub

