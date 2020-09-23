VERSION 5.00
Begin VB.Form Form2oo 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7530
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form2oo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   6375
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   6615
      Begin VB.TextBox txtBedNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "ToBedNo"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtBedCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "ToBedCode"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "EntryType"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Discharge"
         Top             =   5400
         Width           =   1095
      End
      Begin VB.TextBox txtPhNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "PhoneNo"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Address"
         DataSource      =   "Data1"
         Height          =   735
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Age"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "PatientID"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "DischargeDate"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtPStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "DischargeStatus"
         DataSource      =   "Data1"
         Height          =   765
         Left            =   600
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   5400
         Width           =   3135
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Name"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   21
         Top             =   840
         Width           =   6375
      End
      Begin VB.Label lblDInt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   20
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lblBed 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   19
         Top             =   4560
         Width           =   975
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
         Left            =   3840
         TabIndex        =   18
         Top             =   5040
         Width           =   1335
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
         Left            =   120
         TabIndex        =   17
         Top             =   4080
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
         Left            =   240
         TabIndex        =   16
         Top             =   3120
         Width           =   735
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
         Left            =   240
         TabIndex        =   15
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblPStatus 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   14
         Top             =   5040
         Width           =   3255
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
         Left            =   240
         TabIndex        =   13
         Top             =   2160
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
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Form2oo.frx":000C
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   360
      Width           =   6615
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
      TabIndex        =   22
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "Form2oo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Form_Click()
'On Error GoTo Err

'Form2oo.Hide
'Unload Me
'Exit Sub
'Err:
'MsgBox "Printer Error"
'Unload Me

'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo Err

Form2oo.Hide
Unload Me
Exit Sub
Err:
MsgBox "Printer Error"
Unload Me

End Sub

Private Sub Form_Load()
txtDate.Text = DischargeInt.txtDate.Text
txtPId.Text = DischargeInt.txtPId.Text
txtName.Text = DischargeInt.txtName.Text
txtAge.Text = DischargeInt.txtAge.Text
txtAddress.Text = DischargeInt.txtAddress.Text
txtPhNo.Text = DischargeInt.txtPhNo.Text
txtPStatus.Text = DischargeInt.txtPStatus.Text
txtEntry.Text = DischargeInt.txtEntry.Text
txtBedCode.Text = DischargeInt.txtBedCode.Text
txtBedNo.Text = DischargeInt.txtBedNo.Text
End Sub


