VERSION 5.00
Begin VB.Form Form4oo 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6525
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5760
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form4oo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   5175
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   5055
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "date"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PatientID"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ListBox lstTreatment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1980
         ItemData        =   "Form4oo.frx":000C
         Left            =   240
         List            =   "Form4oo.frx":0013
         TabIndex        =   8
         Top             =   2160
         Width           =   1695
      End
      Begin VB.ListBox lstCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1980
         ItemData        =   "Form4oo.frx":0025
         Left            =   2160
         List            =   "Form4oo.frx":002C
         TabIndex        =   7
         Top             =   2160
         Width           =   735
      End
      Begin VB.ListBox lstReceived 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1980
         ItemData        =   "Form4oo.frx":003B
         Left            =   3120
         List            =   "Form4oo.frx":0042
         TabIndex        =   6
         Top             =   2160
         Width           =   735
      End
      Begin VB.ListBox lstBalance 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1980
         ItemData        =   "Form4oo.frx":0053
         Left            =   4080
         List            =   "Form4oo.frx":005A
         TabIndex        =   5
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Charge"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Received"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtBalance 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "balance"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtPaid 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label lblFTBill 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   19
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   18
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label lblPId 
         BackColor       =   &H00C0E0FF&
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
         Left            =   2640
         TabIndex        =   17
         Top             =   1320
         Width           =   855
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
         TabIndex        =   16
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblTreatment 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   15
         Top             =   1680
         Width           =   855
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
         Left            =   2160
         TabIndex        =   14
         Top             =   1680
         Width           =   615
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
         Left            =   3120
         TabIndex        =   13
         Top             =   1680
         Width           =   855
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
         Left            =   4080
         TabIndex        =   12
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblPaid 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   240
         TabIndex        =   11
         Top             =   4440
         Width           =   495
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Form4oo.frx":006A
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   600
      Width           =   5055
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
      TabIndex        =   20
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "Form4oo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub Form_Click()
'On Error GoTo Err

'Form4oo.Hide
'Unload Me
'Exit Sub
'Err:
'MsgBox "Printer Error"
'Unload Me
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Err

Form4oo.Hide
Unload Me
Exit Sub
Err:
MsgBox "Printer Error"
Unload Me
End Sub

