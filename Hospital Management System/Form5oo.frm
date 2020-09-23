VERSION 5.00
Begin VB.Form Form5oo 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6285
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form5oo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   5055
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Date"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtBillNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "BillNo"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PatientID"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "TotalCharge"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Balance"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   4200
         Width           =   855
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         ItemData        =   "Form5oo.frx":000C
         Left            =   2040
         List            =   "Form5oo.frx":000E
         TabIndex        =   3
         Top             =   2400
         Width           =   615
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         ItemData        =   "Form5oo.frx":0010
         Left            =   1320
         List            =   "Form5oo.frx":0012
         TabIndex        =   2
         Top             =   2400
         Width           =   615
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         ItemData        =   "Form5oo.frx":0014
         Left            =   240
         List            =   "Form5oo.frx":0016
         TabIndex        =   1
         Top             =   2400
         Width           =   975
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
         TabIndex        =   16
         Top             =   720
         Width           =   4335
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
         Left            =   2760
         TabIndex        =   15
         Top             =   960
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
         Left            =   240
         TabIndex        =   14
         Top             =   960
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
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblFBill 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Final Bill"
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
         TabIndex        =   12
         Top             =   240
         Width           =   2775
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
         TabIndex        =   11
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label lblReceived 
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
         TabIndex        =   10
         Top             =   4200
         Width           =   975
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
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Form5oo.frx":0018
      Height          =   615
      Left            =   360
      TabIndex        =   18
      Top             =   480
      Width           =   4575
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
      TabIndex        =   17
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form5oo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Form_Click()
'On Error GoTo Err

'Form5oo.Hide
'Unload Me
'Exit Sub
'Err:
'MsgBox "Printer Error"
'Unload Me
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Err

Form5oo.Hide
Unload Me
Exit Sub
Err:
MsgBox "Printer Error"
Unload Me
End Sub

