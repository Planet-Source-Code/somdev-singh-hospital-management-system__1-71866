VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Hospital(Discharge Bill)"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8295
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   8295
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   5055
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         ItemData        =   "Form5.frx":0000
         Left            =   240
         List            =   "Form5.frx":0002
         TabIndex        =   15
         Top             =   2400
         Width           =   975
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         ItemData        =   "Form5.frx":0004
         Left            =   1320
         List            =   "Form5.frx":0006
         TabIndex        =   14
         Top             =   2400
         Width           =   615
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         ItemData        =   "Form5.frx":0008
         Left            =   2040
         List            =   "Form5.frx":000A
         TabIndex        =   13
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Balance"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "TotalCharge"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox txtPId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PatientID"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtBillNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "BillNo"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Date"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   960
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
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
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
         TabIndex        =   12
         Top             =   4200
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
         TabIndex        =   11
         Top             =   4200
         Width           =   855
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
         TabIndex        =   10
         Top             =   240
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
         TabIndex        =   9
         Top             =   1680
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
         TabIndex        =   8
         Top             =   960
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
         Left            =   2760
         TabIndex        =   7
         Top             =   960
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
         TabIndex        =   6
         Top             =   720
         Width           =   4335
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Form5.frx":000C
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
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Me.Hide
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
Me.Hide
End Sub

