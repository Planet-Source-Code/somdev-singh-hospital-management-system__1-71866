VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6795
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Cancel"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Start"
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
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4695
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   5175
      Begin VB.Image Image1 
         Height          =   3750
         Left            =   600
         Picture         =   "frmSplash.frx":5ADA
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "************************************************************************************"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   135
         Left            =   240
         TabIndex        =   5
         Top             =   4440
         Width           =   4935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "**********************************************************************************"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   135
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Patient Monitoring System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'Private Sub Form_KeyPress(KeyAscii As Integer)
'frmSplash.Hide
'Record.Show
'End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'frmSplash.Hide
'Record.Show
'End Sub

'Private Sub Image1_Click()
'frmSplash.Hide
'Record.Show
'End Sub

'Private Sub Image2_Click()
'End
'End Sub

'Private Sub Image3_Click()
'frmSplash.Hide
'Record.Show
'End Sub
Private Sub Command1_Click()

If Text1.Text <> "private" Then
MsgBox "Your password is wrong"
Text1.Text = ""
Text1.SetFocus

Else
Text1.Text = ""
Unload Me
Record.Show
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_LostFocus()
If Text1.Text <> "" Then Command1.Enabled = True
'Command1.SetFocus
End Sub
