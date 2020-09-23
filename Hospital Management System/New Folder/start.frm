VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSplash1 
   ClientHeight    =   5265
   ClientLeft      =   270
   ClientTop       =   1425
   ClientWidth     =   6795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FF0000&
   Icon            =   "start.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5265
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtlog 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtpass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "login"
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   4800
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6360
      Top             =   4320
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Default         =   -1  'True
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
      Left            =   3840
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
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
      Left            =   4680
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdRet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Retry"
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
      Left            =   5520
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblPer 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOBLE HOSPITAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblLoad 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   0
      Picture         =   "start.frx":5ADA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7065
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
Static count As Integer
If count > 1 Then
MsgBox "sorry! Try Again", vbExclamation
End
End If
Adodc1.Recordset.MoveFirst
Do Until Adodc1.Recordset.EOF
If (Adodc1.Recordset.Fields(1) = txtpass.Text) And (Adodc1.Recordset.Fields(0) = txtlog.Text) Then
ProgressBar1.Value = 0
Timer1.Enabled = True
ProgressBar1.Visible = True
lblLoad.Visible = True
Exit Sub
Else
Adodc1.Recordset.MoveNext
End If
Loop
MsgBox "Please enter Valid Login name & Password", vbInformation + vbOKOnly
count = count + 1
End Sub

Private Sub cmdret_Click()
txtpass.Text = ""
txtlog.Text = ""
txtlog.SetFocus
End Sub

Private Sub Form_Load()
Timer1.Enabled = 0
ProgressBar1.Value = 0
ProgressBar1.Visible = False
lblLoad.Visible = False
txtlog = ""
txtpass.Text = ""
End Sub

Private Sub Timer1_Timer()
lblPer.Caption = ProgressBar1.Value & "%"
If ProgressBar1.Value Mod 2 = 0 Then
lblLoad.Visible = True
Else
lblLoad.Visible = False
End If
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value >= 100 Then
Unload Me
Record.Show
Timer1.Enabled = False
End If
End Sub

Private Sub txtlog_Change()
User = txtlog.Text
End Sub

Private Sub txtpass_Change()
txtpass.PasswordChar = "*"
txtpass.ForeColor = vbRed
txtpass.FontSize = 14
End Sub
