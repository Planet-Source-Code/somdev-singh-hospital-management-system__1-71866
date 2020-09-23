VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSecurity 
   Caption         =   "RECEIPT"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1200
      Top             =   6240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\New Folder (2)\Hospital Management System1\Hms.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\New Folder (2)\Hospital Management System1\Hms.mdb;Persist Security Info=False"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Add/Delete User"
      TabPicture(0)   =   "frmReceipt.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Change Passwurd"
      TabPicture(1)   =   "frmReceipt.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "txtCon"
      Tab(1).Control(2)=   "txtnew"
      Tab(1).Control(3)=   "txtOld"
      Tab(1).Control(4)=   "cboUser"
      Tab(1).Control(5)=   "Label4"
      Tab(1).Control(6)=   "Label3"
      Tab(1).Control(7)=   "Label2"
      Tab(1).Control(8)=   "Label1"
      Tab(1).ControlCount=   9
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "User's Details"
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   480
         TabIndex        =   9
         Top             =   1560
         Width           =   4935
         Begin VB.TextBox txtLogn 
            Appearance      =   0  'Flat
            DataField       =   "log in"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   2160
            TabIndex        =   24
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            DataField       =   "password"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   2160
            TabIndex        =   16
            Top             =   1560
            Width           =   1815
         End
         Begin VB.TextBox txtCpass 
            Appearance      =   0  'Flat
            DataField       =   "confpass"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   2160
            TabIndex        =   15
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Login Name"
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
            Left            =   360
            TabIndex        =   19
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "New Password"
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
            Left            =   360
            TabIndex        =   18
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Confirm Password"
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
            Left            =   360
            TabIndex        =   17
            Top             =   2160
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   3000
         TabIndex        =   25
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton cmdnext 
            Height          =   495
            Left            =   1440
            Picture         =   "frmReceipt.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdprev 
            Height          =   495
            Left            =   480
            Picture         =   "frmReceipt.frx":0186
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   -74040
         TabIndex        =   20
         Top             =   3840
         Width           =   4095
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   615
            Left            =   1560
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdChange 
            Caption         =   "Change"
            Height          =   615
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdcExut 
            Caption         =   "Exit"
            Height          =   615
            Left            =   2880
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   360
         TabIndex        =   10
         Top             =   4320
         Width           =   5415
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   615
            Left            =   4200
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdDlete 
            Caption         =   "Delete"
            Height          =   615
            Left            =   2880
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "Add"
            Height          =   615
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdsave 
            Caption         =   "Save"
            Height          =   615
            Left            =   1560
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtCon 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72600
         TabIndex        =   8
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtnew 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72600
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtOld 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -72600
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox cboUser 
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "frmReceipt.frx":02D4
         Left            =   -72600
         List            =   "frmReceipt.frx":02D6
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   1185
         Left            =   600
         Picture         =   "frmReceipt.frx":02D8
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2040
      End
      Begin VB.Label Label4 
         Caption         =   "Confirm Password"
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
         Left            =   -74520
         TabIndex        =   6
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "New Password"
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
         Left            =   -74520
         TabIndex        =   5
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Old Password"
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
         Left            =   -74520
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Login Name"
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
         Left            =   -74520
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim query As String



Public Sub User()
Adodc1.RecordSource = "select * from login Where log in <> 'Admin'"
query = "select * from login"
Adodc1.RecordSource = query
Adodc1.CommandType = adCmdText
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Do Until Adodc1.Recordset.EOF
On Error Resume Next
cboUser.AddItem Adodc1.Recordset("log in")
Adodc1.Recordset.MoveNext
Loop
End Sub

Private Sub cboduser_Click()
Adodc1.RecordSource = "select * from login where log in = " & "'" & cboUser.Text & "'"
Adodc1.CommandType = adCmdText
Adodc1.Refresh
End Sub

Private Sub cboUser_Click()

Adodc1.RecordSource = "select * from login where log in = " & "'" & cboUser.Text & "'"
query = "select * from login"
Adodc1.RecordSource = query
Adodc1.CommandType = adCmdText
Adodc1.Refresh

End Sub

Private Sub cmdAdd_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub cmdcExut_Click()
Dim w As Integer
 w = MsgBox(" Are you wants to exit ?", vbYesNoCancel + vbQuestion)
  If w = vbYes Then
  Unload Me
End If
End Sub

Private Sub cmdChange_Click()
If txtOld.Text = Adodc1.Recordset("Password") And txtnew.Text = txtCon.Text Then

Adodc1.Recordset("Password") = txtCon.Text
Adodc1.Recordset.Update
MsgBox "Password Changed Successfully", vbInformation, "Security"
Else
MsgBox "Please enter correct Password", vbInformation, "Incorrect Password"
End If

End Sub

Private Sub cmdDlete_Click()
Adodc1.Recordset.Delete
Adodc1.Refresh
End Sub

Private Sub cmdexit_Click()
Dim e As Integer
 e = MsgBox(" Are you wants to exit ?", vbYesNoCancel + vbQuestion)
  If e = vbYes Then
  Unload Me
End If
End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
cmdprev.Enabled = True
If Adodc1.Recordset.EOF Then
cmdnext.Enabled = False
Adodc1.Recordset.MoveFirst

End If
End Sub

Private Sub cmdprev_Click()
 Adodc1.Recordset.MovePrevious
cmdnext.Enabled = True
If Adodc1.Recordset.BOF Then
  cmdprev.Enabled = False
   Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub cmdSave_Click()
Adodc1.Recordset("Password") = txtCpass.Text
Adodc1.Recordset.Update
End Sub

Private Sub Form_Load()
cmdnext.Enabled = False
Call User
End Sub

