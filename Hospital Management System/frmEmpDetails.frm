VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form EmployeeDetails 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Employee Details"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5475
   LinkTopic       =   "Form2"
   ScaleHeight     =   4275
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1560
      Top             =   7920
      Width           =   2055
      _ExtentX        =   3625
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\Hms.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\Hms.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Emp_Details"
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Employee Details"
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Height          =   3615
         Left            =   6000
         TabIndex        =   29
         Top             =   2280
         Width           =   1575
         Begin VB.CommandButton cmdAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   240
            Picture         =   "frmEmpDetails.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdSave 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CommandButton CmdExit 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   2520
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   240
         TabIndex        =   22
         Top             =   6120
         Width           =   7335
         Begin VB.CommandButton cmdEfirst 
            BackColor       =   &H00FFC0FF&
            Caption         =   "First"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            Picture         =   "frmEmpDetails.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdEprev 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Previous"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1320
            Picture         =   "frmEmpDetails.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "previous nrecord"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdEnext 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Next"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2520
            Picture         =   "frmEmpDetails.frx":0CC6
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "next record"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdElast 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Last"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3720
            Picture         =   "frmEmpDetails.frx":1108
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtGo 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   5400
            TabIndex        =   24
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdGo 
            BackColor       =   &H00FFC0FF&
            Caption         =   "GO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   6240
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.Line Line1 
            X1              =   5160
            X2              =   5160
            Y1              =   120
            Y2              =   1200
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Official Details"
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   240
         TabIndex        =   15
         Top             =   3720
         Width           =   5535
         Begin VB.ComboBox cboDuty 
            DataField       =   "Duty"
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmEmpDetails.frx":154A
            Left            =   1920
            List            =   "frmEmpDetails.frx":1557
            TabIndex        =   20
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox cboPost 
            DataField       =   "Post"
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmEmpDetails.frx":1585
            Left            =   1920
            List            =   "frmEmpDetails.frx":1592
            TabIndex        =   19
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtSal 
            Appearance      =   0  'Flat
            DataField       =   "Salary"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   16
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Duty"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Salary"
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
            Left            =   240
            TabIndex        =   18
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Post"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.ComboBox cboEsex 
         Appearance      =   0  'Flat
         DataField       =   "Sex"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "frmEmpDetails.frx":15AF
         Left            =   2160
         List            =   "frmEmpDetails.frx":15B9
         TabIndex        =   14
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox TxtEage 
         Appearance      =   0  'Flat
         DataField       =   "Age"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtEmob 
         Appearance      =   0  'Flat
         DataField       =   "EmrgcyContactPhone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtEtel 
         Appearance      =   0  'Flat
         DataField       =   "Phone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtEadd 
         Appearance      =   0  'Flat
         DataField       =   "Address"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtEname 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblRecno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "EmployeeID"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblErec 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Record No:"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblAdd 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblPhone 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Phone No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Mobile No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   3240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "EmployeeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdAdd_Click()
txtEname.Enabled = True
txtEadd.Enabled = True
TxtEage.Enabled = True
cboEsex.Enabled = True
txtEtel.Enabled = True
txtEmob.Enabled = True
cboPost.Enabled = True
txtSal.Enabled = True
cboDuty.Enabled = True
Adodc1.Recordset.AddNew
cmdAdd.Enabled = False
cmdSave.Enabled = True
End Sub

Private Sub cmdEfirst_Click()
Adodc1.Recordset.MoveFirst
cmdEfirst.Enabled = False
cmdEnext.Enabled = True
cmdElast.Enabled = True
End Sub

Private Sub cmdElast_Click()
Adodc1.Recordset.MoveLast
cmdEprev.Enabled = True
cmdEfirst.Enabled = True
cmdElast.Enabled = False
End Sub

Private Sub cmdEnext_Click()
Adodc1.Recordset.MoveNext
cmdEfirst.Enabled = True
cmdEprev.Enabled = True
If Adodc1.Recordset.EOF Then
MsgBox "Sorry!no more record is there", vbInformation
Adodc1.RecordSource = "select * from Employee  "
Adodc1.CommandType = adCmdText
Adodc1.Refresh
cmdEnext.Enabled = False
End If
End Sub

Private Sub cmdEprev_Click()
Adodc1.Recordset.MovePrevious
cmdEnext.Enabled = True
cmdElast.Enabled = True
If Adodc1.Recordset.BOF Then
   MsgBox "Sorry!no previous record is there", vbInformation
Adodc1.RecordSource = "select * from Employee  "
Adodc1.CommandType = adCmdText
Adodc1.Refresh
cmdEprev.Enabled = False
End If
End Sub

Private Sub cmdGo_Click()
Adodc1.RecordSource = "select * from Employee Where EmployeeID = " & txtGo.Text
Adodc1.CommandType = adCmdText
Adodc1.Refresh
If lblRecno.Caption = "" Then
MsgBox "No record", vbInformation, "Result"
Adodc1.RecordSource = "select * from Employee  "
Adodc1.CommandType = adCmdText
Adodc1.Refresh
txtGo.Text = ""
End If
End Sub

Private Sub cmdSave_Click()
Adodc1.Recordset.Update
MsgBox "Save Successfully", vbInformation, "Save"
cmdSave.Enabled = False
cmdAdd.Enabled = True
End Sub

Private Sub Form_Load()
txtEname.Enabled = False
txtEadd.Enabled = False
TxtEage.Enabled = False
cboEsex.Enabled = False
txtEtel.Enabled = False
txtEmob.Enabled = False
cboPost.Enabled = False
txtSal.Enabled = False
cboDuty.Enabled = False
cmdSave.Enabled = False
End Sub
