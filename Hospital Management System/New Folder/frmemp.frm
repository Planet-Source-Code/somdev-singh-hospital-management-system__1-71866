VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3600
      Top             =   7800
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\typroject\Hospital.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\typroject\Hospital.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Employee"
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
      Caption         =   "Employee Details"
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   6000
         ScaleHeight     =   2985
         ScaleWidth      =   1425
         TabIndex        =   23
         Top             =   3000
         Width           =   1455
         Begin VB.CommandButton cmdAdd 
            Appearance      =   0  'Flat
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
            Height          =   615
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdSave 
            Appearance      =   0  'Flat
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
            Height          =   615
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton CmdExit 
            Appearance      =   0  'Flat
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
            Height          =   615
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1920
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Official Details"
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   240
         TabIndex        =   16
         Top             =   3720
         Width           =   5535
         Begin VB.ComboBox cboDuty 
            DataField       =   "Duty"
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmemp.frx":0000
            Left            =   1920
            List            =   "frmemp.frx":000D
            TabIndex        =   21
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox cboPost 
            DataField       =   "Post"
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmemp.frx":003B
            Left            =   1920
            List            =   "frmemp.frx":0048
            TabIndex        =   20
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtSal 
            Appearance      =   0  'Flat
            DataField       =   "Salary"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   17
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label3 
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
            TabIndex        =   22
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label8 
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
            TabIndex        =   19
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label7 
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
            TabIndex        =   18
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   240
         ScaleHeight     =   825
         ScaleWidth      =   7305
         TabIndex        =   15
         Top             =   6120
         Width           =   7335
         Begin VB.CommandButton cmdGo 
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
            Height          =   495
            Left            =   6360
            TabIndex        =   32
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txtGo 
            Height          =   615
            Left            =   5400
            TabIndex        =   31
            Top             =   120
            Width           =   615
         End
         Begin VB.CommandButton cmdEfirst 
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
            Height          =   615
            Left            =   240
            Picture         =   "frmemp.frx":0065
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdEnext 
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
            Height          =   615
            Left            =   2640
            Picture         =   "frmemp.frx":01B3
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "next record"
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdEprev 
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
            Height          =   615
            Left            =   1440
            Picture         =   "frmemp.frx":0301
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "previous nrecord"
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdElast 
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
            Height          =   615
            Left            =   3840
            Picture         =   "frmemp.frx":044F
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   120
            Width           =   1095
         End
         Begin VB.Line Line1 
            X1              =   5160
            X2              =   5160
            Y1              =   0
            Y2              =   840
         End
      End
      Begin VB.ComboBox cboEsex 
         Appearance      =   0  'Flat
         DataField       =   "Sex"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "frmemp.frx":059D
         Left            =   2160
         List            =   "frmemp.frx":05A7
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
         DataField       =   "EmployeeID"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblErec 
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
Attribute VB_Name = "Form2"
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
