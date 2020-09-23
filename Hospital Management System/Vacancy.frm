VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Vacancy 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10770
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   7920
      Top             =   5640
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
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Availability"
      Caption         =   "Adodc2"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Vacancy.frx":0000
      Height          =   4575
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8070
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Search"
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
      Left            =   6240
      Picture         =   "Vacancy.frx":0014
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00C0C0C0&
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
      Left            =   4920
      Picture         =   "Vacancy.frx":0456
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   3975
      Left            =   7800
      TabIndex        =   1
      Top             =   960
      Width           =   2835
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ok"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox cboWards 
         Height          =   315
         ItemData        =   "Vacancy.frx":0898
         Left            =   120
         List            =   "Vacancy.frx":08A5
         TabIndex        =   11
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ComboBox cboGender 
         Height          =   315
         ItemData        =   "Vacancy.frx":08B2
         Left            =   120
         List            =   "Vacancy.frx":08BC
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   480
         Top             =   3600
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\Hospital Management System\HMS.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Availability"
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
      Begin VB.OptionButton optGender_Ward 
         BackColor       =   &H00C0C0C0&
         Caption         =   "By Gender and Ward"
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
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton optWard 
         BackColor       =   &H00C0C0C0&
         Caption         =   "By Ward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optGender 
         BackColor       =   &H00C0C0C0&
         Caption         =   "By Gender"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   3840
         Left            =   0
         Picture         =   "Vacancy.frx":08CF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2820
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Vacancy.frx":1A13
      Height          =   4575
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8070
      _Version        =   393216
      BackColorBkg    =   16761024
      Appearance      =   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"Vacancy.frx":1A27
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
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   7575
   End
   Begin VB.Line Line1 
      X1              =   7680
      X2              =   7680
      Y1              =   0
      Y2              =   7560
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "List Of Vacant Beds"
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
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Vacancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim query As String

Private Sub cboGender_Click()

If optGender_Ward.Value = False Then

If optGender.Value = True Then
    query = "select * from Availability where Reserved='FALSE' and Gender='" & cboGender.Text & "'"
    Adodc.RecordSource = query
    Adodc.Refresh

    'Set cboGender.RowSource = Adodc
    'cboGender.ListField = "Gender"
    'cboGender.Refresh
Else
If optGender.Value = False Then

End If
End If
End If

 'cboGender.Text = ""
 'cboGender.Visible = True
 'cboWards.Visible = False
 'cboWards.Text = ""
 cboGender.SetFocus
If optGender_Ward.Value = True Then
'If optGender.Value = True And optWard.Value = True Then
    query = "select * from Availability where Reserved='FALSE' and Gender='" & cboGender.Text & "' and Wards='" & cboWards.Text & "'"
    Adodc.RecordSource = query
    Adodc.Refresh
End If
'End If
End Sub

Private Sub cboWards_Click()
If optGender_Ward.Value = False Then
If optWard.Value = True Then
    query = "select * from Availability where Reserved='FALSE' and Wards='" & cboWards.Text & "'"
    Adodc.RecordSource = query
    Adodc.Refresh

    'Set cboGender.RowSource = Adodc
    'cboGender.ListField = "Gender"
    'cboGender.Refresh
Else
If optWard.Value = False Then
MsgBox "Select ward"
End If
End If
End If

If optGender_Ward.Value = True Then
'If optGender.Value = True And optWard.Value = True Then
    query = "select * from Availability where Reserved='FALSE' and Gender='" & cboGender.Text & "' and Wards='" & cboWards.Text & "'"
    Adodc.RecordSource = query
    Adodc.Refresh
'End If
End If

End Sub
Private Sub cmdSearch_Click()
 Frame1.Visible = True
 'Width = 10785
 cboGender.Text = ""
 cboWards.Text = ""
End Sub

Private Sub cmdEnd_Click()
 Vacancy.Hide
End Sub

Private Sub Form_Load()
Frame1.Visible = False
cboGender.Visible = False
cboWards.Visible = False
 'Width = 7770
 'Height = 7065
 Adodc1.RecordSource = "SELECT * FROM Availability WHERE Reserved = '" & False & "'"
 'Adodc1.Refresh
End Sub
Private Sub optWard_Click()
 cboWards.Visible = True
 cboGender.Visible = False
 cboWards.Text = ""
 cboWards.SetFocus
End Sub

Private Sub optGender_Click()
cboGender.Visible = True
cboWards.Visible = False
End Sub

Private Sub optGender_Ward_Click()
 cboWards.Visible = True
 cboGender.Visible = True
 cboGender.Text = ""
 cboWards.Text = ""
 cboGender.SetFocus

 cboWards.Enabled = True
End Sub


