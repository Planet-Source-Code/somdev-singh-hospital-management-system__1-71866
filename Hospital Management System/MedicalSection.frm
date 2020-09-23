VERSION 5.00
Begin VB.Form MedicalSec 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medical Section"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9255
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   5040
      Width           =   3015
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "MedicalSection.frx":0000
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "MedicalSection.frx":0018
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "MedicalSection.frx":001E
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "MedicalSection.frx":002A
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "MedicalSection.frx":0034
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "MedicalSection.frx":0041
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "MedicalSection.frx":004C
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "MedicalSection.frx":005A
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "MedicalSection.frx":006B
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "MedicalSection.frx":0076
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
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
      Left            =   5160
      Picture         =   "MedicalSection.frx":0087
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   3000
      Picture         =   "MedicalSection.frx":04C9
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   9120
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"MedicalSection.frx":090B
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   9255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Medical Section"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   240
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   9120
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   9120
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   9120
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      X1              =   3000
      X2              =   3000
      Y1              =   840
      Y2              =   5760
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      X1              =   6240
      X2              =   6240
      Y1              =   840
      Y2              =   5760
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "CARDIOLOGY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "GASTROENTEROLOGY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "PSYCHIATRY"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "NEPHROLOGY"
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
      Left            =   6360
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   9120
      Y1              =   5760
      Y2              =   5760
   End
End
Attribute VB_Name = "MedicalSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 SergicalSection.Show
End Sub

Private Sub Command2_Click()
 MedicalSec.Hide
End Sub

Private Sub Form_Load()
' Height = 6840
 'Width = 9375
End Sub
