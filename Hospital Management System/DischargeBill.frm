VERSION 5.00
Begin VB.Form DischargeBill 
   Caption         =   "Discharge Bill"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   4320
      TabIndex        =   30
      Text            =   "Text16"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   2520
      TabIndex        =   29
      Text            =   "Text15"
      Top             =   5760
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "By Cheque"
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
      Left            =   5520
      TabIndex        =   27
      Top             =   5160
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "By Cash"
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
      Left            =   3960
      TabIndex        =   26
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   1920
      TabIndex        =   25
      Text            =   "Text14"
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   2280
      TabIndex        =   23
      Text            =   "Text13"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   2280
      TabIndex        =   21
      Text            =   "Text12"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   2280
      TabIndex        =   19
      Text            =   "Text11"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   4440
      TabIndex        =   17
      Text            =   "Text10"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Text            =   "Text9"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Text            =   "Text8"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Text            =   "Text7"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Text            =   "Text6"
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6480
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5880
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5280
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label14 
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
      Left            =   3120
      TabIndex        =   31
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label13 
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
      Left            =   480
      TabIndex        =   28
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label12 
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
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label11 
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
      Left            =   480
      TabIndex        =   22
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Deposite"
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
      Left            =   480
      TabIndex        =   20
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label9 
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
      Left            =   480
      TabIndex        =   18
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Other Charges"
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
      Left            =   480
      TabIndex        =   15
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Days"
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
      Left            =   480
      TabIndex        =   14
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Bed No"
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
      Left            =   2400
      TabIndex        =   12
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Bed Code"
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
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
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
      Left            =   480
      TabIndex        =   8
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Patient Code"
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
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Receipt No"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "DischargeBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

