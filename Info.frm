VERSION 5.00
Begin VB.Form Info 
   BackColor       =   &H00DBBE71&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Info"
   ClientHeight    =   6825
   ClientLeft      =   3720
   ClientTop       =   2340
   ClientWidth     =   13260
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   13260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   3300
   End
   Begin VB.Label LblProfName 
      BackStyle       =   0  'Transparent
      Caption         =   "Prof. P. H. Sawlani"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   4920
      TabIndex        =   7
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Label LblGuidedBy 
      BackStyle       =   0  'Transparent
      Caption         =   "Guided By"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   855
      Left            =   5160
      TabIndex        =   6
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label LblName4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mangesh Deole"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   7200
      TabIndex        =   5
      Top             =   2880
      Width           =   3000
   End
   Begin VB.Label LblName3 
      BackStyle       =   0  'Transparent
      Caption         =   "Akshay Magar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   3600
      TabIndex        =   4
      Top             =   2880
      Width           =   3000
   End
   Begin VB.Label LblName2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vijay Ghore"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   7200
      TabIndex        =   3
      Top             =   2040
      Width           =   3000
   End
   Begin VB.Label LblName1 
      BackStyle       =   0  'Transparent
      Caption         =   "Akshay Ingole"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   3600
      TabIndex        =   2
      Top             =   2040
      Width           =   3000
   End
   Begin VB.Label LblSubmittedBy 
      BackStyle       =   0  'Transparent
      Caption         =   "Submitted by"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   855
      Left            =   4920
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label LblTiltle 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Library Management System"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   9615
   End
End
Attribute VB_Name = "Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnNext_Click()
    Load Login
    Login.Show
    Unload Me
End Sub
