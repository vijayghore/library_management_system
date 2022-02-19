VERSION 5.00
Begin VB.Form Welcome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome"
   ClientHeight    =   7185
   ClientLeft      =   3720
   ClientTop       =   2340
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   13425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnGetStarted 
      BackColor       =   &H0080C0FF&
      Caption         =   "Get Started"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   3300
   End
   Begin VB.Label LblQuote 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Knowledge is free at the Library. Just bring your own container."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   2640
      TabIndex        =   0
      Top             =   1560
      Width           =   8415
   End
   Begin VB.Image Img1 
      Height          =   7140
      Left            =   0
      Picture         =   "Welcome.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13440
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnGetStarted_Click()
    Info.Show
    Unload Me
End Sub

