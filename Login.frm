VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00DBBE71&
   Caption         =   "Login"
   ClientHeight    =   6555
   ClientLeft      =   3795
   ClientTop       =   2535
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnLogin 
      BackColor       =   &H00E97C23&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox TxtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   7920
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3360
      Width           =   2900
   End
   Begin VB.TextBox TxtUserName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7920
      TabIndex        =   3
      Top             =   2160
      Width           =   2900
   End
   Begin VB.Image ImgLogin 
      Height          =   4215
      Left            =   3120
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4170
   End
   Begin VB.Label LblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label LblUserName 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnLogin_Click()
    If Me.TxtUserName.Text = "" Or Me.TxtPassword.Text = "" Then
        MsgBox "Please enter your Username & Password", vbOKOnly + vbCritical, "Login Failed"
    Else
        If TxtUserName.Text = "123" And TxtPassword.Text = "123" Then
            MainMenu.Show
            Unload Me
        Else
            MsgBox "Please Enter Correct Username & Password", vbOKOnly + vbCritical, "Login Failed"
        End If
    End If
End Sub
