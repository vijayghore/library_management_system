VERSION 5.00
Begin VB.Form BookDetails 
   BackColor       =   &H00DBBE71&
   Caption         =   "Book Details"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnOk 
      BackColor       =   &H00EC8D40&
      Caption         =   "Ok"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1605
   End
   Begin VB.Label LblShowISBNNo 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   20
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label LblISBNNo 
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN Number"
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
      Left            =   240
      TabIndex        =   19
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label LblShowTitle 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   18
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label LblShowAuthor 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   17
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label LblShowPublishingYear 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   16
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Label LblShowPages 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   15
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label LblShowPrice 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   14
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label LblShowLanguage 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   13
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Label LblShowAccessionNo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   12
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label LblShowLocation 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   11
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Label LblShowPublication 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   10
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label LblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Title"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label LblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Author"
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
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label LblPublishingYear 
      BackStyle       =   0  'Transparent
      Caption         =   "Publishing Year"
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
      Left            =   240
      TabIndex        =   7
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label LblPages 
      BackStyle       =   0  'Transparent
      Caption         =   "Pages"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label LblPrice 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label LblLanguage 
      BackStyle       =   0  'Transparent
      Caption         =   "Language"
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
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label LblAccessionNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Accession No"
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
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label LblLocation 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label LblPublication 
      BackStyle       =   0  'Transparent
      Caption         =   "Publication"
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
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
End
Attribute VB_Name = "BookDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnOk_Click()
    Unload Me
End Sub

