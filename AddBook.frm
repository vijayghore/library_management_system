VERSION 5.00
Begin VB.Form AddBook 
   BackColor       =   &H00DBBE71&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AddBook"
   ClientHeight    =   5355
   ClientLeft      =   750
   ClientTop       =   1560
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9555
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtCopies 
      Appearance      =   0  'Flat
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
      Left            =   3600
      TabIndex        =   6
      Top             =   3240
      Width           =   4500
   End
   Begin VB.CommandButton BtnClear 
      BackColor       =   &H00EC8D40&
      Caption         =   "Clear"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   2000
   End
   Begin VB.CommandButton BtnSubmit 
      BackColor       =   &H00EC8D40&
      Caption         =   "Submit"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   2000
   End
   Begin VB.TextBox TxtAuthor 
      Appearance      =   0  'Flat
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
      Left            =   3600
      TabIndex        =   4
      Top             =   2640
      Width           =   4500
   End
   Begin VB.TextBox TxtTitle 
      Appearance      =   0  'Flat
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
      Left            =   3600
      TabIndex        =   2
      Top             =   2040
      Width           =   4500
   End
   Begin VB.Label LblCopies 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Copies"
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
      Left            =   1320
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label LblAddBookInAcquisition 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Book in Acquisition"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   8415
   End
   Begin VB.Label LblEnterAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Author :"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   2640
      Width           =   2205
   End
   Begin VB.Label LblEnterTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Title :"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   2205
   End
End
Attribute VB_Name = "AddBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnClear_Click()
    TxtTitle.Text = ""
    TxtAuthor.Text = ""
    TxtCopies.Text = ""
End Sub

Private Sub BtnSubmit_Click()

    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    
    Dim Title, Author As String
    Dim Copies As Integer
    
    Title = TxtTitle.Text
    Author = TxtAuthor.Text
    Copies = CInt(Me.TxtCopies.Text)
    
    Dim sql As String
    sql = "insert into AddBook values('" & Title & "' , '" & Author & "'," & Copies & ")"
    
    Dim cs As String
    cs = "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb"
    
    con.ConnectionString = cs
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    
    Dim n As Integer
    cmd.Execute n
    con.Close
    If n = 1 Then
       mb = MsgBox("Book Added..", vbOKOnly + vbInformation, "Message")
    Else
        mb = MsgBox("Cannot add the book..", vbOKOnly + vbCritical, "Message")
    End If
    
End Sub

