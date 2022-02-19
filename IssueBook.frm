VERSION 5.00
Begin VB.Form IssueBook 
   BackColor       =   &H00DBBE71&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issue Book"
   ClientHeight    =   4110
   ClientLeft      =   7575
   ClientTop       =   2400
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtAccessionNo 
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   3000
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
      Height          =   480
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton BtnIssue 
      BackColor       =   &H00EC8D40&
      Caption         =   "Issue"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox TxtBTNumber 
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
      Left            =   2625
      TabIndex        =   4
      Top             =   1800
      Width           =   3000
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
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2130
   End
   Begin VB.Label LblBTNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "BT Number"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1905
   End
   Begin VB.Label LblIssueBook 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Book"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "IssueBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnClear_Click()
    Me.TxtAccessionNo = ""
    Me.TxtBTNumber = ""
End Sub

Private Sub BtnIssue_Click()
    Dim con As New Connection
    Dim cmd As New Command
    Dim AccessionNo As Integer, BTNumber, IDate As String
    
    AccessionNo = CInt(Me.TxtAccessionNo.Text)
    BTNumber = Me.TxtBTNumber.Text
    IDate = Date
    
    Dim sql As String
    sql = "insert into IssueBook values(" & AccessionNo & ", '" & BTNumber & "' , '" & IDate & "')"
    
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
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sqldel As String
    sqldel = "delete from BookEntry where AccessionNo = " & CInt(Me.TxtAccessionNo)
    
    con.ConnectionString = cs
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = sqldel
    
    Dim n1 As Integer
    cmd.Execute n1
    con.Close
    If n = 1 Then
        mb = MsgBox("Book Issued...", vbOKOnly, "Message")
    Else
        mb = MsgBox("Book can't be Issue....", vbOKOnly, "Message")
    End If
    
    Me.TxtAccessionNo = ""
    Me.TxtBTNumber = ""
End Sub

