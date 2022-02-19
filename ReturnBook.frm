VERSION 5.00
Begin VB.Form ReturnBook 
   BackColor       =   &H00DBBE71&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return Book"
   ClientHeight    =   4215
   ClientLeft      =   7575
   ClientTop       =   2520
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton BtnReturned 
      BackColor       =   &H00EC8D40&
      Caption         =   "Returned"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox TxtBTNumber 
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
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   3000
   End
   Begin VB.TextBox TxtAccessionNo 
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
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   3000
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
      Left            =   345
      TabIndex        =   2
      Top             =   2040
      Width           =   2100
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
      Left            =   350
      TabIndex        =   1
      Top             =   1440
      Width           =   2100
   End
   Begin VB.Label LblReturnBook 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Return Book"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "ReturnBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnClear_Click()
    Me.TxtAccessionNo = ""
    Me.TxtBTNumber = ""
End Sub

Private Sub BtnReturned_Click()
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim AccessionNo, BTNumber, RDate As String
        
    AccessionNo = Me.TxtAccessionNo.Text
    BTNumber = Me.TxtBTNumber.Text
    RDate = Date
    
    Dim sql As String
    sql = "insert into ReturnBook values('" & AccessionNo & "', '" & BTNumber & "' , '" & RDate & "')"
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
        MsgBox "Book Returned..", vbOKOnly + vbInformation, "Message"
    Else
       MsgBox "Book can't be Return", vbOKOnly + vbCritical, "Message"
    End If
'----------------------------------------------------------------------------------------------------------------------------------------
    Dim sqldel As String
    sqldel = "delete from IssueBook where BTNumber = '" & BTNumber & "'"

    con.ConnectionString = cs
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = sqldel
    
    cmd.Execute n
    con.Close
'-----------------------------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim ISBNNo, Title, Author, Language, Publication, Location As String
    Dim Pages, Price, PublishingYear As Integer
    
    con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb;persist security info=false"
    Set rs = con.Execute("select * from ActualBookEntry where AccessionNo=" & AccessionNo)
    rs.Requery
    
    AccessionNo = rs.Fields(0)
    ISBNNo = rs.Fields(1)
    Title = rs.Fields(2)
    Author = rs.Fields(3)
    Pages = rs.Fields(4)
    Price = rs.Fields(5)
    Language = rs.Fields(6)
    Publication = rs.Fields(7)
    PublishingYear = rs.Fields(8)
    Location = rs.Fields(9)
    
    con.Close
    sqlnew = "insert into BookEntry values('" & AccessionNo & "', '" & ISBNNo & "', '" & Title & "', '" & Author & "', " & Pages & ", " & Price & ", '" & Language & "', '" & Publication & "', " & PublishingYear & ", '" & Location & "')"
    con.ConnectionString = cs
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = sqlnew
    
    cmd.Execute n
    con.Close
    If n = 1 Then
        MsgBox "Book Returned", vbOKOnly + vbInformation, "Returned Book"
    Else
        MsgBox "Book cannot be Returned", vbOKCancel + vbCritical, "Returned Book"
    End If
End Sub
