VERSION 5.00
Begin VB.Form BookEntry 
   BackColor       =   &H00DBBE71&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Book Details"
   ClientHeight    =   8325
   ClientLeft      =   5340
   ClientTop       =   1680
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtISBNNo 
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
      Left            =   3720
      MaxLength       =   13
      TabIndex        =   4
      Top             =   1920
      Width           =   4000
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
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7560
      Width           =   2200
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
      Height          =   480
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7560
      Width           =   2200
   End
   Begin VB.TextBox TxtPublication 
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
      Left            =   3720
      TabIndex        =   16
      Top             =   5520
      Width           =   4000
   End
   Begin VB.TextBox TxtLocation 
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
      Left            =   3720
      TabIndex        =   20
      Top             =   6720
      Width           =   4000
   End
   Begin VB.TextBox TxtAccessionNo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   3720
      TabIndex        =   2
      Top             =   1320
      Width           =   4000
   End
   Begin VB.TextBox TxtLanguage 
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
      Left            =   3720
      TabIndex        =   14
      Top             =   4920
      Width           =   4000
   End
   Begin VB.TextBox TxtPrice 
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
      Left            =   3720
      TabIndex        =   12
      Top             =   4320
      Width           =   4000
   End
   Begin VB.TextBox TxtPages 
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
      Left            =   3720
      TabIndex        =   10
      Top             =   3720
      Width           =   4000
   End
   Begin VB.TextBox TxtPublishingYear 
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
      Left            =   3720
      TabIndex        =   18
      Top             =   6120
      Width           =   4000
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
      Left            =   3720
      TabIndex        =   8
      Top             =   3120
      Width           =   4000
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
      Left            =   3720
      TabIndex        =   6
      Top             =   2520
      Width           =   4000
   End
   Begin VB.Label LblISBNNo 
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN No"
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
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
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
      Left            =   1200
      TabIndex        =   15
      Top             =   5520
      Width           =   1695
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
      Left            =   1200
      TabIndex        =   19
      Top             =   6720
      Width           =   1335
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
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
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
      Left            =   1200
      TabIndex        =   13
      Top             =   4920
      Width           =   1575
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
      Left            =   1200
      TabIndex        =   11
      Top             =   4320
      Width           =   975
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
      Left            =   1200
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
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
      Left            =   1200
      TabIndex        =   17
      Top             =   6120
      Width           =   2415
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
      Left            =   1200
      TabIndex        =   7
      Top             =   3120
      Width           =   1935
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
      Left            =   1200
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label LblEnterBookDetails 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Book Details"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "BookEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnClear_Click()
    
    Me.TxtAccessionNo.Text = ""
    Me.TxtISBNNo.Text = ""
    Me.TxtTitle.Text = ""
    Me.TxtAuthor.Text = ""
    Me.TxtPages.Text = ""
    Me.TxtPrice.Text = ""
    Me.TxtLanguage.Text = ""
    Me.TxtPublication.Text = ""
    Me.TxtPublishingYear.Text = ""
    Me.TxtLocation.Text = ""
    
    Dim rs As New ADODB.Recordset
    Dim con As New ADODB.Connection
    
    con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb;persist security info=false"
    Set rs = con.Execute("select max(AccessionNo) from ActualBookEntry")
    rs.Requery
    TxtAccessionNo.Text = rs.Fields(0) + 1
    
End Sub

Private Sub BtnSubmit_Click()
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim ISBNNo, Title, Author, Language, Publication, Location As String
    Dim AccessionNo, Pages, Price, PublishingYear As Integer
    
    AccessionNo = BookEntry.TxtAccessionNo.Text
    ISBNNo = BookEntry.TxtISBNNo.Text
    Title = BookEntry.TxtTitle.Text
    Author = BookEntry.TxtAuthor.Text
    Pages = CInt(BookEntry.TxtPages.Text)
    Price = CInt(BookEntry.TxtPrice.Text)
    Language = BookEntry.TxtLanguage.Text
    Publication = BookEntry.TxtPublication.Text
    PublishingYear = CInt(BookEntry.TxtPublishingYear.Text)
    Location = BookEntry.TxtLocation.Text
    
    Dim sql As String
    sql = "insert into BookEntry values(" & AccessionNo & ", '" & ISBNNo & "', '" & Title & "', '" & Author & "', " & Pages & ", " & Price & ", '" & Language & "', '" & Publication & "', " & PublishingYear & ", '" & Location & "')"
    
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
'--------------------------------------------------------------------------------------------------------------------------------------
    Dim sqldel As String
    sqldel = "delete from AddBook where Title = '" & Title & "' AND Author= '" & Author & "'"
    
    con.ConnectionString = cs
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = sqldel
    
    Dim n1 As Integer
    cmd.Execute n1
    con.Close
'-------------------------------------------------------------------------------------------------------------------------------------
    Dim sqladd As String
    sqladd = "insert into ActualBookEntry values(" & AccessionNo & ", '" & ISBNNo & "', '" & Title & "', '" & Author & "', " & Pages & ", " & Price & ", '" & Language & "', '" & Publication & "', " & PublishingYear & ", '" & Location & "')"

    con.ConnectionString = cs
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = sqladd
    
    Dim n2 As Integer
    cmd.Execute n2
    con.Close
    If n = 1 Then
        MsgBox "Book Added in the Database", vbOKOnly + vbInformation, "Add Book"
    Else
        MsgBox "Unable to Add in the Database", vbOKCancel, "Add Book"
    End If
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    Dim con As New ADODB.Connection
    
    con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb;persist security info=false"
    Set rs = con.Execute("select max(AccessionNo) from ActualBookEntry")
    rs.Requery
    TxtAccessionNo.Text = rs.Fields(0) + 1
End Sub
