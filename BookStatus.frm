VERSION 5.00
Begin VB.Form BookStatus 
   BackColor       =   &H00DBBE71&
   Caption         =   "BookStatus"
   ClientHeight    =   3195
   ClientLeft      =   6420
   ClientTop       =   4860
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   8700
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
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   4335
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
      Height          =   615
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.ComboBox DdlSelStatus 
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
      ItemData        =   "BookStatus.frx":0000
      Left            =   3960
      List            =   "BookStatus.frx":000A
      TabIndex        =   3
      Text            =   "----------------Select----------------"
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label LblBTNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter BT Number"
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
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label LblSelStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Select status of book"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
End
Attribute VB_Name = "BookStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnSubmit_Click()
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    
    Dim Status As String, BTNumber As String
    
    Status = Me.DdlSelStatus.Text
    BTNumber = Me.TxtBTNumber
    
    con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb;persist security info=false"
    Set rs = con.Execute("select * from IssueBook where BTNumber = '" & BTNumber & "'")
    rs.Requery
    
    Dim AccNo As String
    AccNo = rs.Fields(0)
    
    Set rs1 = con.Execute("select * from ActualBookEntry where AccessionNo = " & AccNo)
    rs1.Requery
    
    Dim Price As String
    Price = rs1.Fields(4)
    con.Close
'-------------------------------------------------------------------------------------------------------------------------------------------
    Dim con1 As New ADODB.Connection
    Dim cmd As New ADODB.Command
    
    Dim sql As String
    sql = "insert into " & Status & " values('" & BTNumber & "', " & AccNo & ")"
    
    Dim cs As String
    cs = "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb"
    
    con1.ConnectionString = cs
    con1.Open
    cmd.ActiveConnection = con1
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    
    Dim n As Integer
    cmd.Execute n
    con1.Close
    If n = 1 Then
        mb = MsgBox("Book " & Me.DdlSelStatus.Text, vbOKOnly, "Message")
    Else
        mb = MsgBox("Book can't be Add", vbOKOnly, "Message")
    End If
'-------------------------------------------------------------------------------------------------------------------------------------------
    If Me.DdlSelStatus.Text = "Lost" Then
    
        Dim con2 As New ADODB.Connection
        Dim cmd1 As New ADODB.Command
    
        Dim sqldel As String
        sqldel = "delete from ActualBookEntry where AccessionNo = " & AccNo
        Dim cs1 As String
        cs1 = "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb"
    
        con2.ConnectionString = cs1
        con2.Open
        cmd1.ActiveConnection = con2
        cmd1.CommandType = adCmdText
        cmd1.CommandText = sqldel
    
        Dim n1 As Integer
        cmd1.Execute n1
        con2.Close
        
        MsgBox "Calculated Fine = " & Price, vbOKOnly + vbExclamation, "Fine"
        
    Else: Me.DdlSelStatus.Text = "Damage"
        Dim df As Integer
        df = (CInt(Price)) / 5
    
        MsgBox "Calculated Fine = " & df, vbOKOnly + vbExclamation, "Fine"
    End If
'-------------------------------------------------------------------------------------------------------------------------------------------
    Dim con3 As New ADODB.Connection
    Dim cmd2 As New ADODB.Command
    
    Dim sqldel1 As String
    RDate = Date
    sqldel1 = "insert into ReturnBook values(" & AccNo & ", '" & BTNumber & "' ,'" & RDate & "' )"

    Dim cs2 As String
    cs2 = "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb"
    
    con3.ConnectionString = cs2
    con3.Open
    cmd2.ActiveConnection = con3
    cmd2.CommandType = adCmdText
    cmd2.CommandText = sqldel1
    
    Dim n2 As Integer
    cmd2.Execute n2
    con3.Close
'-------------------------------------------------------------------------------------------------------------------------------------------
    Dim con4 As New ADODB.Connection
    Dim cmd3 As New ADODB.Command
    
    Dim sqldel2 As String
    sqldel2 = "delete from IssueBook where BTNumber = '" & BTNumber & "'"

    Dim cs3 As String
    cs3 = "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb"
    
    con4.ConnectionString = cs3
    con4.Open
    cmd3.ActiveConnection = con4
    cmd3.CommandType = adCmdText
    cmd3.CommandText = sqldel2
    
    Dim n3 As Integer
    cmd3.Execute n3
    con4.Close
End Sub

