VERSION 5.00
Begin VB.Form Search 
   BackColor       =   &H00DBBE71&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtEnterAccessionNo 
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
      Height          =   510
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton BtnSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Search"
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
      Left            =   4320
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Shape RoundedRectangle 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   6135
   End
   Begin VB.Label LblEnterAccessionNo 
      BackColor       =   &H00FF8080&
      Caption         =   " Enter Accession No."
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
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnSearch_Click()
    Dim n, MaxBook As Integer
    n = Me.TxtEnterAccessionNo.Text
    
    Dim rs1 As New ADODB.Recordset
    Dim con1 As New ADODB.Connection
    
    con1.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb;persist security info=false"
    Set rs1 = con1.Execute("select max(AccessionNo) from ActualBookEntry")
    rs1.Requery
    MaxBook = rs1.Fields(0) + 1
    
    If n = "" Then
        MsgBox "Please enter Accession No", vbOKOnly + vbInformation, "Search"
    Else
        If n > 1000 And n < MaxBook Then
            Dim rs As New ADODB.Recordset
            Dim con As New ADODB.Connection
    
            con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb;persist security info=false"
            Set rs = con.Execute("select * from ActualBookEntry where AccessionNo=" & n)
            rs.Requery
    
            BookDetails.LblShowAccessionNo.Caption = rs.Fields(0)
            BookDetails.LblShowISBNNo.Caption = rs.Fields(1)
            BookDetails.LblShowTitle.Caption = rs.Fields(2)
            BookDetails.LblShowAuthor.Caption = rs.Fields(3)
            BookDetails.LblShowPages.Caption = rs.Fields(4)
            BookDetails.LblShowPrice.Caption = rs.Fields(5)
            BookDetails.LblShowLanguage.Caption = rs.Fields(6)
            BookDetails.LblShowPublication.Caption = rs.Fields(7)
            BookDetails.LblShowPublishingYear.Caption = rs.Fields(8)
            BookDetails.LblShowLocation.Caption = rs.Fields(9)
    
            Load BookDetails
            BookDetails.Show
        Else
            mb = MsgBox("Enter Valid Accession Number", vbOKOnly + vbExclamation, "Search")
            If mb = vbOK Then
                Me.TxtEnterAccessionNo.Text = ""
            End If
        End If
    End If
End Sub
