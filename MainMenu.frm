VERSION 5.00
Begin VB.Form MainMenu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Main Menu"
   ClientHeight    =   10650
   ClientLeft      =   120
   ClientTop       =   630
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Image Img1 
      Height          =   10605
      Left            =   0
      Picture         =   "MainMenu.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20475
   End
   Begin VB.Menu MenuAquisition 
      Caption         =   "Aquisition"
      Begin VB.Menu AquisitionAddBook 
         Caption         =   "Add Book"
      End
      Begin VB.Menu AquisitionBiA 
         Caption         =   "Books in Aquisition"
      End
   End
   Begin VB.Menu MenuCatloging 
      Caption         =   "Catloging"
      Begin VB.Menu CatlogingBookEntry 
         Caption         =   "Book Entry"
      End
      Begin VB.Menu CatlogingSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu CatlogingIssuedBookList 
         Caption         =   "Issued Book List"
      End
      Begin VB.Menu CatlogingReturnedBookList 
         Caption         =   "Returned Book List"
      End
   End
   Begin VB.Menu MenuCirculation 
      Caption         =   "Circulation"
      Begin VB.Menu CirculationCreatingMembership 
         Caption         =   "Creating Membership"
      End
      Begin VB.Menu CirculationSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu CirculationIssungBook 
         Caption         =   "Issuing A Book"
      End
      Begin VB.Menu CirculationReturnBook 
         Caption         =   "Returning A Book"
      End
      Begin VB.Menu CirculationSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu CirculationBookStatus 
         Caption         =   "Book Status"
      End
   End
   Begin VB.Menu MenuSerialControl 
      Caption         =   "Books Info"
      Begin VB.Menu SCTotalBooks 
         Caption         =   "Total Books"
      End
      Begin VB.Menu SCIssuedBooks 
         Caption         =   "Issued Books"
      End
      Begin VB.Menu SCAllBookList 
         Caption         =   "All Book List"
      End
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AquisitionAddBook_Click()
    Load AddBook
    AddBook.Show
End Sub

Private Sub AquisitionBiA_Click()
    Load BooksInAcquisition
    BooksInAcquisition.Show
End Sub

Private Sub CatlogingBookEntry_Click()
    BookEntry.Show
End Sub

Private Sub CatlogingIssuedBookList_Click()
    Load IssuedBookList
    IssuedBookList.Show
End Sub

Private Sub CatlogingReturnedBookList_Click()
    Load ReturnedBookList
    ReturnedBookList.Show
End Sub

Private Sub CatlogingSearch_Click()
    Load Search
    Search.Show
End Sub

Private Sub CirculationBookStatus_Click()
    Load BookStatus
    BookStatus.Show
End Sub

Private Sub CirculationCreatingMembership_Click()
    Load CreatingAMembership
    CreatingAMembership.Show
End Sub

Private Sub CirculationIssungBook_Click()
    Load IssueBook
    IssueBook.Show
End Sub

Private Sub CirculationReturnBook_Click()
    Load ReturnBook
    ReturnBook.Show
End Sub

Private Sub SCCopiesOfBook_Click()

End Sub

Private Sub SCAllBookList_Click()
    Load AllBooksList
    AllBooksList.Show
End Sub

Private Sub SCIssuedBooks_Click()
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim a As Integer
    
    conn.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb;persist security info=false"

    sql = "SELECT * FROM IssueBook"
    rs.Open sql, conn
    rs.Supports (adApproxPosition)
    a = rs.RecordCount
   ' strmsg = "the no of records is " & rs.RecordCount
   ' MsgBox strmsg
   MsgBox a
    End Sub

Private Sub SCTotalBooks_Click()
    Dim rs As New ADODB.Recordset
    Dim con As New ADODB.Connection
    
    con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb;persist security info=false"
    Set rs = con.Execute("select max(AccessionNo) from ActualBookEntry")
    rs.Requery
    TBook = rs.Fields(0) - 1000
    
    mb = MsgBox(TBook & " are available in the Library", vbOKOnly + vbInformation, "Total Books")
End Sub
