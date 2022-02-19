VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BooksInAcquisition 
   BackColor       =   &H00FFFF80&
   Caption         =   "Books In Acquisition"
   ClientHeight    =   8430
   ClientLeft      =   2010
   ClientTop       =   1635
   ClientWidth     =   17085
   FillColor       =   &H80000008&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "BooksInAcquisition.frx":0000
   ScaleHeight     =   8430
   ScaleWidth      =   17085
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   1  'Align Top
      Height          =   10485
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17085
      _ExtentX        =   30136
      _ExtentY        =   18494
      _Version        =   393216
      BackColor       =   16777088
      BorderStyle     =   0
      DefColWidth     =   333
      ForeColor       =   255
      HeadLines       =   2
      RowHeight       =   26
      RowDividerStyle =   1
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Books in Acquisition"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   120
      Top             =   10440
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "BooksInAcquisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_Click()

End Sub

Private Sub Form_Load()
    Adodc1.ConnectionString = "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\LBMS.mdb;Persist Security Info=False"
    Adodc1.RecordSource = "select * from AddBook"
    Set DataGrid1.DataSource = Adodc1
End Sub
