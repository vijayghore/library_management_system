VERSION 5.00
Begin VB.Form CreatingAMembership 
   BackColor       =   &H00DBBE71&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creating A Membership"
   ClientHeight    =   7815
   ClientLeft      =   5820
   ClientTop       =   1320
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CBYear 
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
      ItemData        =   "CreatingAMembership.frx":0000
      Left            =   3120
      List            =   "CreatingAMembership.frx":0016
      TabIndex        =   14
      Top             =   5880
      Width           =   3975
   End
   Begin VB.ComboBox CBClass 
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
      ItemData        =   "CreatingAMembership.frx":0050
      Left            =   3120
      List            =   "CreatingAMembership.frx":005D
      TabIndex        =   6
      Top             =   2640
      Width           =   4000
   End
   Begin VB.ComboBox CBCourse 
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
      ItemData        =   "CreatingAMembership.frx":0082
      Left            =   3120
      List            =   "CreatingAMembership.frx":008F
      TabIndex        =   4
      Top             =   2040
      Width           =   4000
   End
   Begin VB.TextBox TxtRollNo 
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
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   3240
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
      Height          =   480
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6720
      Width           =   2200
   End
   Begin VB.CommandButton BtnCreate 
      BackColor       =   &H00EC8D40&
      Caption         =   "Create"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Width           =   2200
   End
   Begin VB.TextBox TxtAddress 
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
      Height          =   1335
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   4440
      Width           =   4000
   End
   Begin VB.TextBox TxtMobNo 
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
      Height          =   495
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   10
      Top             =   3840
      Width           =   4000
   End
   Begin VB.TextBox TxtName 
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
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   4000
   End
   Begin VB.Label LblRollNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Roll No"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label LblCourse 
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
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
      Top             =   2040
      Width           =   1800
   End
   Begin VB.Label LblYear 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Top             =   5880
      Width           =   1800
   End
   Begin VB.Label LblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Top             =   4440
      Width           =   1800
   End
   Begin VB.Label LblMobNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Mob No"
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
      Top             =   3840
      Width           =   1800
   End
   Begin VB.Label LblClass 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
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
      Top             =   2640
      Width           =   1800
   End
   Begin VB.Label LblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Top             =   1440
      Width           =   1800
   End
   Begin VB.Label LblCreatingAMembership 
      BackStyle       =   0  'Transparent
      Caption         =   "Creating A Membership"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "CreatingAMembership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCreate_Click()
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Name, Class, BClass, Course, MobNo, Address, IssueDate, BTNumber, Validity, BRollNo, Year, RollNo As String
    Name = Me.TxtName.Text
    Class = Me.CBClass.Text
    Course = Me.CBCourse.Text
    RollNo = Me.TxtRollNo.Text
    MobNo = Me.TxtMobNo.Text
    Address = Me.TxtAddress.Text
    Year = Me.CBYear.Text
    IssueDate = Date
  
    If Me.CBYear.Text = "2017-18" Then
        Validity = "30/04/18"
    Else
        If Me.CBYear.Text = "2018-19" Then
            Validity = "30/04/19"
        Else
            If Me.CBYear.Text = "2019-20" Then
                Validity = "30/04/20"
            Else
                If Me.CBYear.Text = "2020-21" Then
                    Validity = "30/04/21"
                Else
                    If Me.CBYear.Text = "2021-22" Then
                        Validity = "30/04/22"
                    Else
                        If Me.CBYear.Text = "2022-23" Then
                            Validity = "30/04/23"
                        Else
                            If Me.CBYear.Text = "2023-24" Then
                                Validity = "30/04/24"
                            Else
                                If Me.CBYear.Text = "2024-25" Then
                                    Validity = "30/04/25"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If IsNumeric(Me.TxtName) Then
        MsgBox "Enter valid Name", vbOKOnly + vbCritical, "Error"
    Else
        If IsNumeric(Me.TxtRollNo.Text) Then
            If Len(Me.TxtMobNo.Text) < 10 Then
                MsgBox "Enter valid mobile no.", vbOKOnly + vbCritical, "Error"
            Else
                If Class = "Ist Year" Then
                    BClass = "1"
                Else
                    If Class = "IInd Year" Then
                        BClass = "2"
                    Else
                        If Class = "IIIrd Year" Then
                            BClass = "3"
                        End If
                    End If
                End If
                BRollNo = CStr(RollNo)
                BTNumber = Course & BClass & BRollNo
        
                Dim sql As String
                sql = "insert into CreatingMembership values('" & Name & "','" & Course & "','" & Class & "', " & RollNo & ", '" & MobNo & "','" & Address & "','" & Year & "','" & IssueDate & "','" & BTNumber & "','" & Validity & "')"
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
                    mb = MsgBox("Your BT Number is " & BTNumber, vbOKOnly + vbInformation, "Membership Created")
                    If mb = vbOK Then
                        Me.TxtAddress.Text = ""
                        Me.TxtMobNo.Text = ""
                        Me.TxtName.Text = ""
                        Me.TxtRollNo.Text = ""
                        Me.CBCourse.Text = ""
                        Me.CBClass.Text = ""
                        Me.CBYear.Text = ""
                    End If
                Else
                    MsgBox "Unable to Create MemberShip", vbOKCancel, "Membership"
                End If
            End If
        Else
            MsgBox "Enter Valid Roll Number", vbOKOnly + vbCritical, "Error"
        End If
    End If
End Sub
