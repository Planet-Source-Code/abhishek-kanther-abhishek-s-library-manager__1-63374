VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Library Manager 4.51"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10710
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Print List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   20
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Simple Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   6735
      Begin VB.OptionButton Option1 
         Caption         =   "Advance serach"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Simple Search"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "main.frx":0ECA
         Left            =   2160
         List            =   "main.frx":0ECC
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Select Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3255
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   6735
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3135
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   5
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Advance Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   6735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   741
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "Books"
      TabPicture(0)   =   "main.frx":0ECE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSFlexGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Magazines"
      TabPicture(1)   =   "main.frx":0EEA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Reports"
      TabPicture(2)   =   "main.frx":0F06
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Journals"
      TabPicture(3)   =   "main.frx":0F22
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "CDs"
      TabPicture(4)   =   "main.frx":0F3E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Floppies"
      TabPicture(5)   =   "main.frx":0F5A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3255
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Advance Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   6735
      Begin VB.OptionButton Option5 
         Caption         =   "Publication"
         Height          =   255
         Left            =   4080
         TabIndex        =   17
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Top             =   1080
         Width           =   2295
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Author"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Name"
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Book ID"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Please Select One of the following Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Label Label4 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Double click on the BookID to issue it"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
      Begin VB.Menu Add 
         Caption         =   "&Add new"
      End
      Begin VB.Menu return 
         Caption         =   "&Return"
      End
      Begin VB.Menu date 
         Caption         =   "&Check Return Date"
      End
      Begin VB.Menu checkacc1 
         Caption         =   "&Check Accounts"
      End
      Begin VB.Menu dete 
         Caption         =   "&Delete Entry"
      End
      Begin VB.Menu listall1 
         Caption         =   "&List All"
      End
      Begin VB.Menu ui 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim book As Object
Private Sub Add_Click()
Form2.Show


End Sub

Private Sub checkacc1_Click()
checkacc.Show

End Sub

Private Sub Combo1_Click()
Call listbooks
Label3.Visible = True
Label4.Visible = True
Command2.Visible = True



End Sub


Private Sub Command1_Click()

Dim f, d As String

Dim con As Connection
Dim rs As Recordset
Set con = CreateObject("adodb.connection")
con.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "/books.mdb"

If Option2.Value = True Then
f = "bookid"
d = Text1.Text
sql1 = "select * from books where " & f & "= '" & d & "' "
ElseIf Option3.Value = True Then
f = Option3.Caption
d = Text1.Text
sql1 = "select * from books where " & f & " like '%" & d & "%' "
ElseIf Option4.Value = True Then
f = Option4.Caption
d = Text1.Text
sql1 = "select * from books where " & f & " like '%" & d & "%' "
ElseIf Option5.Value = True Then
f = Option5.Caption
d = Text1.Text
sql1 = "select * from books where " & f & "='" & d & "' "

End If



'MsgBox f & ":" & d
'SQL = "select * from books where " & f & " like '%" & d & "%' "
'MsgBox sql1
Set rs = New Recordset
rs.Open sql1, con, adOpenStatic, adLockReadOnly
If rs.EOF Then
MsgBox "no book  found"
Else
Call listbooks1
Label3.Visible = True
Label4.Visible = True
Command2.Visible = True
End If
End Sub

Private Sub Command2_Click()

Set DataReport1.DataSource = MSHFlexGrid1.DataSource
DataReport1.PrintReport (True)

End Sub

Private Sub date_Click()
Form3.Show

End Sub

Private Sub dete_Click()
deleteentry.Show

End Sub

Private Sub Exit_Click()
For Each str1 In Forms
Unload str1
Next
End
End Sub

Private Sub Form_Load()
Label4.Visible = False
Label3.Visible = False
Command2.Visible = False


Frame2.Visible = False


MSHFlexGrid1.ColWidth(1) = MSHFlexGrid1.Width / 4
MSHFlexGrid1.ColWidth(2) = MSHFlexGrid1.Width / 5
MSHFlexGrid1.ColWidth(3) = MSHFlexGrid1.Width / 6
MSHFlexGrid1.ColWidth(4) = MSHFlexGrid1.Width / 5
Frame1.Visible = True
category = SSTab1.Caption
Call comb
'Combo1.ListIndex = 0

End Sub

Private Sub Form_Resize()

Frame1.Move Frame1.Left, Frame1.Top, Width * 0.9315, Frame1.Height
Frame2.Move Frame2.Left, Frame2.Top, Width * 0.9315, Frame2.Height
Frame4.Move Frame4.Left, Frame4.Top, Width * 0.9315, Frame4.Height
Frame3.Move Frame3.Left, Frame3.Top, Width * 0.9315, Frame3.Height
SSTab1.Move SSTab1.Left, SSTab1.Top, Width * 0.85269, SSTab1.Height
MSHFlexGrid1.Move MSHFlexGrid1.Left, MSHFlexGrid1.Top, Width * 0.9315, MSHFlexGrid1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
For Each str1 In Forms
Unload str1
Next
End Sub

Private Sub listall_Click()
listall.Show

End Sub


Private Sub listall1_Click()
listall.Show


End Sub

Private Sub MSHFlexGrid1_DblClick()
s = MSHFlexGrid1.Text
If MSHFlexGrid1.Col = 4 Then
s = MSHFlexGrid1.Text
ind = MSHFlexGrid1.Row
book1 = MSHFlexGrid1.TextMatrix(ind, 2)
Dim con As Connection
Dim rs As Recordset
Set con = CreateObject("ADODB.Connection")
con.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "/books.mdb"
SQL = "select * from issue where bookid='" & s & "'"
Set rs = con.Execute(SQL)
If Not rs.EOF Then
MsgBox " This book is issued to " & rs("name") & "  and Return date is " & rs("date of Return") & ""
Else
issue.Text1.Text = book1
issue.Text5.Text = s
issue.Show



End If



End If
'issue.Show
End Sub



Private Sub Option1_Click(Index As Integer)
Frame4.Visible = True

Frame1.Visible = False

Label3.Visible = False
Label4.Visible = False
Command2.Visible = False



MSHFlexGrid1.Clear
'SHFlexGrid1.Visible = False

'MsgBox " i am in it"
flag1 = True
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub return_Click()
return1.Show
Frame3.Visible = False



End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)



Select Case SSTab1.Tab
Case 1
Frame1.Visible = True
Frame4.Visible = False
Frame3.Visible = False
category = "Books"
Label3.Visible = False
Label4.Visible = False

Case 2
Frame3.Visible = False
category = "Books"
Label3.Visible = False
Label4.Visible = False
Case 3
Frame3.Visible = False
category = "Books"
Label3.Visible = False
Label4.Visible = False

Case 4
Frame3.Visible = False
category = "Books"
Label3.Visible = False
Label4.Visible = False

Case 5
Label3.Visible = False
Label4.Visible = False

Case Default
MSHFlexGrid1.Clear
Option1(0).Value = True
Option1(1).Value = False

Frame1.Visible = True
Frame4.Visible = False
Frame3.Visible = False

Call comb
Label3.Visible = False
Label4.Visible = False










End Select


End Sub

Sub comb()
Dim con As Connection
Dim rs As Recordset
Set con = CreateObject("ADODB.Connection")
con.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "/books.mdb"
SQL = "select distinct category from books"

Frame1.Visible = True


Set rs = CreateObject("ADODB.Recordset")
rs.Open SQL, con, 1, 2
Combo1.Clear
While Not rs.EOF
Combo1.AddItem rs("category")

rs.MoveNext
Wend

End Sub

'Made By :Abhishek kanther
'emailID :kantherabhishek@gmail.com
