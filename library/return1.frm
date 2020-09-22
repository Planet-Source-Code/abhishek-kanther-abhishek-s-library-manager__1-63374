VERSION 5.00
Begin VB.Form return1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Return"
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "return1.frx":0000
         Left            =   1800
         List            =   "return1.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Category"
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
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "return1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Label2.Caption = Combo1.Text & "'s name"
Label3.Caption = Combo1.Text & "'s ID"

End Sub

Private Sub Command1_Click()
Dim rs As Recordset
Dim con As Connection
Set con = CreateObject("ADODB.Connection")
con.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "/books.mdb"
sql5 = "select * from issue where (bookid='" & Text2.Text & "' or bookname='" & Text1.Text & "') and category='" & Combo1.Text & "' "
Set rs = CreateObject("ADODB.Recordset")
rs.Open sql5, con, adOpenDynamic, adLockReadOnly


If Not rs.EOF Then

sql34 = "delete from issue where bookid='" & Text2.Text & "' or bookname='" & Text1.Text & "'"
'MsgBox sql34
con.Execute (sql34)


Set rs = Nothing
sql5 = "update books set issued='Not Issued' where (bookid= '" & Text2.Text & "' or name= '" & Text1.Text & "') "
con.Execute (sql5)
'Set rs = Nothing
con.Close
MsgBox "Your Book has been returned successfully"
Else
MsgBox "we are unable to return your book"


End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
return1.Hide

End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
Label2.Caption = Combo1.Text & "'s name"
Label3.Caption = Combo1.Text & "'s ID"

End Sub

'Made By :Abhishek kanther
'emailID :kantherabhishek@gmail.com
