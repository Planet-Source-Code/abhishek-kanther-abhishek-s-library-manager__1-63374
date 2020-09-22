VERSION 5.00
Begin VB.Form deleteentry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Entry"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   Icon            =   "Form1_1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Book ID"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Delete Entry"
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
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "deleteentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

Dim con As Connection
Dim rs As Recordset
Dim sql12 As String
Set con = CreateObject("adodb.connection")
con.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "/books.mdb"
Set rs = New Recordset

If Text1.Text = "" And Text2.Text = "" Then
MsgBox "Please enter any one of the above"
Exit Sub
ElseIf Text1.Text <> "" Then
sql12 = "select * from books where name='" & Text1.Text & "'"
ElseIf Text2.Text <> "" Then
sql12 = "select * from books where bookid='" & Text2.Text & "'"
End If

rs.Open sql12, con, 1, 2

If rs.EOF Then
MsgBox "No Entry"
Else
rs.Delete
MsgBox "Entry Deleted successfully"
Text1.Text = ""
Text2.Text = ""


Me.Hide

End If
End Sub

Private Sub Form_Load()

End Sub
'Made By :Abhishek kanther
'emailID :kantherabhishek@gmail.com
