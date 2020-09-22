VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Item"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "return.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.Frame Frame2 
         Height          =   5295
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   4815
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   1560
            TabIndex        =   16
            Top             =   3960
            Width           =   1695
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1560
            TabIndex        =   13
            Top             =   3360
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1560
            TabIndex        =   11
            Top             =   2640
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Cancel          =   -1  'True
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2040
            TabIndex        =   10
            Top             =   4800
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Add"
            Height          =   375
            Left            =   720
            TabIndex        =   9
            Top             =   4800
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1560
            TabIndex        =   8
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1560
            TabIndex        =   6
            Top             =   1200
            Width           =   1695
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1680
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Price"
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
            Left            =   360
            TabIndex        =   15
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Remarks"
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
            Left            =   360
            TabIndex        =   14
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Publication"
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
            Left            =   240
            TabIndex        =   12
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Author"
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
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Name"
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
            Left            =   360
            TabIndex        =   5
            Top             =   1200
            Width           =   975
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
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "return.frx":0ECA
         Left            =   1920
         List            =   "return.frx":0EE0
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b, c As Boolean
Dim r As Integer
Dim p, p1, q, q1 As String
Dim price As Double
Dim publication As String
Dim remarks As String
Dim z As String





Private Sub Combo1_Click()
Call listbooks3

End Sub

Private Sub Combo2_Click()
If Combo2.Text = "Add Category" Then
category1.Show

Else
cate = Combo2.Text


End If


End Sub

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please enter book name"
Text1.SetFocus
Exit Sub
End If


If Text2.Text = "" Then
MsgBox "Please enter Author name"
Text2.SetFocus
Exit Sub
End If


If Text3.Text <> "" Then
publication = Text4.Text
Else
publication = "-"

End If

If Text4.Text <> "" Then
remarks = Text4.Text
Else
remarks = "No Remarks"

End If

If Text5.Text <> "" Then
If IsNumeric(Text5.Text) = False Then
MsgBox "please enter numeric value in Price Field"
Text5.SetFocus
Exit Sub
Else
price = Text5.Text
End If
Else
price = 0
End If






p = ""
p1 = ""
q = ""

Dim g, g1 As Integer
s = Combo2.Text

g1 = Len(s)
If g1 > 2 Then
p = Left(s, 3)
ElseIf g1 <= 2 Then
p = Left(s, g1)
End If


q = p

For g = Len(s) To 1 Step -1
chr5 = Mid(s, g, 1)

If chr5 = " " Then
r = g

p1 = Mid(s, r + 1, 1)
q = p + p1
Exit For
End If

Next g

'MsgBox q

Dim ss As Connection
Dim rq As Recordset
Set ss = CreateObject("ADODB.Connection")
ss.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "/books.mdb"
sql56 = "select * from " & Combo1.Text & " where category='" & cate & "' order by bookid"
'MsgBox sql56
Set rq = CreateObject("ADODB.Recordset")

rq.Open sql56, ss, 1, 2
If Not rq.EOF Then


rq.MoveLast
q1 = rq("bookid")
Else
z = "01"
End If






Set rs = Nothing

'MsgBox q1

For d = 1 To Len(q1)
w = Mid(q1, d, 1)
If IsNumeric(w) Then
tt = CStr(tt) & CStr(w)
End If


Next d
If Len(tt) <= 1 Then
z = "0" & tt + 1

Else
z = tt + 1
End If

oo = q & z
sql78 = "insert into " & Combo1.Text & " values('" & cate & "','" & Text1.Text & "','" & Text2.Text & "','" & oo & "','Not Issued','" & publication & "','" & price & "','" & remarks & " ')"
'MsgBox sql78
ss.Execute sql78
MsgBox "New item has been inserted successfully" & vbCrLf & "Item Code :" & oo
Set ss = Nothing

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Form2.Hide



End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Form2.Hide


End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
Combo2.ListIndex = 0
price = 0
publication = "-"
remarks = "No Remarks"

End Sub
'Made By :Abhishek kanther
'emailID :kantherabhishek@gmail.com
