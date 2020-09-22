VERSION 5.00
Begin VB.Form issue 
   Caption         =   "Issue"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Issue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   6
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Book ID"
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
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Date of Returning"
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
         TabIndex        =   10
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Date of Issue"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Book Name"
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
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As Boolean

Private Sub Command1_Click()
For i = 1 To Len(Text3.Text)
Chr1 = Mid(Text3.Text, i, 1)
If (Chr1 < "a" Or "z" < Chr1) And (Chr1 < "A" Or "Z" < Chr1) And Chr1 <> " " Then
b = True
End If
i = i + 1
Next i

If b Then
MsgBox "Invalid Character in the name Field"
End If

b = False

Dim con As Connection
Dim rs As Recordset

sql1 = "select * from issue"
Set con = CreateObject("Adodb.connection")
con.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "/books.mdb"
Set rs = CreateObject("ADODB.Recordset")
rs.Open sql1, con, 1, 2
rs.AddNew
rs("bookname") = Text1.Text
rs("bookid") = Text5.Text
rs("date of issue") = Text2.Text
rs("name") = Text3.Text
rs("date of return") = Text4.Text
rs("category") = "Books"
rs.Update
sql3 = "update books set issued='issued' where bookid='" & Trim(Text5.Text) & "'"

Set rs1 = con.Execute(sql3)

MsgBox "Your Book has been issued successfully"
Call listbooks
Unload Me









End Sub

Private Sub Form_Load()
Text2.Text = date
Text4.Text = DateAdd("d", 5, date)



End Sub

Private Sub Form_Resize()
Frame1.Move Frame1.Left, Frame1.Top, Width / 1.112, Height / 1.179
End Sub

'Made By :Abhishek kanther
'emailID :kantherabhishek@gmail.com
