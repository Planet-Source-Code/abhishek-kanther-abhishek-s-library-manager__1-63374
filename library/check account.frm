VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3450
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Check Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please enter your Name"
Else
Dim con As Connection
Dim rs As Recordset

Set con = CreateObject("adodb.connection")
con.Open "dsn=lib"

sql11 = "select * from issue where name='" & Text1.Text & "' "
Set rs = New Recordset
rs.Open sql11, con, adOpenStatic, adLockReadOnly

If rs.EOF Then
MsgBox "No Books are currently issued by you"
Exit Sub
End If
For i = 1 To 5
MSHFlexGrid1.ColWidth(i) = MSHFlexGrid1.Width / 5
Next i
Set MSHFlexGrid1.DataSource = rs
MSHFlexGrid1.Visible = True
End If

End Sub

