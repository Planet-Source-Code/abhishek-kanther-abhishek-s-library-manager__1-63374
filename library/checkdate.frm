VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Date"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "checkdate.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BorderStyle     =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label Label2 
      Caption         =   "Return Date is today"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "Return Date expired"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   4080
      Width           =   135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Set DataReport2.DataSource = MSHFlexGrid1.DataSource
DataReport2.Show

'DataReport2.PrintReport (True)

End Sub



Private Sub Form_Load()
MSHFlexGrid1.ColWidth(1) = MSHFlexGrid1.Width / 4
MSHFlexGrid1.ColWidth(2) = MSHFlexGrid1.Width / 4
MSHFlexGrid1.ColWidth(3) = MSHFlexGrid1.Width / 3.5
Dim con11 As Connection
Dim rs As Recordset
Dim dt As Date
Dim dt1 As Date
Dim i As Integer

i = 1
MSHFlexGrid1.TextMatrix(0, 0) = "Name"
MSHFlexGrid1.TextMatrix(0, 1) = "Book ID"
MSHFlexGrid1.TextMatrix(0, 2) = "Book Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Date Of Returning"

Set con11 = CreateObject("adodb.connection")
con11.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "/books.mdb"
dt = date
day2 = Day(date)
mon2 = Month(date)
year2 = Year(date)

sql11 = "select * from issue"
Set rs = New Recordset
rs.Open sql11, con11, adOpenStatic, adLockReadOnly

While Not rs.EOF
dt1 = rs(5)
day1 = Day(dt1)
mon1 = Month(dt1)
year1 = Year(dt1)

If year2 > year1 Then
'MsgBox "Greater"
flag1 = True
ElseIf mon2 > mon1 Then
'MsgBox "Greater"
flag1 = True
ElseIf day2 > day1 Then
'MsgBox "Greater"
flag1 = True
Else
'MsgBox "smaller"
flag1 = False
End If

If flag1 Then

MSHFlexGrid1.TextMatrix(i, 0) = rs(4)
MSHFlexGrid1.TextMatrix(i, 1) = rs(2)
MSHFlexGrid1.TextMatrix(i, 2) = rs(1)
MSHFlexGrid1.TextMatrix(i, 3) = rs(5)

MSHFlexGrid1.Row = i
MSHFlexGrid1.Col = 0
If day1 < day2 Then


MSHFlexGrid1.CellBackColor = vbRed
MSHFlexGrid1.CellForeColor = vbWhite
End If

i = i + 1
MSHFlexGrid1.AddItem " ", i
End If

rs.MoveNext
Wend
End Sub

'Made By :Abhishek kanther
'emailID :kantherabhishek@gmail.com
