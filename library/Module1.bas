Attribute VB_Name = "Module1"
Global b1 As Boolean
Global sql1 As String
Global flag1 As Boolean
Global category, cate As String

'Made By :Abhishek kanther
'emailID :kantherabhishek@gmail.com

Sub listbooks()
Form1.MSHFlexGrid1.Clear
Form1.MSHFlexGrid1.Visible = False
Dim con As New Connection
Dim rs As New Recordset

Set con = CreateObject("ADODB.Connection")
con.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "/books.mdb"
SQL = "select *from books where category='" & Form1.Combo1.Text & "'"
Set rs = New Recordset
rs.Open SQL, con, adOpenDynamic, adLockReadOnly
Set Form1.MSHFlexGrid1.DataSource = rs
Form1.Frame3.Visible = True
Form1.MSHFlexGrid1.Visible = True
Form1.MSHFlexGrid1.Refresh



End Sub



Sub listbooks1()
Form1.MSHFlexGrid1.Clear
Form1.MSHFlexGrid1.Visible = False
Dim con As New Connection
Dim rs As New Recordset

Set con = CreateObject("ADODB.Connection")
con.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "/books.mdb"
Set rs = New Recordset
rs.Open sql1, con, adOpenDynamic, adLockReadOnly
Set Form1.MSHFlexGrid1.DataSource = rs
Form1.Frame3.Visible = True
Form1.MSHFlexGrid1.Visible = True
Form1.MSHFlexGrid1.Refresh
End Sub

Sub listbooks3()
Dim con As Connection
Dim rs As Recordset
Set con = CreateObject("ADODB.Connection")
con.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "/books.mdb"
SQL = "select distinct category from books"


Set rs = CreateObject("ADODB.Recordset")
rs.Open SQL, con, 1, 2
Form2.Combo2.Clear
While Not rs.EOF
Form2.Combo2.AddItem rs("category")

rs.MoveNext
Wend
Form2.Combo2.AddItem "Add Category"

End Sub
