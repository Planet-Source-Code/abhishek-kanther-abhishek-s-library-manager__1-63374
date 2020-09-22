VERSION 5.00
Begin VB.Form welcome 
   ClientHeight    =   4785
   ClientLeft      =   2055
   ClientTop       =   2055
   ClientWidth     =   6765
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "welcome.frx":0000
   ScaleHeight     =   4785
   ScaleWidth      =   6765
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2760
      Top             =   2280
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
welcome.Enabled = False

End Sub

Private Sub Timer1_Timer()
welcome.Hide
Form1.Show
Timer1.Enabled = False
End Sub
'Made By :Abhishek kanther
'emailID :kantherabhishek@gmail.com
