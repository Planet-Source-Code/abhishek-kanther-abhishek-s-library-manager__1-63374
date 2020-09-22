VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5250
   ClientLeft      =   3045
   ClientTop       =   2040
   ClientWidth     =   6780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   16  'Merge Pen
   DrawStyle       =   5  'Transparent
   FillStyle       =   7  'Diagonal Cross
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   452
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   3120
      Top             =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Form1.Hide
End Sub
