VERSION 5.00
Begin VB.Form frmThumb 
   BorderStyle     =   0  'None
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2280
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   0
      Picture         =   "frmThumb.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2280
   End
End
Attribute VB_Name = "frmThumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Image1.Width = Me.Width
    Image1.Height = Me.Height
End Sub
