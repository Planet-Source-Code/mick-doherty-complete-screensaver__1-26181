VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Saver Setup"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optShape 
      Caption         =   "Combination"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Square"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Circle"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox chkThumb 
      Caption         =   "Mini Preview"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.HScrollBar HScrollTime 
      Height          =   255
      LargeChange     =   100
      Left            =   240
      Max             =   1000
      Min             =   100
      SmallChange     =   10
      TabIndex        =   3
      Top             =   960
      Value           =   100
      Width           =   4215
   End
   Begin VB.HScrollBar HScrollSize 
      Height          =   255
      LargeChange     =   10
      Left            =   240
      Max             =   100
      Min             =   1
      TabIndex        =   1
      Top             =   360
      Value           =   20
      Width           =   4215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Interval:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Max Shape Size:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iOpt As Integer

Private Sub cmdOK_Click()
    SaveSetting "ScrnSave Base", "Settings", "Shape", iOpt
    SaveSetting "ScrnSave Base", "Settings", "Interval", HScrollTime.Value
    SaveSetting "ScrnSave Base", "Settings", "Max", HScrollSize.Value
    SaveSetting "ScrnSave Base", "Settings", "Preview", chkThumb.Value
    Unload Me
End Sub

Private Sub Form_Load()
    optShape(GetSetting("ScrnSave Base", "Settings", "Shape", "0")).Value = True
    HScrollTime.Value = GetSetting("ScrnSave Base", "Settings", "Interval", "100")
    HScrollSize.Value = GetSetting("ScrnSave Base", "Settings", "Max", "20")
    chkThumb.Value = GetSetting("ScrnSave Base", "Settings", "Preview", "1")
End Sub

Private Sub HScrollSize_Change()
    Label1 = "Max Shape Size: " & HScrollSize.Value & " %"
End Sub

Private Sub HScrollTime_Change()
    Label2 = "Interval: " & HScrollTime.Value / 1000 & " s"
End Sub

Private Sub optShape_Click(Index As Integer)
    iOpt = Index
End Sub
