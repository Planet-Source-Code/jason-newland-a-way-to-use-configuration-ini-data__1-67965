VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demonstration Hash Table"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2220
      IntegralHeight  =   0   'False
      Left            =   105
      TabIndex        =   0
      Top             =   840
      Width           =   4440
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   660
      Left            =   195
      TabIndex        =   1
      Top             =   75
      Width           =   4290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    hMake "section1"
    hMake "section2"
    LoadINIToHash "section1", "testini.ini", "section1"
    LoadINIToHash "section2", "testini.ini", "section2"

    Me.List1.AddItem "[section1]"
    Me.List1.AddItem "item1=" & hGet("section1", "item1")
    Me.List1.AddItem "item2=" & hGet("section1", "item2")
    Me.List1.AddItem "[section2]"
    Me.List1.AddItem "config=" & hGet("section2", "config")
    Me.List1.AddItem "colors=" & hGet("section2", "colors")
    Me.List1.AddItem "test=" & hGet("section2", "test")
    '
    SaveHashToINI "section1", "testsave.ini", "section1"
    SaveHashToINI "section2", "testsave.ini", "section2"
End Sub
