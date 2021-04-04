VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "end"
      Height          =   855
      Left            =   6120
      TabIndex        =   2
      Top             =   4560
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5160
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5160
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2640
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Combo2.Clear
If Combo1.Text = "BCA" Then
Combo2.AddItem "BCA"
ElseIf Combo1.Text = "BSC" Then
Combo2.AddItem "PMCS"
Combo2.AddItem "EMCS"
Combo2.AddItem "PCM"
Combo2.AddItem "CBZ"
Combo2.AddItem "PMG"
Combo2.AddItem "PME"
Combo2.AddItem "CBEZ"
Combo2.AddItem "CZBC"
Else
End If
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub form_load()
Combo1.AddItem "BCA"
Combo1.AddItem "BSC"
End Sub
