VERSION 5.00
Begin VB.Form Form15 
   Caption         =   "Form15"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   4560
   LinkTopic       =   "Form15"
   ScaleHeight     =   10995
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   7320
      TabIndex        =   3
      Text            =   "BSC"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SUBMIT"
      Height          =   615
      Left            =   7320
      TabIndex        =   2
      Top             =   4200
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7320
      TabIndex        =   1
      Text            =   "BSC"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "SELECT ANOTHER COURSE"
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form12.Show
End Sub

Private Sub Combo1_Change()
combo
End Sub

Private Sub Form_Load()
Combo1.AddItem "PMCS"
Combo1.AddItem "EMCS"
Combo1.AddItem "PCM"
Combo1.AddItem "CBZ"
Combo1.AddItem "PMG"
Combo1.AddItem "PME"
Combo1.AddItem "CZBC"
End Sub

Private Sub Command2_Click()
Form12.Show
End Sub
