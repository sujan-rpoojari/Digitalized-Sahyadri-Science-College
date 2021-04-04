VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   3075
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "SUBMIT"
      Height          =   495
      Left            =   8280
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply For More than One Course"
      Height          =   615
      Left            =   5040
      TabIndex        =   2
      Top             =   4080
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      Text            =   "COMBINATION"
      Top             =   3240
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5280
      TabIndex        =   0
      Text            =   "COURSE"
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Select Academic Course"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   3
      Top             =   1800
      Width           =   3135
   End
End
Attribute VB_Name = "Form8"
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
Form9.Show
End Sub

Private Sub Command2_Click()
Form12.Show
End Sub

Private Sub form_load()
Combo1.AddItem "BCA"
Combo1.AddItem "BSC"
End Sub

