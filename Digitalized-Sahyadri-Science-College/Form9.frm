VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   4560
   LinkTopic       =   "Form9"
   ScaleHeight     =   3075
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "SUBMIT"
      Height          =   495
      Left            =   9120
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6120
      TabIndex        =   1
      Text            =   "COURSE"
      Top             =   3360
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6120
      TabIndex        =   0
      Text            =   "COMBINATION"
      Top             =   3840
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
      Left            =   5880
      TabIndex        =   2
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CON As New ADODB.Connection
Dim RS As New ADODB.Recordset
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


Private Sub Command2_Click()
Form12.Show
End Sub

Private Sub form_load()
Combo1.AddItem "BCA"
Combo1.AddItem "BSC"
End Sub

