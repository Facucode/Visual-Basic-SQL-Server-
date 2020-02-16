VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2520
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   3375
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtValor 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   3495
   End
   Begin VB.ComboBox cboCursos 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label lbValor 
      Caption         =   "Label1"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCursos_Change()
lbValor.Caption = cboCursos.Text


End Sub

Private Sub cboCursos_Click()
cboCursos_Change
End Sub

Private Sub cmdAgregar_Click()
cboCursos.AddItem txtValor.Text
End Sub

Private Sub cmdLimpiar_Click()
cboCursos.Clear

End Sub

Private Sub Command1_Click()
List1.AddItem Text1.Text



End Sub

Private Sub Command2_Click()
List1.Clear

End Sub

Private Sub Command3_Click()
List1.RemoveItem List1.ListIndex


End Sub

Private Sub Form_Load()
cboCursos.AddItem "Programación"

End Sub


