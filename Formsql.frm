VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CloseBtn 
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton AddBtn 
      Caption         =   "New User"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox passwordtxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox usernametxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "Formsql.frx":0000
      Height          =   1455
      Left            =   600
      TabIndex        =   0
      Top             =   2760
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   5
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0)._NumMapCols=   4
      _Band(0)._MapCol(0)._Name=   "id"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Alignment=   7
      _Band(0)._MapCol(1)._Name=   "RefId"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "Username"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "Password"
      _Band(0)._MapCol(3)._RSIndex=   3
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4440
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SQLServer;Initial Catalog=Prueba"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SQLServer;Initial Catalog=Prueba"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tabla1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddBtn_Click()
If usernametxt.Text <> "" And passwordtxt.Text <> "" Then
    Dim v As String
    Randomize
    v = Int(Rnd * 999) * 100
    Adodc1.Refresh
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields("RefId") = v
    Adodc1.Recordset.Fields("Username") = usernametxt.Text
    Adodc1.Recordset.Fields("Password") = passwordtxt.Text
    
    Adodc1.Recordset.Update
    Adodc1.Recordset.Close

    Unload Me
    MsgBox "New user is added in the database", vbOKOnly + vbInformation, "New User"
    Me.Show

Else
    Unload Me
    MsgBox "Please complete the form", vbOKOnly + vbCritical, "New User"
    Me.Show

End If

End Sub

Private Sub CloseBtn_Click()
Unload Me

End Sub

Private Sub Form_Load()

Adodc1.RecordSource = "Select * from tabla1"
Adodc1.Refresh
End Sub

