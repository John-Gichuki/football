VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form FrmLoginAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrator Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdminInfo 
      Height          =   330
      Left            =   795
      Top             =   1185
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=football.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=football.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from admin"
      Caption         =   "AdminInfo"
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
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "FrmLoginAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valid As Boolean
Dim usrname As String
Dim psword As String
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    validate
    If valid = False Then
        Exit Sub
    End If
    usrname = txtUserName.Text
    psword = txtPassword.Text
    
    AdminInfo.Refresh
    AdminInfo.Recordset.MoveFirst
    Do Until FrmLoginAdmin.AdminInfo.Recordset.EOF
        If FrmLoginAdmin.AdminInfo.Recordset.Fields("username").Value = usrname _
             And FrmLoginAdmin.AdminInfo.Recordset.Fields("password").Value = psword Then
            MsgBox "Welcome! You may now view the database", , "Login Successful"
            FrmAdmin.Show
            admin = True
            Unload Me
            Unload FrmLogin
            Exit Sub
        Else
            FrmLoginAdmin.AdminInfo.Recordset.MoveNext
        End If
    Loop
    
    MsgBox "Invalid Username/Password, try again!", , "Login Error"
    txtPassword.Text = ""
    txtUserName.Text = ""
    txtUserName.SetFocus
    SendKeys "{Home}+{End}"
End Sub

Private Sub Form_Load()
    valid = False
End Sub

Private Sub txtPassword_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtUserName_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub validate()
    Dim x As Integer
    If txtPassword.Text = "" Or txtUserName.Text = "" Then
        x = MsgBox("Please fill in all the details", vbCritical, "Incomplete Form")
        valid = False
        Exit Sub
    End If
    valid = True
End Sub
