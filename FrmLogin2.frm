VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form FrmLoginGuest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guest Login"
   ClientHeight    =   2190
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1293.924
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc UserInfo 
      Height          =   330
      Left            =   1695
      Top             =   1770
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      RecordSource    =   "select * from guest"
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
   Begin VB.CommandButton cmdregister 
      Caption         =   "&Register"
      Height          =   390
      Left            =   2100
      TabIndex        =   7
      Top             =   1590
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
      DataSource      =   "UserInfo"
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
      DataSource      =   "UserInfo"
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "New Users Click Here:"
      Height          =   270
      Left            =   180
      TabIndex        =   6
      Top             =   1680
      Width           =   1800
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
Attribute VB_Name = "FrmLoginGuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim usrname As String
Dim psword As String

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    usrname = "vineet" 'txtUserName.Text
    psword = "potdar" 'txtPassword.Text
    
    FrmLoginGuest.UserInfo.Recordset.MoveFirst
    Do Until FrmLoginGuest.UserInfo.Recordset.EOF
        If FrmLoginGuest.UserInfo.Recordset.Fields("username").Value = usrname _
             And FrmLoginGuest.UserInfo.Recordset.Fields("password").Value = psword Then
            MsgBox "Welcome! You may now view the database", , "Login Successful"
            FrmGuest.Show
            'FrmCountry.Show
            admin = False
            Unload Me
            Unload FrmLogin
            Exit Sub
        
        Else
            FrmLoginGuest.UserInfo.Recordset.MoveNext
        End If
    Loop
    
    MsgBox "Invalid Username/Password, try again!", , "Login Error"
    txtPassword.Text = ""
    txtUserName.Text = ""
    txtUserName.SetFocus
    SendKeys "{Home}+{End}"
End Sub

Private Sub cmdregister_Click()
    FrmRegister.Show
    Unload Me
End Sub


Private Sub txtPassword_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtUserName_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
