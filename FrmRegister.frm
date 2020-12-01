VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form FrmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration Form"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2700
      Top             =   5235
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin VB.TextBox txtclub 
      DataSource      =   "Adodc1"
      Height          =   330
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   13
      Top             =   3585
      Width           =   3030
   End
   Begin VB.TextBox txtpwd 
      DataSource      =   "Adodc1"
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "(max 10 characters)"
      Top             =   3120
      Width           =   3030
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1590
      TabIndex        =   16
      ToolTipText     =   "GO BACK TO LOGIN PAGE"
      Top             =   4725
      Width           =   1395
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2430
      TabIndex        =   15
      ToolTipText     =   "RESET FORM"
      Top             =   4065
      Width           =   1395
   End
   Begin VB.CommandButton cmdsubmit 
      Caption         =   "&Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   795
      TabIndex        =   14
      ToolTipText     =   "SUBMIT REGISTRATION"
      Top             =   4065
      Width           =   1395
   End
   Begin VB.TextBox txtadd 
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   1560
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   645
      Width           =   3030
   End
   Begin VB.TextBox txtname 
      DataSource      =   "Adodc1"
      Height          =   330
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   1
      Top             =   165
      Width           =   3030
   End
   Begin VB.TextBox txtuname 
      DataSource      =   "Adodc1"
      Height          =   330
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   9
      Text            =   "(max 10 characters)"
      Top             =   2655
      Width           =   3030
   End
   Begin VB.TextBox txtemail 
      DataSource      =   "Adodc1"
      Height          =   330
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   7
      Top             =   2190
      Width           =   3030
   End
   Begin VB.TextBox txttelno 
      DataSource      =   "Adodc1"
      Height          =   330
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1740
      Width           =   3030
   End
   Begin VB.Label Label8 
      Caption         =   "* fields are compulsory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   17
      Top             =   5325
      Width           =   2265
   End
   Begin VB.Label Label7 
      Caption         =   "Favourite Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   3615
      Width           =   1410
   End
   Begin VB.Label Label6 
      Caption         =   "Password *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   3180
      Width           =   1185
   End
   Begin VB.Label Label5 
      Caption         =   "Username *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   2700
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "Email-id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   2250
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "Tel No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   1785
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Name *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   1005
   End
End
Attribute VB_Name = "FrmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valid As Boolean

Private Sub validate()
    Dim x As Integer
    valid = True
    
    If txtname.Text = "" Then
        valid = False
        MsgBox "Please enter your name", , "Validation Fail"
        Exit Sub
    End If
    
    If txtuname.Text = "" Then
        valid = False
        MsgBox "Please enter a username", , "Validation Fail"
        Exit Sub
    End If
    If txtpwd.Text = "" Then
        valid = False
        MsgBox "Please enter a password", , "Validation Fail"
        Exit Sub
    End If
    
    If (IsNumeric(txttelno.Text) = False) Then
        valid = False
        x = MsgBox("Please enter valid telephone number", vbCritical, "Invalid Telephone Number")
        Exit Sub
    End If

    Adodc1.Refresh
    Do While Adodc1.Recordset.EOF = False
        If Adodc1.Recordset.Fields("username").Value = txtuname.Text Then
            valid = False
            x = MsgBox("Username already exists. Select a different username", vbCritical, "Username Taken")
            Exit Sub
        Else
            Adodc1.Recordset.MoveNext
        End If
    Loop
    
    If (InStr(txtemail.Text, "@") = 0) Then
        x = MsgBox("Please enter a valid email id", vbCritical, "Invalid Email-id")
        valid = False
        Exit Sub
    Else
        If (InStr(txtemail.Text, ".") = 0) Then
            x = MsgBox("Please enter a valid email id", vbCritical, "Invalid Email-id")
            valid = False
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdback_Click()
    FrmLoginGuest.Show
    Unload Me
End Sub

Private Sub cmdreset_Click()
    txtadd.Text = ""
    txtname.Text = ""
    txtuname.Text = ""
    txtpwd.Text = ""
    txttelno.Text = ""
    txtemail.Text = ""
    txtclub.Text = ""
End Sub

Private Sub cmdsubmit_Click()
    validate
    If valid = True Then
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields("name") = txtname.Text
        Adodc1.Recordset.Fields("address") = txtadd.Text
        Adodc1.Recordset.Fields("telno") = txttelno.Text
        Adodc1.Recordset.Fields("email") = txtemail.Text
        Adodc1.Recordset.Fields("username") = txtuname.Text
        Adodc1.Recordset.Fields("password") = txtpwd.Text
        Adodc1.Recordset.Fields("favclub") = txtclub.Text
        Adodc1.Recordset.Update
        FrmCountry.Show
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    MsgBox "Welcome to Registration!", , "Welcome"
    
End Sub

Private Sub txtname_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtpwd_GotFocus()
    SendKeys "{Home}+{End}"
    If txtpwd.Text = "(max 10 characters)" Then
        txtpwd.Text = ""
        txtpwd.PasswordChar = "*"
    End If
End Sub

Private Sub txtuname_GotFocus()
    SendKeys "{Home}+{End}"
    If txtuname.Text = "(max 10 characters)" Then
        txtuname.Text = ""
    End If
End Sub
