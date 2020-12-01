VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form FrmSearchPlayer 
   Caption         =   "Player Search"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Search 
      Height          =   390
      Left            =   7740
      Top             =   630
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   688
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
      RecordSource    =   "select * from players"
      Caption         =   "Search"
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
      Height          =   525
      Left            =   4875
      TabIndex        =   5
      Top             =   960
      Width           =   2085
   End
   Begin VB.CommandButton cmdnewsearch 
      Caption         =   "&New Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2610
      TabIndex        =   4
      Top             =   960
      Width           =   2085
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   345
      TabIndex        =   3
      Top             =   960
      Width           =   2085
   End
   Begin VB.TextBox txtsearchname 
      Height          =   390
      Left            =   2370
      TabIndex        =   1
      Top             =   180
      Width           =   3360
   End
   Begin VB.Frame Frame1 
      Caption         =   "Player Details"
      Height          =   4230
      Left            =   240
      TabIndex        =   6
      Top             =   1755
      Width           =   9480
      Begin VB.TextBox txtclub 
         BackColor       =   &H8000000F&
         DataSource      =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1635
         Width           =   2370
      End
      Begin VB.TextBox txtcountry 
         BackColor       =   &H8000000F&
         DataSource      =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2505
         Width           =   2370
      End
      Begin VB.TextBox txtdoj 
         BackColor       =   &H8000000F&
         DataSource      =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3375
         Width           =   2370
      End
      Begin VB.TextBox txtpos 
         BackColor       =   &H8000000F&
         DataSource      =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1635
         Width           =   2370
      End
      Begin VB.TextBox txtdob 
         BackColor       =   &H8000000F&
         DataSource      =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2505
         Width           =   2370
      End
      Begin VB.TextBox txttfrom 
         BackColor       =   &H8000000F&
         DataSource      =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3375
         Width           =   2370
      End
      Begin VB.Image imageplayer 
         Height          =   1725
         Left            =   6735
         Stretch         =   -1  'True
         Top             =   2265
         Width           =   1725
      End
      Begin VB.Label lblclub 
         Caption         =   "Club"
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
         Left            =   270
         TabIndex        =   18
         Top             =   1335
         Width           =   1215
      End
      Begin VB.Label lblcountry 
         Caption         =   "Country"
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
         Left            =   270
         TabIndex        =   17
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label lbldoj 
         Caption         =   "Date of Join"
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
         Left            =   270
         TabIndex        =   16
         Top             =   3075
         Width           =   1215
      End
      Begin VB.Label lblname 
         Caption         =   "Player name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   270
         TabIndex        =   13
         Top             =   480
         Width           =   2700
      End
      Begin VB.Label lblposition 
         Caption         =   "Position"
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
         Left            =   3090
         TabIndex        =   12
         Top             =   1335
         Width           =   1215
      End
      Begin VB.Label lbldob 
         Caption         =   "Date of Birth"
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
         Left            =   3090
         TabIndex        =   11
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label lbltfrom 
         Caption         =   "Transferred from"
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
         Left            =   3090
         TabIndex        =   10
         Top             =   3075
         Width           =   1560
      End
      Begin VB.Image imageclub 
         Height          =   1725
         Left            =   6735
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1725
      End
   End
   Begin VB.Label Label2 
      Caption         =   "eg. Wayne Rooney"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   2
      Top             =   255
      Width           =   2190
   End
   Begin VB.Label Label1 
      Caption         =   "Player Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "FrmSearchPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim team As String

Private Sub cmdback_Click()
    FrmSearch.Show
    Unload Me
End Sub

Private Sub cmdnewsearch_Click()
    txtsearchname.Text = ""
    lblname.Caption = ""
    txtclub.Text = ""
    txtpos.Text = ""
    txtdob.Text = ""
    txtcountry.Text = ""
    txtdoj.Text = ""
    txttfrom.Text = ""
    imageclub.Picture = LoadPicture("")
    imageplayer.Picture = LoadPicture("")
End Sub

Private Sub cmdsearch_Click()
    Dim sqlstring As String
    Dim x As Integer
    
    If txtsearchname.Text = "" Then
        Exit Sub
    Else
        sqlstring = "select name, club, position, dob, country, doj, tfrom from players where lower(name)='" & LCase(txtsearchname.Text) & "'"
        Search.RecordSource = sqlstring
        Search.Refresh
        
        Frame1.Visible = True
        If Search.Recordset.RecordCount = 0 Then
            x = MsgBox("Player not found", vbExclamation, "Not Found")
            Exit Sub
        End If
        
        lblname.Caption = Search.Recordset.Fields(0).Value
        txtclub.Text = Search.Recordset.Fields(1).Value
        txtpos.Text = Search.Recordset.Fields(2).Value
        txtdob.Text = Search.Recordset.Fields(3).Value
        txtcountry.Text = Search.Recordset.Fields(4).Value
        txtdoj.Text = Search.Recordset.Fields(5).Value
        txttfrom.Text = Search.Recordset.Fields(6).Value
        imageplayer.Picture = LoadPicture(App.Path & "/images/players/" & LCase(txtsearchname.Text) & ".jpg")
        imageclub.Picture = LoadPicture(App.Path & "/images/logos/" & LCase(txtclub.Text) & ".gif")
    End If
End Sub
