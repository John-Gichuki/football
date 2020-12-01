VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcom to CHAMPIONS LEAGUE FOOTBALL DATABASE"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11415
      TabIndex        =   3
      Top             =   9345
      Width           =   1110
   End
   Begin VB.CommandButton cmdguest 
      Caption         =   "&Guest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click here to Login as Guest"
      Top             =   7050
      Width           =   2880
   End
   Begin VB.CommandButton cmdadmin 
      Caption         =   "&Administrator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   885
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click here to Login as Admin"
      Top             =   7050
      Width           =   2880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Champions League Football Database"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1575
      Left            =   1020
      TabIndex        =   0
      Top             =   -45
      Width           =   10545
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   9930
      Left            =   -30
      Picture         =   "FrmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   -15
      Width           =   12660
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadmin_Click()
    FrmLoginAdmin.Show
End Sub

Private Sub cmdexit_Click()
    Unload Me
    Unload FrmAdmin
    Unload FrmClub
    Unload FrmCountry
    Unload FrmLogin
    Unload FrmLoginAdmin
    Unload FrmLoginGuest
    Unload FrmRegister
    Unload frmSplash
    Unload FrmTeam
    Unload FrmTeamInfo
    Unload FrmUserInfo
    Unload frmBrowse
    dbcon.Close
End Sub

Private Sub cmdguest_Click()
    FrmLoginGuest.Show
End Sub


