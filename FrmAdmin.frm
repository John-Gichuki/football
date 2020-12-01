VERSION 5.00
Begin VB.Form FrmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome, Mr. Vineet Potdar"
   ClientHeight    =   2445
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdlogout 
      Caption         =   "&Logout"
      Height          =   435
      Left            =   2625
      TabIndex        =   3
      Top             =   1890
      Width           =   1020
   End
   Begin VB.CommandButton cmdclub 
      Caption         =   "View &Club Information"
      Height          =   510
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "CLICK TO VIEW CLUB INFO"
      Top             =   600
      Width           =   2820
   End
   Begin VB.CommandButton cmduser 
      Caption         =   "View &User Information"
      Height          =   510
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "CLICK TO VIEW REGISTERED USERS"
      Top             =   1110
      Width           =   2820
   End
   Begin VB.Label Label1 
      Caption         =   "Select a Choice"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3105
   End
End
Attribute VB_Name = "FrmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdclub_Click()
    admin = True
    FrmCountry.Show
    Unload FrmLogin
    Unload Me
End Sub

Private Sub cmdlogout_Click()
    FrmLogin.Show
    admin = False
    Unload Me
End Sub

Private Sub cmduser_Click()
    FrmUserInfo.Show
    Unload FrmLogin
    Unload Me
End Sub

