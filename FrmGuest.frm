VERSION 5.00
Begin VB.Form FrmGuest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdlogout 
      Caption         =   "&Logout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2415
      TabIndex        =   3
      Top             =   1860
      Width           =   1020
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "&Player Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "CLICK TO VIEW REGISTERED USERS"
      Top             =   1095
      Width           =   2820
   End
   Begin VB.CommandButton cmdclub 
      Caption         =   "View &Club Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "CLICK TO VIEW CLUB INFO"
      Top             =   585
      Width           =   2820
   End
   Begin VB.Label Label1 
      Caption         =   "Select  a  task"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "FrmGuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclub_Click()
    FrmCountry.Show
    Unload Me
End Sub

Private Sub cmdlogout_Click()
    FrmLogin.Show
    Unload Me
End Sub

Private Sub cmdsearch_Click()
    FrmSearch.Show
    Unload Me
End Sub
