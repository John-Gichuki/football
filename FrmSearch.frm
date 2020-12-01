VERSION 5.00
Begin VB.Form FrmSearch 
   Caption         =   "Player Search"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdname 
      Caption         =   "Search by &Name"
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
      Left            =   540
      TabIndex        =   1
      ToolTipText     =   "CLICK TO VIEW CLUB INFO"
      Top             =   660
      Width           =   2820
   End
   Begin VB.CommandButton cmdcountry 
      Caption         =   "Search by &Country"
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
      Left            =   540
      TabIndex        =   2
      ToolTipText     =   "CLICK TO VIEW REGISTERED USERS"
      Top             =   1170
      Width           =   2820
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
      Height          =   435
      Left            =   2475
      TabIndex        =   3
      Top             =   1935
      Width           =   1020
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
      Left            =   180
      TabIndex        =   0
      Top             =   195
      Width           =   1665
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
    FrmGuest.Show
    Unload Me
End Sub

Private Sub cmdcountry_Click()
    FrmSearchCountry.Show
    Unload Me
End Sub

Private Sub cmdname_Click()
    FrmSearchPlayer.Show
    Unload Me
End Sub

