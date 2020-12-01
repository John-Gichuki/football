VERSION 5.00
Begin VB.Form FrmCountry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Country Selection"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9690
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdfra 
      Caption         =   "&France"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   345
      TabIndex        =   3
      ToolTipText     =   "GO TO FRENCH LEAGUE PAGE"
      Top             =   5625
      Width           =   1530
   End
   Begin VB.CommandButton cmdita 
      Caption         =   "&Italy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   345
      TabIndex        =   2
      ToolTipText     =   "GO TO ITALIAN SERIE A PAGE"
      Top             =   3870
      Width           =   1530
   End
   Begin VB.CommandButton cmdeng 
      Caption         =   "&England"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   345
      TabIndex        =   1
      ToolTipText     =   "GO TO ENGLISH PREMIER LEAGUE PAGE"
      Top             =   2115
      Width           =   1530
   End
   Begin VB.PictureBox Picture7 
      Height          =   915
      Left            =   7785
      Picture         =   "FrmCountry.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   1215
      TabIndex        =   13
      Top             =   4710
      Width           =   1275
   End
   Begin VB.CommandButton cmdhol 
      Caption         =   "&Holland"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7665
      TabIndex        =   6
      ToolTipText     =   "GO TO DUTCH LEAGUE PAGE"
      Top             =   5625
      Width           =   1530
   End
   Begin VB.PictureBox Picture6 
      Height          =   810
      Left            =   7785
      Picture         =   "FrmCountry.frx":04FA
      ScaleHeight     =   750
      ScaleWidth      =   1215
      TabIndex        =   12
      Top             =   3045
      Width           =   1275
   End
   Begin VB.CommandButton cmdger 
      Caption         =   "&Germany"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7665
      TabIndex        =   5
      ToolTipText     =   "GO TO GERMAN BUNDESLIGA PAGE"
      Top             =   3870
      Width           =   1530
   End
   Begin VB.PictureBox Picture5 
      Height          =   915
      Left            =   7785
      Picture         =   "FrmCountry.frx":0942
      ScaleHeight     =   855
      ScaleWidth      =   1215
      TabIndex        =   11
      Top             =   1200
      Width           =   1275
   End
   Begin VB.CommandButton cmdesp 
      Caption         =   "&Spain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7665
      TabIndex        =   4
      ToolTipText     =   "GO TO SPANISH PRIMERA LIGA PAGE"
      Top             =   2115
      Width           =   1530
   End
   Begin VB.PictureBox Picture4 
      Height          =   915
      Left            =   510
      Picture         =   "FrmCountry.frx":0EEB
      ScaleHeight     =   855
      ScaleWidth      =   1215
      TabIndex        =   10
      Top             =   4710
      Width           =   1275
   End
   Begin VB.PictureBox Picture3 
      Height          =   825
      Left            =   495
      Picture         =   "FrmCountry.frx":1320
      ScaleHeight     =   765
      ScaleWidth      =   1215
      TabIndex        =   9
      Top             =   3030
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      Height          =   915
      Left            =   495
      Picture         =   "FrmCountry.frx":1784
      ScaleHeight     =   855
      ScaleWidth      =   1215
      TabIndex        =   8
      Top             =   1200
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Height          =   4665
      Left            =   2550
      Picture         =   "FrmCountry.frx":2162
      ScaleHeight     =   4605
      ScaleWidth      =   4560
      TabIndex        =   14
      Top             =   1320
      Width           =   4620
   End
   Begin VB.CommandButton cmdlogout 
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
      Left            =   4215
      TabIndex        =   7
      Top             =   6240
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select a Country"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2985
      TabIndex        =   0
      Top             =   375
      Width           =   3570
   End
End
Attribute VB_Name = "FrmCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdeng_Click()
    league = 1
    FrmClub.Show
    Unload Me
End Sub

Private Sub cmdesp_Click()
    league = 4
    FrmClub.Show
    Unload Me
End Sub

Private Sub cmdfra_Click()
    league = 3
    FrmClub.Show
    Unload Me
End Sub

Private Sub cmdger_Click()
    league = 5
    FrmClub.Show
    Unload Me
End Sub

Private Sub cmdhol_Click()
    league = 6
    FrmClub.Show
    Unload Me
End Sub

Private Sub cmdita_Click()
    league = 2
    FrmClub.Show
    Unload Me
End Sub

Private Sub cmdlogout_Click()
    If admin = False Then
        FrmGuest.Show
        Unload Me
    Else
        FrmAdmin.Show
        Unload Me
    End If
End Sub
