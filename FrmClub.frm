VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmClub 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome to English Premier League"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7965
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   5715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   210
      ImageHeight     =   240
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClub.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClub.frx":5CCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClub.frx":8622
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClub.frx":A441
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pctbox2 
      Height          =   1560
      Left            =   270
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   8
      Top             =   3375
      Width           =   1560
   End
   Begin VB.PictureBox pctbox1 
      Height          =   1560
      Left            =   270
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   7
      Top             =   1065
      Width           =   1560
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
      Height          =   510
      Left            =   2205
      TabIndex        =   5
      Top             =   5730
      Width           =   1815
   End
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
      Height          =   510
      Left            =   4020
      TabIndex        =   6
      Top             =   5730
      Width           =   1815
   End
   Begin VB.CommandButton cmdclub4 
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
      Left            =   5940
      TabIndex        =   4
      Top             =   4950
      Width           =   1965
   End
   Begin VB.CommandButton cmdclub3 
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
      Left            =   5940
      TabIndex        =   3
      Top             =   2640
      Width           =   1965
   End
   Begin VB.CommandButton cmdclub2 
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
      Left            =   135
      TabIndex        =   2
      Top             =   4950
      Width           =   1965
   End
   Begin VB.CommandButton cmdclub1 
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
      Left            =   75
      TabIndex        =   1
      Top             =   2640
      Width           =   1965
   End
   Begin VB.PictureBox pctbox4 
      Height          =   1560
      Left            =   6120
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   10
      Top             =   3375
      Width           =   1560
   End
   Begin VB.PictureBox pctbox3 
      Height          =   1560
      Left            =   6120
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   9
      Top             =   1065
      Width           =   1560
   End
   Begin VB.Label lblna 
      BackStyle       =   0  'Transparent
      Caption         =   "LEAGUE LOGO NOT AVAILABLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3090
      TabIndex        =   11
      Top             =   1530
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Image ImageLeague 
      Height          =   3195
      Left            =   2190
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   3645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SELECT A CLUB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2370
      TabIndex        =   0
      Top             =   240
      Width           =   3225
   End
End
Attribute VB_Name = "FrmClub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsclub As New ADODB.Recordset
Dim sqlstr As String

Private Sub cmdback_Click()
    FrmCountry.Show
    Unload Me
End Sub

Private Sub cmdclub1_Click()
    team = cmdclub1.Caption
    club = 1
    If admin = False Then
        FrmTeam.Show
    Else
        FrmTeamInfo.Show
    End If
    Unload Me
End Sub

Private Sub cmdclub2_Click()
    If league = 3 Or league = 6 Then
        MsgBox "There are only two clubs in this league", , "Error"
        Exit Sub
    End If
    team = cmdclub2.Caption
    club = 2
    If admin = False Then
        FrmTeam.Show
    Else
        FrmTeamInfo.Show
    End If
    Unload Me
End Sub

Private Sub cmdclub3_Click()
    team = cmdclub3.Caption
    club = 3
    If admin = False Then
        FrmTeam.Show
    Else
        FrmTeamInfo.Show
    End If
    Unload Me
End Sub

Private Sub cmdclub4_Click()
    If league = 3 Or league = 6 Then
        MsgBox "There are only two clubs in this league", , "Error"
        Exit Sub
    End If
    team = cmdclub4.Caption
    club = 4
    If admin = False Then
        FrmTeam.Show
    Else
        FrmTeamInfo.Show
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    
    Select Case league
        Case 1:
            FrmClub.Caption = "ENGLISH PREMIER LEAGUE"
            ImageLeague.Picture = ImageList1.ListImages(1).Picture
            sqlstr = "select * from clubs where league like 'epl';"
            
        Case 2:
            FrmClub.Caption = "ITALIAN SERIE A"
            ImageLeague.Picture = ImageList1.ListImages(3).Picture
            sqlstr = "select * from clubs where league like 'seriea';"
            
        Case 3:
            lblna.Visible = True
            FrmClub.Caption = "FRENCH LEAGUE"
            sqlstr = "select * from clubs where league like 'french';"
            cmdclub2.Visible = False
            pctbox2.Visible = False
            cmdclub4.Visible = False
            pctbox4.Visible = False

        Case 4:
            FrmClub.Caption = "SPANISH PRIMERA LIGA"
            ImageLeague.Picture = ImageList1.ListImages(2).Picture
            sqlstr = "select * from clubs where league like 'laliga';"
            
        Case 5:
            FrmClub.Caption = "GERMAN BUNDESLIGA"
            ImageLeague.Picture = ImageList1.ListImages(4).Picture
            sqlstr = "select * from clubs where league like 'bundesliga';"
            
        Case 6:
            lblna.Visible = True
            FrmClub.Caption = "DUTCH LEAGUE"
            sqlstr = "select * from clubs where league like 'dutch';"
            cmdclub2.Visible = False
            pctbox2.Visible = False
            cmdclub4.Visible = False
            pctbox4.Visible = False

    End Select
    
    If rsclub.State = adStateOpen Then
        rsclub.Close
    End If
    
    rsclub.Open sqlstr, dbcon, adOpenStatic, adLockOptimistic
    rsclub.Sort = "club"
    
    cmdclub1.Caption = UCase(rsclub.Fields("club").Value)
    pctbox1.Picture = LoadPicture(App.Path & rsclub.Fields("logo").Value)
    rsclub.MoveNext
    
    cmdclub3.Caption = UCase(rsclub.Fields("club").Value)
    pctbox3.Picture = LoadPicture(App.Path & rsclub.Fields("logo").Value)
    rsclub.MoveNext
    
    If league <> 3 And league <> 6 Then
        cmdclub2.Caption = rsclub.Fields("club").Value
        pctbox2.Picture = LoadPicture(App.Path & rsclub.Fields("logo").Value)
        rsclub.MoveNext
        
        cmdclub4.Caption = UCase(rsclub.Fields("club").Value)
        pctbox4.Picture = LoadPicture(App.Path & rsclub.Fields("logo").Value)
        rsclub.MoveNext
    End If
    rsclub.Close
    cmdclub1.ToolTipText = "Go to " & cmdclub1.Caption & " page"
    cmdclub2.ToolTipText = "Go to " & cmdclub2.Caption & " page"
    cmdclub3.ToolTipText = "Go to " & cmdclub3.Caption & " page"
    cmdclub4.ToolTipText = "Go to " & cmdclub4.Caption & " page"
End Sub

Private Sub cmdlogout_Click()
    MsgBox "Thank you for visiting!!! Please Come Again!!!", , "Logged out"
    FrmLogin.Show
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsclub.State = adStateOpen Then
        rsclub.Close
    End If
End Sub
