VERSION 5.00
Begin VB.Form frmBrowse 
   Caption         =   "Picture Browser"
   ClientHeight    =   4410
   ClientLeft      =   6090
   ClientTop       =   3735
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4410
   ScaleWidth      =   6135
   Begin VB.CommandButton cmdaddpicture 
      Caption         =   "Add &Picture"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   2535
   End
   Begin VB.FileListBox FilList 
      Height          =   3405
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.DriveListBox drvList 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaddpicture_Click()
    If FilList = "" Then
        MsgBox "Please select a picture.", vbExclamation, "Error"
    Else
        FrmTeamInfo.txtthumb.Text = FilList.Path & "\" & FilList
        FrmTeamInfo.txtthumbdata.Text = FilList
        Unload Me
    End If

End Sub

Private Sub Dir1_Change()
    FilList.Path = Dir1.Path
End Sub

Private Sub drvList_Change()
    On Error Resume Next
    Dir1.Path = drvList.Drive
    drvList.Drive = Dir1.Path
End Sub

Private Sub Form_Load()
    FilList.Pattern = "*.*"
    drvList.Drive = "g:\"
    Select Case img
        Case "stadium":
            Dir1.Path = App.Path & "\Images\stadiums"
        Case "home"
            Dir1.Path = App.Path & "\Images\kits"
        Case "away"
            Dir1.Path = App.Path & "\Images\kits"
        Case "logo"
            Dir1.Path = App.Path & "\Images\logos"
    End Select
End Sub

