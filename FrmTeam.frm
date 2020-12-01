VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmTeam 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   12675
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc TeamInfo 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=football.mdb;Mode=Read;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=football.mdb;Mode=Read;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "admin"
      Password        =   ""
      RecordSource    =   "select * from players"
      Caption         =   "TeamInfo"
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
   Begin VB.Frame Frame2 
      Caption         =   "Club Information"
      Height          =   1680
      Left            =   210
      TabIndex        =   8
      Top             =   1065
      Width           =   10455
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "FrmTeam.frx":0000
         Height          =   885
         Left            =   90
         TabIndex        =   9
         Top             =   300
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   1561
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "MANAGER"
            Caption         =   "MANAGER"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "STADIUM"
            Caption         =   "STADIUM"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "CAPACITY"
            Caption         =   "CAPACITY"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "FOUNDED"
            Caption         =   "FOUNDED"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "HONOURS"
            Caption         =   "HONOURS"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1934.929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2025.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3734.929
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc clubinfo 
      Height          =   330
      Left            =   0
      Top             =   330
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
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
      RecordSource    =   "select * from clubs"
      Caption         =   "ClubInfo"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5700
      Left            =   210
      TabIndex        =   0
      Top             =   2910
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   10054
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Team Profile"
      TabPicture(0)   =   "FrmTeam.frx":0017
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Picture Gallery"
      TabPicture(1)   =   "FrmTeam.frx":0033
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   5175
         Left            =   -74925
         TabIndex        =   4
         Top             =   450
         Width           =   11640
         Begin VB.Image ImageStad 
            BorderStyle     =   1  'Fixed Single
            Height          =   4350
            Left            =   2595
            Stretch         =   -1  'True
            Top             =   210
            Width           =   8790
         End
         Begin VB.Image ImageHome 
            BorderStyle     =   1  'Fixed Single
            Height          =   1560
            Left            =   495
            Stretch         =   -1  'True
            Top             =   465
            Width           =   1560
         End
         Begin VB.Label lblhome 
            Alignment       =   2  'Center
            Caption         =   "Home Kit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   675
            TabIndex        =   7
            Top             =   2100
            Width           =   1125
         End
         Begin VB.Image ImageAway 
            BorderStyle     =   1  'Fixed Single
            Height          =   1560
            Left            =   495
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   1560
         End
         Begin VB.Label lblaway 
            Alignment       =   2  'Center
            Caption         =   "Away Kit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   780
            TabIndex        =   6
            Top             =   4320
            Width           =   990
         End
         Begin VB.Label lblstadname 
            Alignment       =   2  'Center
            Caption         =   "Home Ground"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   5
            Top             =   4635
            Width           =   8745
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5235
         Left            =   90
         TabIndex        =   1
         Top             =   345
         Width           =   12105
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "FrmTeam.frx":004F
            Height          =   4560
            Left            =   150
            Negotiate       =   -1  'True
            TabIndex        =   2
            Top             =   285
            Width           =   11790
            _ExtentX        =   20796
            _ExtentY        =   8043
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            Enabled         =   0   'False
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   17
            WrapCellPointer =   -1  'True
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Profiles"
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "name"
               Caption         =   "name"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "club"
               Caption         =   "club"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "position"
               Caption         =   "position"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "dob"
               Caption         =   "dob"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "country"
               Caption         =   "country"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "doj"
               Caption         =   "doj"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "tfrom"
               Caption         =   "tfrom"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1920.189
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2204.788
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   675.213
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1950.236
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   2174.74
               EndProperty
            EndProperty
         End
      End
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
      Height          =   450
      Left            =   10635
      TabIndex        =   3
      Top             =   8865
      Width           =   1245
   End
   Begin VB.Label lblteam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Team Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   0
      TabIndex        =   10
      Top             =   180
      Width           =   11910
   End
   Begin VB.Image ImgCntLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1560
      Left            =   10815
      Stretch         =   -1  'True
      Top             =   1140
      Width           =   1560
   End
End
Attribute VB_Name = "FrmTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim delete As Boolean
Dim confirm As Integer

Private Sub cmdback_Click()
    FrmClub.Show
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim clubsql As String
    Dim playersql As String
    
    clubsql = "select * from clubs where club like '" & UCase(team) & "'; "
    playersql = "select * from players where club like '" & UCase(team) & "';"
    
    TeamInfo.RecordSource = playersql
    TeamInfo.Refresh
    
    clubinfo.RecordSource = clubsql
    clubinfo.Refresh
    
    SSTab1.Tab = 0
    
    lblstadname.Caption = clubinfo.Recordset.Fields("stadium").Value
    
    txtthumb = clubinfo.Recordset.Fields("imgstad").Value
    ImageStad.Picture = LoadPicture(App.Path & txtthumb)
    
    txtthumb = clubinfo.Recordset.Fields("imghome").Value
    ImageHome.Picture = LoadPicture(App.Path & txtthumb)
    
    txtthumb = clubinfo.Recordset.Fields("imgaway").Value
    If txtthumb Is Not Empty Then
        ImageAway.Picture = LoadPicture(App.Path & txtthumb)
    End If
    
    txtthumb = clubinfo.Recordset.Fields("logo").Value
    ImgCntLogo.Picture = LoadPicture(App.Path & txtthumb)
    
    Select Case club
    Case 1:
        If league = 1 Then
            FrmTeam.Caption = "Arsenal Football Club"
            lblteam.Caption = "ARSENAL FOOTBALL CLUB"
        End If
            
        If league = 2 Then
            FrmTeam.Caption = "AC Milan"
            lblteam.Caption = "AC MILAN"
        End If
        
        If league = 3 Then
            FrmTeam.Caption = "AS Monaco"
            lblteam.Caption = "AS MONACO"
        End If
        
        If league = 4 Then
            FrmTeam.Caption = "Football Club de Barcelona"
            lblteam.Caption = "BARCELONA FOOTBALL CLUB"
        End If
        
        If league = 5 Then
            FrmTeam.Caption = "Bayer Leverkusen"
            lblteam.Caption = "BAYER LEVERKUSEN"
            lblaway.Visible = False
            ImageAway.Visible = False
        End If
        
        If league = 6 Then
            FrmTeam.Caption = "Ajax Amsterdam"
            lblaway.Visible = False
            lblteam.Caption = "AJAX AMSTERDAM"
            ImageAway.Visible = False
        End If
        
    Case 2:
        If league = 1 Then
            FrmTeam.Caption = "Chelsea Football Club"
            lblteam.Caption = "CHELSEA FOOTBALL CLUB"
        End If
            
        If league = 2 Then
            FrmTeam.Caption = "AS Roma"
            lblteam.Caption = "AS ROMA"
        End If
        
        If league = 4 Then
            FrmTeam.Caption = "Real Madrid"
            lblteam.Caption = "REAL MADRID"
        End If
        
        If league = 5 Then
            FrmTeam.Caption = "Bayern Munchen"
            lblteam.Caption = "BAYERN MUNCHEN"
        End If
        
    Case 3:
        If league = 1 Then
            FrmTeam.Caption = "Liverpool Football Club"
            lblteam.Caption = "LIVERPOOL FOOTBALL CLUB"
        End If
            
        If league = 2 Then
            FrmTeam.Caption = "Internazionale Football Club"
            lblteam.Caption = "INTERNAZIONALE FOOTBALL CLUB"
        End If
        
        If league = 3 Then
            FrmTeam.Caption = "Olympique Lyon"
            lblteam.Caption = "OLYMPIQUE LYON"
        End If
        
        If league = 4 Then
            FrmTeam.Caption = "Valencia Club de Football"
            lblteam.Caption = "VALENCIA CLUB DE FOOTBALL"
        End If
        
        If league = 5 Then
            FrmTeam.Caption = "SV Werder Bremen"
            lblteam.Caption = "SV WERDER BREMEN"
        End If
        
        If league = 6 Then
            FrmTeam.Caption = "PSV Eindhoven"
            lblteam.Caption = "PSV EINDHOVEN"
        End If
        
    Case 4:
        If league = 1 Then
            FrmTeam.Caption = "Manchester United Football Club"
            lblteam.Caption = "MANCHESTER UNITED FOOTBALL CLUB"
        End If
            
        If league = 2 Then
            FrmTeam.Caption = "Juventus Football Club"
            lblteam.Caption = "JUVENTUS FOOTBALL CLUB"
        End If
        
        If league = 4 Then
            FrmTeam.Caption = "Villareal Club de Football"
            lblteam.Caption = "VILLAREAL CLUB DE FOOTBALL"
            lblaway.Visible = False
            ImageAway.Visible = False
        End If
        
        If league = 5 Then
            FrmTeam.Caption = "VFB Stuttgart"
            lblteam.Caption = "VFB STUTTGART"
        End If
    End Select
End Sub
