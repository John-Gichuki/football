VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmUserInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Information Window"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcan 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10335
      TabIndex        =   7
      Top             =   4800
      Width           =   1605
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5355
      TabIndex        =   8
      ToolTipText     =   "GO BACK TO SELECTION PAGE"
      Top             =   5475
      Width           =   1605
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save Changes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7845
      TabIndex        =   6
      ToolTipText     =   "CLICK TO SAVE CHANGES"
      Top             =   4800
      Width           =   1605
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "&Delete User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5355
      TabIndex        =   5
      ToolTipText     =   "CLICK TO DELETE USER"
      Top             =   4800
      Width           =   1605
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2865
      TabIndex        =   4
      ToolTipText     =   "CLICK TO EDIT USER INFO"
      Top             =   4800
      Width           =   1605
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Add User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   375
      TabIndex        =   3
      ToolTipText     =   "CLICK TO ADD USER"
      Top             =   4800
      Width           =   1605
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmUserInfo.frx":0000
      Height          =   3795
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   6694
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "User Information"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "NAME"
         Caption         =   "NAME"
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
         DataField       =   "ADDRESS"
         Caption         =   "ADDRESS"
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
         DataField       =   "TELNO"
         Caption         =   "TELNO"
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
         DataField       =   "EMAIL"
         Caption         =   "EMAIL"
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
         DataField       =   "USERNAME"
         Caption         =   "USERNAME"
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
         DataField       =   "PWD"
         Caption         =   "PWD"
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
         DataField       =   "FAVCLUB"
         Caption         =   "FAVCLUB"
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
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2369.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc UserInfo 
      Height          =   330
      Left            =   3135
      Top             =   5490
      Visible         =   0   'False
      Width           =   1950
      _ExtentX        =   3440
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
      Caption         =   "UserInfo"
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
   Begin VB.TextBox txtuserno 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2955
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   180
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Registered Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   2715
   End
End
Attribute VB_Name = "FrmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim userno As Long
Dim delete As Boolean
Dim confirm As Integer

Private Sub cmdadd_Click()
    DataGrid1.AllowAddNew = True
    DataGrid1.AllowUpdate = True
    DataGrid1.Enabled = True
    cmdsave.Enabled = True
    cmdcan.Enabled = True
    cmdadd.Enabled = False
    cmdedit.Enabled = False
    cmddel.Enabled = False
End Sub

Private Sub cmdcan_Click()
    UserInfo.Refresh
    DataGrid1.Refresh
    DataGrid1.AllowUpdate = False
    DataGrid1.AllowAddNew = False
    DataGrid1.Enabled = False
    delete = False
    
    cmdsave.Enabled = False
    cmdcan.Enabled = False
    cmdadd.Enabled = True
    cmdedit.Enabled = True
    cmddel.Enabled = True
End Sub

Private Sub cmddel_Click()
    Dim temp As Integer
    DataGrid1.AllowUpdate = True
    DataGrid1.Enabled = True
    delete = True
    temp = MsgBox("Select the record to be deleted and then press 'Save'", vbInformation, "To Delete")
    
    cmdsave.Enabled = True
    cmdcan.Enabled = True
    cmdadd.Enabled = False
    cmdedit.Enabled = False
    cmddel.Enabled = False
End Sub

Private Sub cmdedit_Click()
    DataGrid1.AllowUpdate = True
    DataGrid1.Enabled = True
    
    cmdsave.Enabled = True
    cmdcan.Enabled = True
    cmdadd.Enabled = False
    cmdedit.Enabled = False
    cmddel.Enabled = False
End Sub

Private Sub cmdexit_Click()
    FrmAdmin.Show
    Unload Me
End Sub

Private Sub cmdsave_Click()

    If delete = True Then
        confirm = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Deletion Confirmation")
        If confirm = vbYes Then
            UserInfo.Recordset.delete
            MsgBox "Record Deleted!", , "Message"
        Else
            MsgBox "Record Not Deleted!", , "Message"
        End If
    End If
    delete = False
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowUpdate = False
    DataGrid1.Refresh
    
    UserInfo.Recordset.MoveFirst
    userno = 0
    While UserInfo.Recordset.EOF = False
        userno = userno + 1
        UserInfo.Recordset.MoveNext
    Wend
    txtuserno.Text = userno
    DataGrid1.Enabled = False
    
    cmdsave.Enabled = False
    cmdadd.Enabled = True
    cmdedit.Enabled = True
    cmddel.Enabled = True
    cmdcan.Enabled = False

End Sub

Private Sub Form_Load()
    userno = 0
    UserInfo.Recordset.MoveFirst
    While UserInfo.Recordset.EOF = False
        userno = userno + 1
        UserInfo.Recordset.MoveNext
    Wend
    txtuserno.Text = userno
    cmdsave.Enabled = False
    cmdcan.Enabled = False
End Sub
