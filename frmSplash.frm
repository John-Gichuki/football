VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4800
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4800
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   195
      Top             =   4200
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub form_click()
    FrmLogin.Show
    admin = False
    Unload Me
End Sub

Private Sub Form_Load()
    
    Set dbcon = New ADODB.Connection
    dbcon.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\football.mdb"
       
    dbcon.Mode = adModeReadWrite
    dbcon.CursorLocation = adUseClient
    dbcon.Open

End Sub

Private Sub Timer1_Timer()
    form_click
End Sub


