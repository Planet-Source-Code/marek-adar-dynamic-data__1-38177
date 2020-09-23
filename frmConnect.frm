VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Connect"
   ClientHeight    =   2760
   ClientLeft      =   5160
   ClientTop       =   2835
   ClientWidth     =   3720
   ControlBox      =   0   'False
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1065
      TabIndex        =   4
      Top             =   2250
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2355
      TabIndex        =   5
      Top             =   2250
      Width           =   1230
   End
   Begin VB.ComboBox cboTable 
      Height          =   315
      Left            =   195
      TabIndex        =   3
      Top             =   1785
      Width           =   3390
   End
   Begin VB.ComboBox cboDatabase 
      Height          =   315
      Left            =   210
      TabIndex        =   2
      Top             =   1110
      Width           =   3390
   End
   Begin VB.TextBox txtServer 
      Height          =   330
      Left            =   210
      TabIndex        =   1
      Top             =   435
      Width           =   3360
   End
   Begin VB.Label lblTable 
      Caption         =   "Table"
      Height          =   270
      Left            =   195
      TabIndex        =   0
      Top             =   1515
      Width           =   1905
   End
   Begin VB.Label lblDatabase 
      Caption         =   "Database"
      Height          =   270
      Left            =   210
      TabIndex        =   7
      Top             =   840
      Width           =   1905
   End
   Begin VB.Label lblServer 
      Caption         =   "Server"
      Height          =   270
      Left            =   210
      TabIndex        =   6
      Top             =   180
      Width           =   1905
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDatabase_Click()
    On Error GoTo Err
    If Trim(Me.txtServer.Text) = "" Then
        MsgBox "Please insert a server !", vbInformation
        Exit Sub
    End If
    SubCreateConnection Me.txtServer.Text, Me.cboDatabase.Text
    SubCreateRecordset RSMetaData, "Select table_name from information_schema.tables where table_type='BASE TABLE' order by table_name"
    SubFillCombos RSMetaData, Me.cboTable
    Exit Sub
Err:
    SubShowError Err.Number, Err.Description
End Sub

Private Sub cboDatabase_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboTable_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    If CN.State = adStateOpen Then CN.Close
    Set CN = Nothing
    Unload Me
    End
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    If CN.State = adStateClosed Then
        MsgBox "Not connected !", vbInformation
        Exit Sub
    End If
    strTable = Me.cboTable.Text
    strDatabase = Me.cboDatabase.Text
    strServer = Me.txtServer.Text
    Unload Me
End Sub

Private Sub Form_Load()
    SubCenterForm Me
End Sub

Private Sub txtServer_Change()
    Me.cboDatabase.Clear
    Me.cboTable.Clear
End Sub


Private Sub txtServer_LostFocus()
    On Error GoTo Err
    If Trim(Me.txtServer.Text) = "" Then
        MsgBox "Please insert a server !", vbInformation
        Exit Sub
    End If
    SubCreateConnection Me.txtServer.Text
    SubCreateRecordset RSMetaData, "Select Name from sysdatabases order by name"
    SubFillCombos RSMetaData, Me.cboDatabase
    Exit Sub
Err:
    SubShowError Err.Number, Err.Description
End Sub
