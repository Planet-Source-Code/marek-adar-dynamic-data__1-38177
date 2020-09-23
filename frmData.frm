VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmData 
   Caption         =   "Dynamic Data"
   ClientHeight    =   3900
   ClientLeft      =   3810
   ClientTop       =   2895
   ClientWidth     =   7500
   Icon            =   "frmData.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   7500
   Begin MSDataGridLib.DataGrid Data 
      Height          =   2145
      Left            =   60
      TabIndex        =   16
      Top             =   375
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   3784
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkData 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "Caption"
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1140
      Visible         =   0   'False
      Width           =   7110
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   0
      Left            =   165
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   750
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.Frame frEditData 
      Caption         =   "Edit"
      Height          =   1260
      Left            =   15
      TabIndex        =   0
      Top             =   2580
      Visible         =   0   'False
      Width           =   7395
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         Height          =   255
         Left            =   1395
         TabIndex        =   12
         Tag             =   "L"
         Top             =   270
         Width           =   390
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         Height          =   255
         Left            =   165
         TabIndex        =   11
         Tag             =   "L"
         Top             =   270
         Width           =   390
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   255
         Left            =   975
         TabIndex        =   10
         Tag             =   "L"
         Top             =   270
         Width           =   390
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   255
         Left            =   570
         TabIndex        =   9
         Tag             =   "L"
         Top             =   270
         Width           =   390
      End
      Begin VB.CommandButton cmdReconnect 
         Caption         =   "&Reconnect"
         Height          =   435
         Left            =   4650
         TabIndex        =   7
         Tag             =   "L"
         Top             =   735
         Width           =   1275
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   435
         Left            =   2805
         TabIndex        =   3
         Tag             =   "L"
         Top             =   735
         Width           =   1290
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   435
         Left            =   5970
         TabIndex        =   2
         Tag             =   "L"
         Top             =   735
         Width           =   1275
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   435
         Left            =   165
         TabIndex        =   4
         Top             =   735
         Width           =   1290
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   1485
         TabIndex        =   5
         Top             =   735
         Width           =   1290
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   435
         Left            =   1485
         TabIndex        =   1
         Tag             =   "L"
         Top             =   735
         Width           =   1290
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   435
         Left            =   165
         TabIndex        =   6
         Tag             =   "L"
         Top             =   735
         Width           =   1290
      End
   End
   Begin TabDlg.SSTab tabData 
      Height          =   2535
      Left            =   30
      TabIndex        =   8
      Top             =   45
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "Data"
      TabPicture(0)   =   "frmData.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.Label lblData 
      Caption         =   "Caption"
      Height          =   225
      Index           =   0
      Left            =   165
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   7095
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    On Error Resume Next
    SubUnlockLockForm Me, False
    RSTableData.AddNew
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    RSTableData.CancelBatch
    SubUnlockLockForm Me, True
End Sub

Private Sub cmdClose_Click()
    Unload Me
    End
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Err
    Dim ret As Long
    ret = MsgBox("Do you want to delete this record ?", vbYesNo + vbQuestion)
    If ret = vbNo Then Exit Sub
    RSTableData.Delete
    Exit Sub
Err:
    SubShowError Err.Number, Err.Description
End Sub

Private Sub cmdEdit_Click()
    On Error Resume Next
    If RSTableData.RecordCount = 0 Then
        MsgBox "No data to edit !", vbInformation
        Exit Sub
    End If
    SubUnlockLockForm Me, False
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    RSTableData.MoveFirst
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    RSTableData.MoveLast
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    RSTableData.MoveNext
    If RSTableData.EOF = True Then RSTableData.MoveLast
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    RSTableData.MovePrevious
    If RSTableData.BOF = True Then RSTableData.MoveFirst
End Sub

Private Sub cmdReconnect_Click()
    Unload Me
    Set frmData = Nothing
    Main
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo Err
    If funcCheckInputs(Me) = True Then Exit Sub
    RSTableData.Update
    SubUnlockLockForm Me, True
    Exit Sub
Err:
    SubShowError Err.Number, Err.Description
End Sub

Private Sub Data_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
End Sub

Private Sub Form_Activate()
    SubBuildForm
    SubCenterForm Me
    SubUnlockLockForm Me, True
    SubBinddata
End Sub

Sub SubBuildForm()
    'Builds the Controls on the Form.
    'The Var i is the Index for the Control
    'The Var j is counter for controls per tab
    
    Dim i As Integer
    Dim j As Integer
    Const constControlsPerTab As Integer = 10
    Me.Caption = "Table: " & strTable & " / Server: " & strServer
    SubCreateRecordset RSMetaData, "Select * from information_schema.columns where table_name='" & strTable & "' and table_schema='dbo'"
    Set Me.Data.Container = Me.tabData
    Me.tabData.Tabs = Me.tabData.Tabs + 1
    Me.tabData.TabCaption(Me.tabData.Tabs - 1) = "Page " & CStr(Me.tabData.Tabs - 1)
    RSMetaData.MoveFirst
    Do Until RSMetaData.EOF = True
        i = i + 1
        j = j + 1
        Select Case RSMetaData.Fields("Data_type").Value
            Case "datetime", "smalldatetime":
                SubCreateTextControl "D", i, j
                Me.txtData(i).MaxLength = 10
            Case "varchar", "char", "nchar", "nvarchar"
                SubCreateTextControl "V", i, j
                Me.txtData(i).MaxLength = CInt(RSMetaData.Fields("character_maximum_length").Value)
                Me.lblData(i).Caption = Me.lblData(i).Caption & " Length:" & RSMetaData.Fields("character_maximum_length").Value
            Case "text", "ntext":
                SubCreateTextControl "T", i, j
            Case "int", "smallint", "tinyint":
                SubCreateTextControl "I", i, j
                Me.lblData(i).Caption = Me.lblData(i).Caption & " Precision:" & RSMetaData.Fields("NUMERIC_PRECISION").Value
            Case "float", "decimal", "money", "smallmoney":
                SubCreateTextControl "F", i, j
                Me.lblData(i).Caption = Me.lblData(i).Caption & " Precision:" & RSMetaData.Fields("NUMERIC_PRECISION").Value & "/" & RSMetaData.Fields("NUMERIC_SCALE").Value
            Case "binary", "varbinary", "image":
                SubCreateTextControl "B", i, j
                Me.txtData(i).Enabled = False
            Case "bit":
                SubCreateCheckControl i, j
            Case "uniqueidentifier"
                SubCreateTextControl "U", i, j
        End Select
        
        RSMetaData.MoveNext
        If j = constControlsPerTab Then
            j = 0
            Me.tabData.Tabs = Me.tabData.Tabs + 1
            Me.tabData.Tab = Me.tabData.Tabs - 1
            Me.tabData.TabCaption(Me.tabData.Tabs - 1) = "Page " & CStr(Me.tabData.Tabs - 1)
        End If
    Loop
    
    'Resizes the Controls on the form
    Me.tabData.Height = constSpaceBetweenControls * constControlsPerTab + 450
    Me.Height = constSpaceBetweenControls * constControlsPerTab + 1000 + frEditData.Height
    Me.Data.Height = Me.tabData.Height - 400
    Me.frEditData.Top = constSpaceBetweenControls * constControlsPerTab + 580
    Me.frEditData.Visible = True
    Me.tabData.Tab = 0
End Sub

Sub SubCreateTextControl(strDatatype As String, i As Integer, j As Integer)
    'Creates the Textcontrols and Lables
    Me.tabData.Tab = Me.tabData.Tabs - 1
    Load Me.txtData(i)
    Load Me.lblData(i)
    Set Me.txtData(i).Container = Me.tabData
    Set Me.lblData(i).Container = Me.tabData
    Me.txtData(i).Visible = True
    Me.lblData(i).Visible = True
    Me.txtData(i).DataField = RSMetaData.Fields("column_name").Value
    Me.lblData(i).Caption = UCase$(RSMetaData.Fields("column_name").Value & " (" & RSMetaData.Fields("Data_type").Value & ")")
    If RSMetaData.Fields("is_nullable").Value = "YES" Then
        Me.txtData(i).Tag = "Y"
    Else
        Me.txtData(i).Tag = "N"
    End If
    Me.txtData(i).Tag = Me.txtData(i).Tag & strDatatype
    Me.txtData(i).Top = constSpaceBetweenControls * j
    Me.lblData(i).Top = (txtData(i).Top - Me.lblData(i).Height)
End Sub

Sub SubCreateCheckControl(i As Integer, j As Integer)
    'Creates Checkboxen
    Me.tabData.Tab = Me.tabData.Tabs - 1
    Load Me.chkData(i)
    Set Me.chkData(i).Container = Me.tabData
    Me.chkData(i).Visible = True
    Me.chkData(i).DataField = RSMetaData.Fields("column_name").Value
    Me.chkData(i).Caption = RSMetaData.Fields("column_name").Value
    If RSMetaData.Fields("is_nullable").Value = "YES" Then
        Me.chkData(i).Tag = "Y"
    Else
        Me.chkData(i).Tag = "N"
    End If
    Me.chkData(i).Tag = Me.chkData(i).Tag & "S"
    Me.chkData(i).Top = constSpaceBetweenControls * j
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    RSMetaData.Close
    Set RSMetaData = Nothing
    RSTableData.Close
    Set RSTableData = Nothing
    CN.Close
    Set CN = Nothing
End Sub

Sub SubBinddata()
    'Binds the Recordset on the controls
    Dim i As Integer
    SubCreateRecordset RSTableData, "Select * from [" & strTable & "]"
    For i = 0 To Me.Controls.Count - 1
        If (TypeOf Me.Controls(i) Is TextBox Or TypeOf Me.Controls(i) Is CheckBox) Then
            If Me.Controls(i).DataField <> "" Then Set Me.Controls(i).DataSource = RSTableData
        End If
    Next i
    Set Me.Data.DataSource = RSTableData
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
    'Validates the Data typed in the Textboxes, they must match the datatypes
    If KeyAscii = 8 Then Exit Sub
    Select Case Right(txtData(Index).Tag, 1)
        Case "I":
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
        Case "D":
            If Chr(KeyAscii) = "." Then Exit Sub
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
        Case "F":
            If Chr(KeyAscii) = "," Then Exit Sub
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
    End Select
End Sub

Private Sub txtData_Validate(Index As Integer, Cancel As Boolean)
    'Validates the Data typed in the Textboxes, they must match the datatypes
    Select Case Right(txtData(Index).Tag, 1)
        Case "I", "F":
            If Not IsNumeric(txtData(Index).Text) Then
                MsgBox "Please insert numeric data !", vbInformation
                txtData(Index).SetFocus
            End If
        Case "D":
            If Not IsDate(txtData(Index).Text) Then
                MsgBox "Please insert a date !", vbInformation
                txtData(Index).SetFocus
            End If
    End Select
End Sub
