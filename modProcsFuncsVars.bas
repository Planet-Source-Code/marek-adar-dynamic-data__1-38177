Attribute VB_Name = "modProcsFuncsVars"
Public Const constSpaceBetweenControls As Integer = 600

Sub Main()
    frmConnect.Show vbModal
    frmData.Show
End Sub

Sub SubFillCombos(RS As ADODB.Recordset, CBO As ComboBox)
    On Error Resume Next
    
    CBO.Clear
    CBO.Text = ""
    RS.MoveFirst
    Do Until RS.EOF = True
        CBO.AddItem RS.Fields(0)
        RS.MoveNext
    Loop
    CBO.ListIndex = 0
End Sub

Sub SubUnlockLockForm(Formular As Form, Locked As Boolean)
    'Locks or unlocks the Controls
    Dim i As Integer
    
    For i = 0 To Formular.Controls.Count - 1
        If TypeOf Formular.Controls(i) Is TextBox Then
            Formular.Controls(i).Locked = Locked
            If Locked = True Then
                Formular.Controls(i).BackColor = vbWhite
            Else
                Formular.Controls(i).BackColor = vbBlue
            End If
        End If
        If TypeOf Formular.Controls(i) Is CheckBox Then Formular.Controls(i).Enabled = Not Locked
        If TypeOf Formular.Controls(i) Is CommandButton Then
            If Formular.Controls(i).Tag = "L" Then
                Formular.Controls(i).Visible = Locked
            Else
                Formular.Controls(i).Visible = Not Locked
            End If
        End If
        If TypeOf Formular.Controls(i) Is DataGrid Then Formular.Controls(i).Enabled = Locked
            
    Next i

End Sub

Sub SubShowError(ErrNumber As Long, ErrDescription As String)
    'Controls Error messages
    On Error Resume Next
    
    Dim intFileID As Integer
    intFileID = FreeFile()
    Open App.Path & "\ErrMsg.log" For Append As #intFileID
        Print #intFileID, Format$(Date, "DD.MM.YYYY") & " " & Format$(Time, "HH:MM:SS") & " " & CStr(ErrNumber) & " " & ErrDescription
    Close #intFileID
    MsgBox "An error occurred: " & vbNewLine & ErrDescription, vbExclamation
End Sub

Sub SubCenterForm(Formular As Form)
    'Centers the form
    Dim lngTop As Long
    
    lngTop = (Screen.Height - Formular.Height) / 2
    Formular.Left = (Screen.Width - Formular.Width) / 2
    If lngTop < 0 Then
        Formular.Top = 10
    Else
        Formular.Top = lngTop
    End If
End Sub

Function funcCheckInputs(Formular As Form) As Boolean
    'checks requirde fields
    Dim i As Integer
    
    For i = 0 To Formular.Controls.Count - 1
        If TypeOf Formular.Controls(i) Is TextBox Then
            If Trim(Formular.Controls(i).Text) = "" And Left(Formular.Controls(i).Tag, 1) = "N" Then
                Formular.Controls(i).BackColor = vbRed
                funcCheckInputs = True
            End If
        End If
    Next i
    If funcCheckInputs = True Then MsgBox "You have to insert data into the red marked textfileds !", vbInformation

End Function
