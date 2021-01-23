Module M_Functions
    Public strRequire As String

    Public clrDeactivated = Color.FromArgb(255, 188, 54)
    Public clrDeleted = Color.FromArgb(255, 84, 84)

    Public Sub fn_ClearField(container As Form)
        On Error GoTo errClear
        For Each ctrl As Control In container.Controls
            Select Case ctrl.GetType()
                Case GetType(TextBox)
                    ctrl.Text = ""
                Case GetType(GroupBox)
                    groupControls(ctrl)
                Case GetType(Label)
                    'No Event
                Case GetType(ComboBox)
                    Dim cbo As New ComboBox
                    cbo = ctrl
                    cbo.SelectedIndex = -1
                Case Else
                    'No Event
            End Select
        Next
errClear:
        If Err.Number <> 0 Then
            MsgBox(Err.Number & "-->" & Err.Description, MsgBoxStyle.Critical, "Invalid Action")
        End If
        
    End Sub

    Private Sub groupControls(group As GroupBox)
        For Each ctrl As Control In group.Controls
            Select Case ctrl.GetType()
                Case GetType(Label)
                    'No Event
                Case GetType(TextBox)
                    ctrl.Text = ""
                Case GetType(MaskedTextBox)
                    ctrl.ResetText()
                Case GetType(ComboBox)
                    Dim cbo As New ComboBox
                    cbo = ctrl
                    cbo.SelectedIndex = -1
                Case GetType(GroupBox)
                    groupControls(ctrl)
                Case Else
                    'No Event
            End Select
        Next
    End Sub

    Public Function subCheckRequire(container As Control) As Boolean
        For Each ctrl As Control In container.Controls
            Select Case ctrl.GetType()
                Case GetType(GroupBox)
                    subCheckRequire(ctrl)
                Case Else
                    If Left(ctrl.Tag, 1) = "*" Then
                        If ctrl.Text = "" Then
                            strRequire &= "- " & Mid(ctrl.Tag, 2, Len(ctrl.Tag) - 1) & vbCrLf
                        End If
                    End If
            End Select
        Next
        If strRequire <> "" Then
            Return True
        End If
    End Function

    Public Sub subRowColor(datTable As DataGridView)
        Try
            For Each row As DataGridViewRow In datTable.Rows
                Select Case row.Cells("status").Value
                    Case 0 'Deleted
                        row.DefaultCellStyle.BackColor = clrDeleted
                    Case 1 'Deactivated
                        row.DefaultCellStyle.BackColor = clrDeactivated
                End Select
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function fn_checkNull(var As Object) As String
        Dim strReturn As String

        Select Case var.GetType()
            Case GetType(String)
                If IsDBNull(var) Then
                    strReturn = ""
                Else
                    strReturn = var
                End If
            Case GetType(Integer)
                If IsDBNull(var) Then
                    strReturn = "0"
                Else
                    strReturn = var
                End If
        End Select
        Return strReturn
    End Function

    Public Sub taskMode(intTaskMode As Integer, container As Control) '0->ReadOnly || 1->Create/Register || 2-> Modify/Update
        Dim blnWrite As Boolean = False
        Dim blnShow As Boolean = False

        Select Case intTaskMode
            Case 0 'Read Only
                blnWrite = False
                blnShow = False
            Case 1 'Create/Register
                blnWrite = True
                blnShow = True
            Case 2 'Modify/Update
                blnWrite = True
                blnShow = False
        End Select

        Try
            For Each ctrl As Control In container.Controls
                Select Case ctrl.GetType()
                    Case GetType(GroupBox)
                        taskMode(intTaskMode, ctrl)
                    Case Else
                        If Strings.Right(ctrl.Tag, 3) = "_up" Then
                            If ctrl.GetType() = GetType(TextBox) Then
                                Dim txt As TextBox
                                txt = ctrl
                                txt.ReadOnly = Not blnWrite
                                txt.BackColor = Color.White
                            ElseIf ctrl.GetType() = GetType(MaskedTextBox) Then
                                Dim mtxt As MaskedTextBox
                                mtxt = ctrl
                                mtxt.ReadOnly = Not blnWrite
                                mtxt.BackColor = Color.White
                            Else
                                ctrl.Enabled = blnWrite
                            End If
                        End If
                        If ctrl.Tag = "reqSign" Then
                            ctrl.Visible = blnShow
                        End If
                End Select
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Module
