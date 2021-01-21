Public Class F_Login

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        On Error GoTo errLogin
        For Each ctrl As Control In Me.Controls
            If ctrl.GetType() Is GetType(TextBox) Then
                If ctrl.Text = "" Then
                    MsgBox("Please Fill out User's " & IIf(ctrl.Name = "txtUser", "Name", "Password"), MsgBoxStyle.Exclamation, "Invalid Input")
                Else
                    If ctrl.Name = "txtUser" Then
                        Call MISConnect()
                        rsDB = New ADODB.Recordset
                        strQuery = ""
                        strQuery = "SELECT * FROM M_UserAccounts" & vbCrLf
                        strQuery &= "WHERE UserName COLLATE Latin1_General_CS_AS = '" & txtUser.Text & "'"
                        rsDB.Open(strQuery, conDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)

                        If rsDB.EOF Then
                            MsgBox("User does not exist!", MsgBoxStyle.Exclamation, "Invalid Input")
                            Exit Sub
                        Else
                            rsDB = New ADODB.Recordset
                            strQuery = ""
                            strQuery = "SELECT * FROM M_UserAccounts" & vbCrLf
                            strQuery &= "WHERE Password COLLATE Latin1_General_CS_AS = '" & txtPassword.Text & "'"
                            rsDB.Open(strQuery, conDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)

                            If rsDB.EOF Then
                                MsgBox("User's Password is incorrect!", MsgBoxStyle.Exclamation, "Invalid Input")
                                Exit Sub
                            Else
                                intUserLevel = rsDB.Fields(3).Value()
                                strUserName = rsDB.Fields(1).Value()
                                Me.Hide()
                                _mdi_MIS.Show()
                                Call fn_ClearField(Me)
                            End If

                        End If
                    End If
                End If
            End If
        Next

errLogin:
        If Err.Number <> 0 Then
            MsgBox(Err.Number & "-->" & Err.Description, MsgBoxStyle.Critical, "Login Error")
        End If
    End Sub

    Private Sub chkShowHide_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowHide.CheckedChanged
        If chkShowHide.Checked Then
            chkShowHide.Text = "&Hide"
            txtPassword.PasswordChar = ""
        Else
            chkShowHide.Text = "&Show"
            txtPassword.PasswordChar = "•"
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class