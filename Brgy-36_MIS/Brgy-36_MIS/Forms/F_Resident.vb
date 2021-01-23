Public Class F_Resident
    Public intTaskMode As Integer = 1 '0->ReadOnly || 1->Create/Register || 2-> Modify/Update
    Dim strCompleteAddress(4) As String
    Private Sub chkPWD_CheckedChanged(sender As Object, e As EventArgs) Handles chkPWD.CheckedChanged
        Call subDisabled()
    End Sub

    Public Sub loadResidentRecord(strResidentCode As String)
        Try
            Dim dsResidentInfo As New DataSet

            strQuery = "SELECT R.Code, R.FamilyName, R.GivenName, R.MiddleName, R.ExtensionName, R.BirthPlace, R.BirthDate," & vbCrLf
            strQuery &= "R.Gender, R.Citizenship, R.CivilStatus, R.ContactNo, R.Occupation, R.DateOfCaseStudy," & vbCrLf
            strQuery &= "R.SamahanID, R.isVoter, R.inHabitant, R.Indigent, R.isPWD, R.Disabilities," & vbCrLf
            strQuery &= "H.HouseholdNo, H.Barangay, H.Street, HM.Role, H.Municipality, H.Province, H.ContactNo FROM Residents R" & vbCrLf
            strQuery &= "LEFT JOIN HouseholdMember HM ON R.Code = HM.ResidentCode" & vbCrLf
            strQuery &= "LEFT JOIN Household H ON HM.HouseholdNo = H.HouseholdNo" & vbCrLf
            strQuery &= "WHERE R.Code = '" & strResidentCode & "'"
            dsResidentInfo = SqlCli_MIS(strQuery)

            With dsResidentInfo
                'Personal Info
                txtID.Text = .Tables(0).Rows(0)(0)
                txtLName.Text = .Tables(0).Rows(0)(1)
                txtFName.Text = .Tables(0).Rows(0)(2)
                txtMName.Text = fn_checkNull(.Tables(0).Rows(0)(3))
                txtEName.Text = fn_checkNull(.Tables(0).Rows(0)(4))
                txtBirthPlace.Text = .Tables(0).Rows(0)(5)
                dtpBirthdate.Value = .Tables(0).Rows(0)(6)
                cboSex.SelectedIndex = IIf(.Tables(0).Rows(0)(7) = "M", 1, 0)
                txtCitizenship.Text = .Tables(0).Rows(0)(8)
                cboCivilStatus.SelectedIndex = .Tables(0).Rows(0)(9)
                txtContactNo.Text = fn_checkNull(.Tables(0).Rows(0)(10))
                txtOccupation.Text = fn_checkNull(.Tables(0).Rows(0)(11))
                dtpCaseStudy.Value = .Tables(0).Rows(0)(12)
                cboSamahan.SelectedIndex = .Tables(0).Rows(0)(13)
                chkVoter.Checked = .Tables(0).Rows(0)(14)
                chkInHabitant.Checked = .Tables(0).Rows(0)(15)
                chkIndigent.Checked = .Tables(0).Rows(0)(16)
                chkPWD.Checked = .Tables(0).Rows(0)(17)
                txtDisability.Text = fn_checkNull(.Tables(0).Rows(0)(18))
                'Household Info
                txtHouseNo.Text = fn_checkNull(.Tables(0).Rows(0)(19))
                txtBarangay.Text = fn_checkNull(.Tables(0).Rows(0)(20))
                txtStreet.Text = fn_checkNull(.Tables(0).Rows(0)(21))
                cboRole.SelectedIndex = fn_checkNull(.Tables(0).Rows(0)(22))
                txtMunicipality.Text = fn_checkNull(.Tables(0).Rows(0)(23))
                txtProvince.Text = fn_checkNull(.Tables(0).Rows(0)(24))
                txtHouseContactNo.Text = fn_checkNull(.Tables(0).Rows(0)(25))
            End With
            Me.MdiParent = _mdi_MIS
            Me.Show()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub F_Resident_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'taskMode(intTaskMode, Me)
        Call subDisabled()
        strCompleteAddress(0) = txtHouseNo.Text & " "
        strCompleteAddress(1) = txtStreet.Text & ", "
        strCompleteAddress(2) = txtBarangay.Text & ", "
        strCompleteAddress(3) = txtMunicipality.Text & ", "
        strCompleteAddress(4) = txtProvince.Text
        Call subCompleteAddress()
    End Sub

    Private Sub subDisabled()
        If chkPWD.Checked Then
            lblDiasability.Visible = True
            txtDisability.Visible = True
        Else
            lblDiasability.Visible = False
            txtDisability.Visible = False
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        Call fn_ClearField(Me)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        On Error GoTo errSaving
        If subCheckRequire(Me) Then
            MsgBox("Please complete the required fields(*):" & vbCrLf & strRequire, MsgBoxStyle.Exclamation, "Required Items")
            strRequire = ""
        Else
            intTaskMode = 1
            If intTaskMode = 1 Then
                Call MISConnect()
                strQuery = ""
                strQuery = "EXECUTE sp_registerResident " & vbCrLf
                strQuery &= "'" & txtLName.Text & "'," & vbCrLf
                strQuery &= "'" & txtFName.Text & "'," & vbCrLf
                strQuery &= "'" & txtMName.Text & "'," & vbCrLf
                strQuery &= "'" & txtEName.Text & "'," & vbCrLf
                strQuery &= "'" & Strings.Left(UCase(cboSex.Text), 1) & "'," & vbCrLf
                strQuery &= "'" & Format(dtpBirthdate.Value, "Short Date") & "'," & vbCrLf
                strQuery &= "'" & txtBirthPlace.Text & "'," & vbCrLf
                strQuery &= "'" & txtCitizenship.Text & "'," & vbCrLf
                strQuery &= cboCivilStatus.SelectedIndex & "," & vbCrLf
                strQuery &= "'" & txtContactNo.Text & "'," & vbCrLf
                strQuery &= "'" & txtOccupation.Text & "'," & vbCrLf
                strQuery &= cboSamahan.SelectedIndex & "," & vbCrLf
                strQuery &= chkVoter.Checked & "," & vbCrLf
                strQuery &= chkInHabitant.Checked & "," & vbCrLf
                strQuery &= chkIndigent.Checked & "," & vbCrLf
                strQuery &= chkPWD.Checked & "," & vbCrLf
                strQuery &= "'" & txtDisability.Text & "'," & vbCrLf
                strQuery &= "'" & dtpCaseStudy.Value & "'," & vbCrLf
                strQuery &= "'" & txtHouseNo.Text & "'," & vbCrLf
                strQuery &= "'" & txtBarangay.Text & "'," & vbCrLf
                strQuery &= "'" & txtStreet.Text & "'," & vbCrLf
                strQuery &= "'" & txtMunicipality.Text & "'," & vbCrLf
                strQuery &= "'" & txtProvince.Text & "'," & vbCrLf
                strQuery &= "'" & txtHouseContactNo.Text & "'," & vbCrLf
                strQuery &= cboRole.SelectedIndex & "," & vbCrLf
                strQuery &= "'" & strUserName & "'"
                conDB.Execute(strQuery)

                MsgBox("Resident Registered Successfully", MsgBoxStyle.Information, "Registration")
            End If
        End If
errSaving:
        If Err.Number <> 0 Then
            MsgBox(Err.Number & "-->" & Err.Description, MsgBoxStyle.Critical, "Resident's Information Error")
        End If
    End Sub

    Private Sub dtpBirthdate_ValueChanged(sender As Object, e As EventArgs) Handles dtpBirthdate.ValueChanged
        txtAge.Text = Math.Floor(DateDiff(DateInterval.Day, dtpBirthdate.Value, Now) / 365.25)
    End Sub

    Private Sub txtHouseNo_Leave(sender As Object, e As EventArgs) Handles txtHouseNo.Leave
        If txtHouseNo.Text <> "" Then
            strCompleteAddress(0) = txtHouseNo.Text & " "
            Call subCompleteAddress()
        End If
    End Sub

    Private Sub txtStreet_LostFocus(sender As Object, e As EventArgs) Handles txtStreet.LostFocus
        If txtStreet.Text <> "" Then
            If UCase(txtStreet.Text).Contains("STREET") Then
                strCompleteAddress(1) = txtStreet.Text & ", "
            Else
                strCompleteAddress(1) = txtStreet.Text & " Street , "
            End If
            Call subCompleteAddress()
        End If
    End Sub

    Private Sub txtBarangay_LostFocus(sender As Object, e As EventArgs) Handles txtBarangay.LostFocus
        If txtBarangay.Text <> "" Then
            If UCase(txtBarangay.Text).Contains("BARANGAY") Then
                Replace(UCase(txtBarangay.Text), "BARANGAY", "Brgy")
            ElseIf UCase(txtBarangay.Text).Contains("BRGY") Then
                'No Event
            Else
                txtBarangay.Text = "Brgy " & txtBarangay.Text
            End If

            strCompleteAddress(2) = txtBarangay.Text & ", "
            Call subCompleteAddress()
        End If
    End Sub

    Private Sub txtMunicipality_LostFocus(sender As Object, e As EventArgs) Handles txtMunicipality.LostFocus
        If txtMunicipality.Text <> "" Then
            strCompleteAddress(3) = txtMunicipality.Text & ", "
            Call subCompleteAddress()
        End If
    End Sub

    Private Sub txtProvince_LostFocus(sender As Object, e As EventArgs) Handles txtProvince.LostFocus
        If txtProvince.Text <> "" Then
            strCompleteAddress(4) = txtProvince.Text
            Call subCompleteAddress()
        End If
    End Sub

    Private Sub subCompleteAddress()
        txtCompleteAdd.Text = strCompleteAddress(0)
        txtCompleteAdd.Text &= strCompleteAddress(1)
        txtCompleteAdd.Text &= strCompleteAddress(2)
        txtCompleteAdd.Text &= strCompleteAddress(3)
        txtCompleteAdd.Text &= strCompleteAddress(4)
    End Sub

End Class