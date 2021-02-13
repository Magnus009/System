Public Class F_ResidentsRecord

    Private Sub F_ResidentsRecord_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call loadResidentRecords()
    End Sub

    Private Sub loadResidentRecords(Optional strFilter As String = "")
        Try
            datResidents.Columns.Clear()
            strQuery = "SELECT R.Code 'ID', R.FamilyName + ', ' + R.GivenName + ' ' + R.MiddleName 'NAME'," & vbCrLf
            strQuery &= "R.ContactNo AS 'CONTACT No.', H.HouseholdNo + ' ' + H.Barangay + ', ' + H.Street + ', ' + H.Municipality 'ADDRESS' FROM Residents R" & vbCrLf
            strQuery &= "LEFT JOIN HouseholdMember HM ON R.Code = HM.ResidentCode" & vbCrLf
            strQuery &= "LEFT JOIN Household H ON H.HouseholdNo = HM.HouseholdNo" & vbCrLf
            If strFilter <> "" Then
                strQuery &= "WHERE" & strFilter
            End If

            datResidents.DataSource = SqlCli_MIS(strQuery)
            datResidents.DataMember = "Table"

            'add Button
            Dim btnView As New DataGridViewButtonColumn

            btnView.Text = "•••"
            btnView.UseColumnTextForButtonValue = True
            datResidents.Columns.Add(btnView)
            datResidents.Columns(datResidents.ColumnCount - 1).Width = 50
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub datResidents_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles datResidents.CellClick
        If e.ColumnIndex = 0 Then
            F_Resident.intTaskMode = IIf(intUserLevel = 0, 2, 0)
            Call F_Resident.loadResidentRecord(datResidents.Rows(e.RowIndex).Cells(1).Value)
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim strFilter As String = ""

        If txtCode.Text <> "" Then
            strFilter = " R.Code = " & txtCode.Text & " OR"
        End If
        If txtName.Text <> "" Then
            strFilter &= " R.FamilyName + ', ' + R.GivenName + ' ' + R.MiddleName LIKE '%" & txtName.Text & "%' OR"
        End If
        If txtHouseNo.Text <> "" Then
            strFilter &= " H.HouseholdNo = " & txtHouseNo.Text & " OR"
        End If
        If strFilter <> "" Then
            strFilter = Strings.Left(strFilter, Len(strFilter) - 2)
        End If
        Call loadResidentRecords(strFilter)
    End Sub
End Class