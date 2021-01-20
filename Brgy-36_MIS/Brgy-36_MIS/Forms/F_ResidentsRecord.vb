Public Class F_ResidentsRecord

    Private Sub F_ResidentsRecord_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call loadResidentRecords()
    End Sub

    Private Sub loadResidentRecords()
        Try
            strQuery = "SELECT R.Code 'ID', R.FamilyName + ', ' + R.GivenName + ' ' + R.MiddleName 'NAME'," & vbCrLf
            strQuery &= "R.ContactNo AS 'CONTACT No.', H.HouseholdNo 'HOUSE No.', H.Street 'STREET' FROM Residents R" & vbCrLf
            strQuery &= "LEFT JOIN HouseholdMember HM ON R.Code = HM.ResidentCode" & vbCrLf
            strQuery &= "LEFT JOIN Household H ON H.HouseholdNo = HM.HouseholdNo"
            datResidents.DataSource = SqlCli_MIS(strQuery)
            datResidents.DataMember = "table"

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
            Call F_Resident.loadResidentRecord(datResidents.Rows(e.RowIndex).Cells(1).Value)
        End If
    End Sub

End Class