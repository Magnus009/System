Public Class F_HouseholdList
    Private Sub loadHouses()
        Try
            With datHouses
                Dim dsHouses As New DataSet

                .Columns.Clear()
                strQuery = ""
                strQuery = "SELECT H.HouseholdNo 'HOUSE #'," & vbCrLf
                strQuery &= "H.Barangay + ', ' + H.Street + ' Street, ' + H.Municipality  + ', ' +  H.Province 'ADDRESS'," & vbCrLf
                strQuery &= "H.ContactNo 'CONTACT #', R.FamilyName 'FAMILY NAME' FROM Household H" & vbCrLf
                strQuery &= "INNER JOIN HouseholdMember HM ON H.HouseholdNo = HM.HouseholdNo AND HM.Role = 1" & vbCrLf
                strQuery &= "LEFT JOIN Residents R ON HM.ResidentCode = R.Code" & vbCrLf
                strQuery &= "WHERE H.DeletedDate IS NULL" & vbCrLf
                If cboSearch.Text <> "" And txtSearch.Text <> "" Then
                    strQuery &= "AND " & IIf(cboSearch.SelectedIndex = 5, "R.", "H.") & cboSearch.Text & " LIKE '%" & txtSearch.Text & "%'"
                End If
                dsHouses = SqlCli_MIS(strQuery)
                .DataSource = dsHouses
                .DataMember = "table"

                'add Button
                Dim btnSelect As New DataGridViewButtonColumn

                btnSelect.Text = "SELECT"
                btnSelect.UseColumnTextForButtonValue = True
                .Columns.Add(btnSelect)

                .Columns(0).Width = .Width * 0.12
                .Columns(2).Width = .Width * 0.2
                .Columns(4).Width = .Width * 0.1
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub F_HouseholdList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call loadHouses()
    End Sub

    Private Sub datHouses_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles datHouses.CellContentClick
        If e.ColumnIndex = 4 Then
            If MsgBox("Do you want to select this house as your reference?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "SELECT HOUSE") = vbYes Then
                F_Resident.txtHouseNo.Text = datHouses.Rows(e.RowIndex).Cells(0).Value
                Me.Close()
                F_Resident.getHouseInfo(datHouses.Rows(e.RowIndex).Cells(0).Value)
                Call F_Resident.subCompleteAddress()
            End If
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Call loadHouses()
    End Sub
End Class