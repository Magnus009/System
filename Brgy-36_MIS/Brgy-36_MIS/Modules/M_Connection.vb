Module M_Connection
    'ADODB
    Public conDB As New ADODB.Connection
    Public rsDB As New ADODB.Recordset
    'User's Information
    Public strUserName As String
    Public intUserLevel As Integer

    Public strQuery As String
    Public Function SqlCli_MIS(strQueryCommand As String)
        Dim result = Nothing
        Try
            Dim sqlConnect As New SqlClient.SqlConnection(My.Resources.ConnectionString)
            Dim sqlAdapter As New SqlClient.SqlDataAdapter(strQueryCommand, sqlConnect)

            If InStr(strQueryCommand, "SELECT") > 0 Then
                Dim dsNew As New DataSet

                sqlConnect.Open()
                sqlAdapter.Fill(dsNew)
                sqlConnect.Close()
                result = dsNew
            Else
                Dim sqlCommand As New SqlClient.SqlCommand(strQueryCommand, sqlConnect)

                sqlCommand.CommandType = CommandType.Text
                sqlConnect.Open()
                If sqlCommand.ExecuteNonQuery.Equals(0) Then
                    result = False
                Else
                    result = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return result
    End Function

    Public Sub MISConnect()
        On Error GoTo errMIS

        With conDB
            If .State = 1 Then
                .Close()
            End If

            .ConnectionString = "Provider=SQLOLEDB;Data Source=sd_sql_training;Persist Security Info=True;User ID=sa;Password=81at84;Initial Catalog=MIS"
            .ConnectionTimeout = 0
            .CommandTimeout = 0
            .Open()
        End With

errMIS:
        If Err.Number <> 0 Then
            MsgBox(Err.Number & "-->" & Err.Description, MsgBoxStyle.Critical, "Connection Error")
            conDB = Nothing
        End If
    End Sub
End Module
