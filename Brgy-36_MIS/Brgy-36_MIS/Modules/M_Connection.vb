Module M_Connection

    Public conDB As New ADODB.Connection
    Public rsDB As New ADODB.Recordset
    'User's Information
    Public strUserName As String
    Public intUserLevel As Integer

    Public strQuery As String


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
