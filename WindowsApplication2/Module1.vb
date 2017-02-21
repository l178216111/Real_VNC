Module Module1
    Dim Wmi As New System.Management.ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapterConfiguration")
    Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
    Dim Database = "Data source=C:\Database4.mdb"
    Dim Conn As OleDb.OleDbConnection
    Dim Cmd As OleDb.OleDbCommand
    Dim Rd As OleDb.OleDbDataReader
    Dim SQL, right1, d As String
    Sub main()
        Dim a As String
        For Each WmiObj As Management.ManagementObject In Wmi.Get
            If CBool(WmiObj("IPEnabled")) Then
                a = WmiObj("MACAddress")
            End If
        Next
        Conn = New OleDb.OleDbConnection(Provider & ";" & Database)
        Conn.Open()
        SQL = "Select Mac From table2 where Mac =" & d
        Cmd = New OleDb.OleDbCommand(SQL, Conn)
        Try
            Rd = Cmd.ExecuteReader()
        Catch ex As Exception
            MsgBox("pls check your pression")
            right1 = 1
        End Try
        If right1 <> 1 Then
            Form1.Show()
        End If
    End Sub
End Module
