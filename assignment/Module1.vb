Imports System.Data.SqlClient
Module Module1
    Public connectionstring As String = "workstation id=PS03157.mssql.somee.com;packet size=4096;user id=quangtnps03157_SQLLogin_1;pwd=lt1n7y1mn9;data source=PS03157.mssql.somee.com;persist security info=False;initial catalog=PS03157"
    'tao bien thuc thi
    Public Sub excutenonquery(sql As String)
        Dim connection As New SqlConnection(connectionstring)
        Dim command As New SqlCommand(sql, connection)
        connection.Open()
        command.ExecuteNonQuery()
        connection.Close()
    End Sub
End Module
