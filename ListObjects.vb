Partial Module Module1

    Sub ListObjects(CurrentDbNumber As Integer)
        Tables = New List(Of String)
        Views = New List(Of String)
        Try
            CN1 = New SqlClient.SqlConnection(My.Settings.SQLServerConnectionString1)
            CN1.Open()
            Dim CMD1 = New SqlClient.SqlCommand($"SELECT * FROM {Databases(CurrentDbNumber)}.INFORMATION_SCHEMA.TABLES order by TABLE_TYPE,TABLE_NAME", CN1)
            Dim RDR1 = CMD1.ExecuteReader
            Dim Fields As DataTable = RDR1.GetSchemaTable
            While RDR1.Read
                Dim Arr1() As Object = New Object(Fields.Rows.Count) {}
                Dim ColumnsNumber As Integer = RDR1.GetValues(Arr1)
                If RDR1("TABLE_TYPE") = "BASE TABLE" Then
                    Tables.Add(RDR1("TABLE_NAME"))
                ElseIf RDR1("TABLE_TYPE") = "VIEW" Then
                    Views.Add(RDR1("TABLE_NAME"))
                Else
                    Console.WriteLine(RDR1("TABLE_NAME"))
                End If
            End While
            For i As Integer = 0 To Tables.Count - 1
                Console.WriteLine($"{i + 1}. {Tables(i)} (tables)")
            Next
            For i As Integer = 0 To Views.Count - 1
                Console.WriteLine($"{Tables.Count + i + 1}. {Views(i)} (view)")
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        CN1.Close()
    End Sub

End Module
