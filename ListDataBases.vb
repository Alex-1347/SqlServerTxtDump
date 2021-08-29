Partial Module Module1
    Sub ListDataBases()
        Str1 = New Text.StringBuilder
        Databases = New List(Of String)
        Try
            CN1 = New SqlClient.SqlConnection(My.Settings.SQLServerConnectionString1)
            CN1.Open()
            Dim CMD1 = New SqlClient.SqlCommand("select * from sys.databases order by name", CN1)
            Dim RDR1 = CMD1.ExecuteReader
            Dim Fields As DataTable = RDR1.GetSchemaTable
            While RDR1.Read
                Dim Arr1() As Object = New Object(Fields.Rows.Count) {}
                Dim ColumnsNumber As Integer = RDR1.GetValues(Arr1)
                For Column As Integer = 0 To ColumnsNumber - 1
                    'ignore four system DB
                    If Arr1(0) = "master" Or Arr1(0) = "tempdb" Or Arr1(0) = "model" Or Arr1(0) = "msdb" Then
                        Exit For
                    Else
                        CheckDatabasesColumn(Arr1(Column), Fields.Rows(Column)("ColumnName"), Fields.Rows(Column)("DataType"), Fields.Rows(Column)("ColumnSize"))
                    End If
                Next
            End While
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        CN1.Close()
        Console.WriteLine(Left(Str1.ToString, Len(Str1.ToString) - 2))
    End Sub

    Dim Str1 As Text.StringBuilder
    Sub CheckDatabasesColumn(Value As Object, ColumnName As String, DataType As Type, Length As Integer)
        Select Case ColumnName
            Case "name"
                Databases.Add(SqlFieldToString(Value, ColumnName, DataType, Length))
                Str1.Append($"{vbCrLf & vbCrLf}{Databases.Count}. database={SqlFieldToString(Value, ColumnName, DataType, Length)}, ")
            Case "database_id"
                Str1.Append($"{ColumnName}={SqlFieldToString(Value, ColumnName, DataType, Length)}, ")
            Case "owner_sid"
                Str1.Append($"{ColumnName}={SqlFieldToString(Value, ColumnName, DataType, Length)}, ")
            Case "create_date"
                Str1.Append($"{ColumnName}={SqlFieldToString(Value, ColumnName, DataType, Length)}, ")
            Case "compatibility_level"
                Str1.Append($"{ColumnName}={SqlFieldToString(Value, ColumnName, DataType, Length)}, ")
            Case "collation_name"
                Str1.Append($"{ColumnName}={SqlFieldToString(Value, ColumnName, DataType, Length)}, ")
            Case "user_access_desc"
                Str1.Append($"{ColumnName}={SqlFieldToString(Value, ColumnName, DataType, Length)}, ")
            Case "state_desc"
                Str1.Append($"{ColumnName}={SqlFieldToString(Value, ColumnName, DataType, Length)}, ")
            Case "recovery_model_desc"
                Str1.Append($"{ColumnName}={SqlFieldToString(Value, ColumnName, DataType, Length)}, ")
        End Select

    End Sub
End Module
