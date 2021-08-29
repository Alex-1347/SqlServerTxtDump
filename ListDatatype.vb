Partial Module Module1
    Sub ListDataType()
        ColumnTypeList = New ArrayList
        Try
            CN1 = New SqlClient.SqlConnection(My.Settings.SQLServerConnectionString1)
            CN1.Open()
            Dim CMD1 = New SqlClient.SqlCommand("select * from sys.databases", CN1)
            Dim RDR1 = CMD1.ExecuteReader
            'https://docs.microsoft.com/en-us/dotnet/api/system.data.sqlclient.sqldatareader.getschematable?view=dotnet-plat-ext-5.0
            Dim Fields As DataTable = RDR1.GetSchemaTable
            While RDR1.Read
                Dim Arr1() As Object = New Object(Fields.Rows.Count) {}
                Dim ColumnsNumber As Integer = RDR1.GetValues(Arr1)
                For Column As Integer = 0 To ColumnsNumber - 1
                    AddColumnDataType(Arr1(Column), Fields.Rows(Column)("ColumnName"), Fields.Rows(Column)("DataType"), Fields.Rows(Column)("ColumnSize"))
                Next
            End While
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        CN1.Close()
        Dim Arr3 As List(Of String) = ColumnTypeList.ToArray.GroupBy(Function(X) X.ToString).Select(Function(X) X.Key).OrderBy(Function(X) X.ToString).ToList()
        Console.WriteLine(String.Join(vbCrLf, Arr3))
    End Sub

    Sub AddColumnDataType(Value As Object, ColumnName As String, DataType As Type, Length As Integer)
        ColumnTypeList.Add(DataType)
    End Sub

    'Print
    'System.Boolean
    'System.Byte
    'System.Byte[]
    'System.DateTime
    'System.Guid
    'System.Int16
    'System.Int32
    'System.String
End Module
