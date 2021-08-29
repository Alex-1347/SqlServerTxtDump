Partial Module Module1

    Dim CN1 As SqlClient.SqlConnection
    Sub Main()
Again:
        Console.WriteLine(My.Settings.SQLServerConnectionString1)
        Console.WriteLine("Select mode:")
        Console.WriteLine("1. List DataType, Test Connection.")
        Console.WriteLine("2. List DataBases")
        Console.Write(">")
        Dim Mode1 As Integer
        Try
            Mode1 = CInt(Console.ReadLine)
        Catch ex As Exception
            GoTo Again
        End Try

        Select Case Mode1
            Case 1 : ListDataType()
            Case 2
DataBases:
                ListDataBases()
                GoTo Objects
            Case 3
Objects:
                If Databases Is Nothing Then GoTo DataBases
                Console.WriteLine()
                Console.WriteLine("Set current DB number or exit(empty)")
                Console.Write(">")
                Dim Inpt1 As String = Console.ReadLine
                If Inpt1 = "" Then End
                Try
                    CurrentDbNumber = Inpt1
                Catch ex As Exception
                    GoTo Objects
                End Try

                ListObjects(CurrentDbNumber - 1)
DumpObj:
                Console.WriteLine()
                Console.WriteLine("Dump object (type number), or all objects (type zero), or exit to previous menu (empty)")
                Console.Write(">")
                Dim ObjNumber As Integer
                Dim Inpt2 As String = Console.ReadLine
                If Inpt2 = "" Then GoTo Objects
                Try
                    ObjNumber = CInt(Inpt2)
                Catch ex As Exception
                    GoTo DumpObj
                End Try
                If ObjNumber > Tables.Count + Views.Count Then
                    GoTo DumpObj
                End If
                OutputPath = My.Settings.DefaultOutputPath
                If OutputPath = "" Then
                    Console.WriteLine()
                    Console.WriteLine("Get output path")
                    Console.Write(">")
                    OutputPath = Console.ReadLine
                End If
                Try
                    Dim OutFileName As String
                    If ObjNumber = 0 Then
                        For i As Integer = 0 To Tables.Count - 1
                            OutFileName = $"{Databases(CurrentDbNumber - 1) & "." & Tables(i) & "." & Now.ToString("dd-MM-yyyy") & ".csv"}"
                            Console.WriteLine($"Dumping {i + 1}.{Tables(i)} to {IO.Path.Combine(OutputPath, OutFileName)}")
                            DBDumper.DumpTableToFile(CN1, $"[{Databases(CurrentDbNumber - 1)}].dbo.[{Tables(i)}]", IO.Path.Combine(OutputPath, OutFileName))
                        Next
                        For i As Integer = 0 To Views.Count - 1
                            OutFileName = $"{Databases(CurrentDbNumber - 1) & "." & Views(i) & "." & Now.ToString("dd-MM-yyyy") & ".csv"}"
                            Console.WriteLine($"Dumping {i + 1}.{Views(i)} to {IO.Path.Combine(OutputPath, OutFileName)}")
                            DBDumper.DumpTableToFile(CN1, $"[{Databases(CurrentDbNumber - 1)}].dbo.[{Views(i)}]", IO.Path.Combine(OutputPath, OutFileName))
                        Next
                        Console.WriteLine("Done.")
                    Else
                        Console.WriteLine()
                        Console.WriteLine("Get WHERE clause or (empty)")
                        Console.Write(">")
                        Dim Inpt3 As String = Console.ReadLine
                        Console.WriteLine()
                        Console.WriteLine("Get ColumnList  or (empty)")
                        Console.Write(">")
                        Dim Inpt4 As String = Console.ReadLine
                        If ObjNumber > Tables.Count Then
                            OutFileName = $"{Databases(CurrentDbNumber - 1) & "." & Views(ObjNumber - Tables.Count - 1) & "." & Now.ToString("dd-MM-yyyy") & ".csv"}"
                            Console.WriteLine($"Dumping {ObjNumber}.{Views(ObjNumber - Tables.Count - 1)} To {IO.Path.Combine(OutputPath, OutFileName)}")
                            DBDumper.DumpTableToFile(CN1, $"[{Databases(CurrentDbNumber - 1)}].dbo.[{Views(ObjNumber - Tables.Count - 1)}]", IO.Path.Combine(OutputPath, OutFileName), IIf(String.IsNullOrEmpty(Inpt3), "1=1", Inpt3), IIf(String.IsNullOrEmpty(Inpt4), "*", Inpt4))
                        Else
                            OutFileName = $"{Databases(CurrentDbNumber - 1) & "." & Tables(ObjNumber - 1) & "." & Now.ToString("dd-MM-yyyy") & ".csv"}"
                            Console.WriteLine($"Dumping {ObjNumber}.{Tables(ObjNumber - 1)} to {IO.Path.Combine(OutputPath, OutFileName)}")
                            DBDumper.DumpTableToFile(CN1, $"[{Databases(CurrentDbNumber - 1)}].dbo.[{Tables(ObjNumber - 1)}]", IO.Path.Combine(OutputPath, OutFileName), IIf(String.IsNullOrEmpty(Inpt3), "1=1", Inpt3), IIf(String.IsNullOrEmpty(Inpt4), "*", Inpt4))
                        End If
                        Console.WriteLine("Done.")
                        GoTo DumpObj
                    End If
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try

                GoTo Objects
            Case Else
                Console.WriteLine("Try again")
                GoTo Again
        End Select
        Console.ReadLine()
    End Sub

    Dim CurrentDbNumber As Integer = 0
    Dim ColumnTypeList As ArrayList
    Dim Databases As List(Of String)
    Dim Tables As List(Of String)
    Dim Views As List(Of String)
    Dim OutputPath As String

    Function SqlFieldToString(Value As Object, ColumnName As String, DataType As Type, Length As Integer) As String
        Select Case DataType
            Case GetType(DBNull)
                Return "NULL"
            Case GetType(System.Boolean)
                Return Value.ToString
            Case GetType(System.Byte)
                Return Value.ToString
            Case GetType(System.Byte())
                If Length > 1000 Then
                    Return $"HUGE BINARY DATA {Length} bytes start with {String.Join("", Enumerable.Range(0, 10).Select(Function(I) CByte(Value(I)).ToString("x2")).ToList)}..."
                Else
                    Dim Str1 = New Text.StringBuilder()
                    For Each One As Byte In Value
                        Str1.Append(One.ToString("x2"))
                    Next
                    Return Str1.ToString()
                End If

            Case GetType(System.DateTime)
                Return Value.ToString
            Case GetType(System.Double)
                Return Value.ToString
            Case GetType(System.Guid)
                Return Value.ToString
            Case GetType(System.Int16)
                Return Value.ToString
            Case GetType(System.Int32)
                Return Value.ToString
            Case GetType(System.Int64)
                Return Value.ToString
            Case GetType(System.String)
                Return Value.ToString
            Case GetType(System.Decimal)
                Return Value.ToString
            Case Else
                Console.WriteLine($"New datatype faced {DataType.Name}")
                Return Value.ToString
        End Select
    End Function




End Module
