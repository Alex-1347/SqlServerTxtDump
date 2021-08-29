Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Partial Public Class DBDumper
    Public Shared Sub DumpTableToFile(ByVal CN1 As SqlConnection, ByVal TableName As String, ByVal DestinationFile As String, Optional WhereClause As String = "1=1", Optional ColumnList As String = "*")
        CN1.Open()
        Using OutFile = File.CreateText(DestinationFile)
            Using CMD = New SqlCommand("select " & ColumnList & "  from " & TableName & " where " & WhereClause, CN1)
                Try
                    Using RDR1 As SqlDataReader = CMD.ExecuteReader()
                        Dim SchemaTable As Tuple(Of String, Type, Integer)() = GetColumnsProperty(RDR1).ToArray()
                        Dim NumFields As Integer = SchemaTable.Length
                        Dim RowCount As Integer
                        RowCount = 0
                        OutFile.WriteLine(String.Join(My.Settings.FieldDivider, SchemaTable.Select(Function(X) X.Item1)))
                        If RDR1.HasRows Then
                            While RDR1.Read()
                                RowCount += 1
                                Dim Arr1() As Object = New Object(NumFields) {}
                                Dim ColumnsNumber As Integer = RDR1.GetValues(Arr1)
                                Dim ColumnValues As String() = Enumerable.Range(0, NumFields).
                                    Select(Function(i) SqlFieldToString(Arr1(i), SchemaTable(i).Item1, Arr1(i).GetType, SchemaTable(i).Item3)).
                                    Select(Function(SqlStringResult) String.Concat("""", Trim(SqlStringResult).Replace("""", """"""), """")).ToArray()
                                OutFile.WriteLine(String.Join(My.Settings.FieldDivider, ColumnValues))
                                If (RowCount Mod 1000) = 0 Then Console.Write(".")
                            End While
                        End If

                    End Using
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                    OutFile.WriteLine(ex.Message)
                End Try
            End Using
        End Using
        CN1.Close()
        Console.WriteLine()
    End Sub
    Private Shared Iterator Function GetColumnsProperty(ByVal RDR As IDataReader) As IEnumerable(Of Tuple(Of String, Type, Integer))
        For Each Row As DataRow In RDR.GetSchemaTable().Rows
            Yield New Tuple(Of String, Type, Integer)(Row("ColumnName"), Row("DataType"), Row("ColumnSize"))
        Next
    End Function

End Class
