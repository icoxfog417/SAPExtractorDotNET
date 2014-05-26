Imports SAPExtractorDotNET

Namespace Util

    ''' <summary>
    ''' Utility Class to display result(Dictionary,DataTable)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ResultWriter

        Public Shared Sub Write(ByVal table As DataTable, Optional ByVal maxCount As Integer = 10)

            Dim count As Integer = 0
            Dim columns = (From col As DataColumn In table.Columns).ToList
            For Each row As DataRow In table.Rows
                Dim line As String = ""
                columns.ForEach(Sub(c) line += c.ColumnName + ":" + row(c.ColumnName) + " ")
                Console.WriteLine(Trim(line))
                If count > 10 Then Exit For
                count += 1
            Next

        End Sub

        Public Shared Sub Write(ByVal fields As List(Of SAPFieldItem))

            For Each f As SAPFieldItem In fields
                Console.WriteLine(f.FieldId + ":" + f.FieldText)
            Next

        End Sub

    End Class

End Namespace

