Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports SAP.Middleware.Connector
Imports SAPExtractorDotNET
Imports System.Globalization

<TestClass()>
Public Class TableExtraction

    Private Const TestTable As String = "T001"
    Private Const TestDestination As String = "SILENT_LOGIN"

    <TestMethod()>
    Public Sub GetColumnDefine()
        Dim connector As New SAPConnector(TestDestination)

        Try
            Dim connection As RfcDestination = connector.Login

            Dim tableExtractor As New SAPTableExtractor(TestTable)
            Dim columns As List(Of SAPFieldItem) = tableExtractor.GetColumnFields(connection)

            For Each column As SAPFieldItem In columns
                Console.WriteLine(column.FieldId + ":" + column.FieldText)
            Next

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Assert.Fail()
        End Try

    End Sub

    <TestMethod()>
    Public Sub ExtractTable()

        Dim connector As New SAPConnector(TestDestination)

        Try
            Dim connection As RfcDestination = connector.Login

            Dim tableExtractor As New SAPTableExtractor(TestTable)
            Dim conditions As New List(Of SAPFieldItem)
            Dim fields As New List(Of SAPFieldItem)

            For Each column As String In {"BUKRS", "BUTXT"}
                conditions.Add(New SAPFieldItem(column))
            Next

            fields.Add(New SAPFieldItem("SPRAS").IsEqualTo(CultureInfo.CurrentCulture.TwoLetterISOLanguageName.Substring(0, 1).ToUpper))

            Dim table As DataTable = tableExtractor.Invoke(connection, conditions, fields)

            Dim count As Integer = 0
            Dim columns = (From col As DataColumn In table.Columns).ToList
            For Each row As DataRow In table.Rows
                Dim line As String = ""
                columns.ForEach(Sub(c) line += c.ColumnName + ":" + row(c.ColumnName) + " ")
                Console.WriteLine(Trim(line))
                If count > 10 Then Exit For
                count += 1
            Next

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Assert.Fail()
        End Try

    End Sub

End Class