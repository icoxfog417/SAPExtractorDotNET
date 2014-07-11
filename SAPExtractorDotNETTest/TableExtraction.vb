Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports SAP.Middleware.Connector
Imports SAPExtractorDotNET
Imports System.Globalization
Imports SAPExtractorDotNETTest.Util

<TestClass()>
Public Class TableExtraction

    Private Const TestTable As String = "T001"
    Private Const TestTableWide As String = "MARA"
    Private Const TestDestination As String = "SILENT_LOGIN"

    <TestMethod()>
    Public Sub FindTable()

        Dim connector As New SAPConnector(TestDestination)

        Try
            Dim connection As RfcDestination = connector.Login
            Dim tables As List(Of SAPTableExtractor) = SAPTableExtractor.Find(connection, "DD*")

            For i As Integer = 0 To If(tables.Count > 10, 10, tables.Count) - 1
                Dim t = tables(i)
                Console.WriteLine(t.Table + ":" + t.TableText)
            Next

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Assert.Fail()
        End Try

    End Sub

    <TestMethod()>
    Public Sub GetColumnDefine()
        Dim connector As New SAPConnector(TestDestination)

        Try
            Dim connection As RfcDestination = connector.Login

            Dim tableExtractor As New SAPTableExtractor(TestTable)
            Dim columns As List(Of SAPFieldItem) = tableExtractor.GetColumnFields(connection)

            ResultWriter.Write(columns)

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
            Dim table As DataTable = tableExtractor.Invoke(connection)
            ResultWriter.Write(table)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Assert.Fail()
        End Try

    End Sub

    <TestMethod()>
    Public Sub ExtractTableWithFields()

        Dim connector As New SAPConnector(TestDestination)

        Try
            Dim connection As RfcDestination = connector.Login

            Dim tableExtractor As New SAPTableExtractor(TestTable)
            Dim table As DataTable = tableExtractor.Invoke(connection, {"BUKRS", "BUTXT"})
            ResultWriter.Write(table)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Assert.Fail()
        End Try

    End Sub


    <TestMethod()>
    Public Sub ExtractTableWithCriteria()

        Dim connector As New SAPConnector(TestDestination)

        Try
            Dim connection As RfcDestination = connector.Login

            Dim tableExtractor As New SAPTableExtractor(TestTableWide)
            Dim options As New List(Of SAPFieldItem)

            options.Add(New SAPFieldItem("MATNR").StartsWith("A"))

            Dim table As DataTable = tableExtractor.Invoke(connection, Nothing, options)
            ResultWriter.Write(table)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Assert.Fail()
        End Try

    End Sub


End Class