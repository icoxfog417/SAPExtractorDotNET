Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports SAP.Middleware.Connector
Imports SAPExtractorDotNET
Imports System.Configuration


<TestClass()>
Public Class QueryExtraction

    Private Const TestDestination As String = "SILENT_LOGIN"

    Public ReadOnly Property TestQuery As String
        Get
            Return ConfigurationManager.AppSettings("testQuery")
        End Get
    End Property

    Public ReadOnly Property TestUserGroup As String
        Get
            Return ConfigurationManager.AppSettings("testUserGroup")
        End Get
    End Property


    <TestMethod()>
    Public Sub GetQueryParameters()

        Dim connector As New SAPConnector(TestDestination)

        Try
            Dim connection As RfcDestination = connector.Login

            Dim query As New SAPQueryExtractor(TestQuery, TestUserGroup)
            Dim queryParams As List(Of SAPFieldItem) = query.GetSelectFields(connection)

            For Each param As SAPFieldItem In queryParams
                Console.WriteLine(param.FieldId + ":" + param.FieldText)
            Next

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Assert.Fail()
        End Try

    End Sub

    <TestMethod()>
    Public Sub ExtractQuery()

        Dim connector As New SAPConnector(TestDestination)

        Try
            Dim connection As RfcDestination = connector.Login

            Dim query As New SAPQueryExtractor(TestQuery, TestUserGroup)
            Dim param As SAPFieldItem = query.GetSelectFields(connection).Where(Function(p) Not p.isIgnore).FirstOrDefault
            param.Likes("*")

            Dim table As DataTable = query.Invoke(connection, New List(Of SAPFieldItem) From {param})

            Dim columns = (From col As DataColumn In table.Columns).ToList
            Dim count As Integer = 0
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