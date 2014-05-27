Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports SAP.Middleware.Connector
Imports SAPExtractorDotNET
Imports System.Configuration
Imports SAPExtractorDotNETTest.Util


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
    Public Sub FindQuery()
        Dim connector As New SAPConnector(TestDestination)

        Try
            Dim connection As RfcDestination = connector.Login
            Dim list As List(Of SAPQueryExtractor) = SAPQueryExtractor.Find(connection, "Y*", "Y*")

            For i As Integer = 0 To If(list.Count > 10, 10, list.Count) - 1
                Dim q = list(i)
                Console.WriteLine(q.UserGroup + "/" + q.Query + ":" + q.QueryText)
            Next

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Assert.Fail()
        End Try

    End Sub

    <TestMethod()>
    Public Sub GetQueryParameters()

        Dim connector As New SAPConnector(TestDestination)

        Try
            Dim connection As RfcDestination = connector.Login

            Dim query As New SAPQueryExtractor(TestQuery, TestUserGroup)
            Dim queryParams As List(Of SAPFieldItem) = query.GetSelectFields(connection)

            ResultWriter.Write(queryParams)

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
            ResultWriter.Write(table)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Assert.Fail()
        End Try

    End Sub

End Class