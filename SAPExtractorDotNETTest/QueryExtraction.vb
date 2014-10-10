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
    Public Sub MakeFieldByStatement()

        Dim f1 As SAPFieldItem = SAPFieldItem.createByStatement("BUKRS = 1000")
        Assert.AreEqual("BUKRS", f1.FieldId)
        Assert.AreEqual("EQ", f1.Operand)
        Assert.AreEqual("1000", f1.Value)

    End Sub

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
            Dim table As DataTable = query.Invoke(connection)
            ResultWriter.Write(table)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Assert.Fail()
        End Try

    End Sub

    <TestMethod()>
    Public Sub ExtractQueryWithCriteria()

        Dim connector As New SAPConnector(TestDestination)

        Try
            Dim connection As RfcDestination = connector.Login

            Dim query As New SAPQueryExtractor(TestQuery, TestUserGroup)

            Dim param As SAPFieldItem = SAPFieldItem.createByStatement(">=1000/1/1", True)
            Dim table As DataTable = query.Invoke(connection, New List(Of SAPFieldItem) From {param})
            ResultWriter.Write(table)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Assert.Fail()
        End Try

    End Sub

    <TestMethod()>
    Public Sub LineDataSeparation()
        Dim queryData As New List(Of String)
        queryData.Add("000000000000000001")
        queryData.Add("20010101")
        queryData.Add("")
        queryData.Add("")

        queryData.Add("A00000000000000001")
        queryData.Add("2002/01/01:")
        queryData.Add("5:00000,")
        queryData.Add("XX")

        queryData.Add("B00000000000000001")
        queryData.Add("20030101")
        queryData.Add("025:aaa, 000: aaaa, 022: bbbb;") 'most complicated case
        queryData.Add("X")

        Dim queryLine As String = ""
        For Each el In queryData
            If Not String.IsNullOrEmpty(queryLine) Then
                queryLine += ","
            End If
            queryLine += el.Length.ToString.PadLeft(3, "0") + ":" + el
        Next
        queryLine += ";/"
        Console.WriteLine(queryLine)

        Dim line As New LineData(0, queryLine)
        line.Split()

        For i As Integer = 0 To line.Elements.Count - 1
            Console.WriteLine(line.Elements(i))
            Assert.AreEqual(queryData(i), line.Elements(i))
        Next

    End Sub

End Class