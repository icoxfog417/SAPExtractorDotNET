Imports SAPExtractorDotNET
Imports SAP.Middleware.Connector
Imports System.IO

Module MainModule

    Sub Main(args As String())

        Dim destination As String = ""
        Dim target As String = ""
        Dim filename As String = ""
        Dim isTable As Boolean = True

        If args.Length < 3 Then
            Console.WriteLine("You have to set obligate parameters (destination,target,filename)" + vbCrLf)
            showUsage()
        Else
            destination = args(0)
            target = args(1)
            filename = args(2)
        End If

        If IndexOfOption(args, "/Q") > -1 Then isTable = False


        Try

            'create connector to SAP
            Dim connector As New SAPConnector(destination)

            'get common options
            Dim filters As List(Of SAPFieldItem) = makeFilters(args)
            Dim result As DataTable = Nothing

            Dim connection As RfcDestination = connector.Login
            Console.WriteLine("Connect to SAP " + connector.Destination)

            'extract from SAP
            If isTable Then
                'extract from table

                Dim conditions As List(Of SAPFieldItem) = makeConditions(args)
                Dim order As String = GetOptionValue(args, "/O")

                Dim tableLoader As New SAPTableExtractor(target)
                result = tableLoader.Invoke(connection, conditions, filters, order)
                Console.WriteLine("Extraction done from " + tableLoader.Table)

            Else
                'extract from query
                Dim queryInfos As String() = target.Split("/")
                If queryInfos.Count = 2 Then
                    Dim queryName As String = queryInfos(0)
                    Dim userGroup As String = queryInfos(1)
                    Dim queryLoader As New SAPQueryExtractor(queryName, userGroup)
                    queryLoader.QueryVariant = GetOptionValue(args, "/V")
                    result = queryLoader.Invoke(connection, filters)
                    Console.WriteLine("Extraction done from " + queryLoader.Query + "/" + queryLoader.UserGroup)

                Else
                    Throw New Exception("You have to set query by queryName/userGroup")
                End If


            End If

            'write to file
            Dim writer As New StreamWriter(filename, False)
            Dim line As New List(Of String)
            If IndexOfOption(args, "/H") > -1 Or IndexOfOption(args, "/HC") > -1 Then
                For Each column As DataColumn In result.Columns
                    If IndexOfOption(args, "/H") > -1 Then
                        line.Add(column.ColumnName)
                    Else
                        line.Add(column.Caption)
                    End If
                Next
                writer.WriteLine(String.Join(vbTab, line))
            End If

            For Each row As DataRow In result.Rows
                line.Clear()
                For Each column As DataColumn In result.Columns
                    line.Add(row(column.ColumnName))
                Next
                writer.WriteLine(String.Join(vbTab, line))
            Next

            writer.Close()

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

    End Sub

    Private Function makeConditions(ByVal arguments As String()) As List(Of SAPFieldItem)
        Const ConditionOption As String = "/C"
        Dim conditions As New List(Of SAPFieldItem)
        Dim conditionValue As String = GetOptionValue(arguments, ConditionOption)

        If conditionValue IsNot Nothing Then
            Dim columns As String() = conditionValue.Split(",")

            For Each column As String In columns
                conditions.Add(New SAPFieldItem(column))
            Next

        End If

        Return conditions

    End Function

    Private Function makeFilters(ByVal arguments As String()) As List(Of SAPFieldItem)
        Const FilterOption As String = "/F"
        Dim filters As New List(Of SAPFieldItem)
        Dim filterValue As String = GetOptionValue(arguments, FilterOption)

        If filterValue IsNot Nothing Then
            Dim statements As String() = filterValue.Split(",")
            For i As Integer = 0 To statements.Count - 1
                'greater part of query fields are select-option 
                Dim f As SAPFieldItem = SAPFieldItem.createByStatement(statements(i), True)
                filters.Add(f)
            Next
        End If

        Return filters

    End Function

    Private Function GetOptionValue(ByVal arguments As String(), ByVal optionName As String) As String
        Dim optIndex As Integer = IndexOfOption(arguments, optionName)
        If optIndex > -1 Then
            If optIndex + 1 < arguments.Length Then
                Return arguments(optIndex + 1)
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function

    Private Function IndexOfOption(ByVal arguments As String(), ByVal optionName As String) As Integer
        Dim optIndex As Integer = -1
        For i As Integer = 0 To arguments.Count - 1
            If arguments(i).ToUpper = optionName.ToUpper Then
                optIndex = i
                Exit For
            End If
        Next

        Return optIndex
    End Function

    Sub showUsage()
        Dim usages As New List(Of String)
        usages.Add("saploader v" + Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString)
        usages.Add("usage:")
        usages.Add("  from table : saploader <destination> <tablename> <filename> [/c <columns>] [/f <filters>] [/o <order>]")
        usages.Add("  from query : saploader <destination> <queryname/userGroup> <filename> /q [/f <filters>] [/v <query variant>]")
        usages.Add("    if you want to set condition or order to query, create query variant in SAP at present.")
        usages.Add("  other options: ")
        usages.Add("    /h  : write header")
        usages.Add("    /hc : write header by caption text")
        usages.Add("examples:")
        usages.Add("  saploader MY_SAP T001 table.txt /t /c BUKRS,BUTXT /f BUKRS=C* /o BUKRS")
        usages.Add("  saploader MY_SAP ZQUERY01/MYGROUP query.txt /q /v MY_VARIANT")
        usages.ForEach(Sub(u) Console.WriteLine(u))

    End Sub

End Module
