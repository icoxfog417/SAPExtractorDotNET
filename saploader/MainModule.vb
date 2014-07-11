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
            Dim options As List(Of SAPFieldItem) = makeOptions(args)
            Dim result As DataTable = Nothing

            Dim connection As RfcDestination = connector.Login
            Console.WriteLine("Connect to SAP " + connector.Destination)

            'extract from SAP
            If isTable Then
                'extract from table

                Dim fields As List(Of SAPFieldItem) = makeFields(args)
                Dim tableLoader As New SAPTableExtractor(target)
                result = tableLoader.Invoke(connection, fields, options)
                Console.WriteLine("Extraction done from " + tableLoader.Table)

            Else
                'extract from query
                Dim queryInfos As String() = target.Split("/")
                If queryInfos.Count = 2 Then
                    Dim queryName As String = queryInfos(0)
                    Dim userGroup As String = queryInfos(1)
                    Dim queryLoader As New SAPQueryExtractor(queryName, userGroup)
                    queryLoader.QueryVariant = GetOptionValue(args, "/V")
                    result = queryLoader.Invoke(connection, options)
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

    Private Function makeFields(ByVal arguments As String()) As List(Of SAPFieldItem)
        Const FieldOption As String = "/F"
        Dim fields As New List(Of SAPFieldItem)
        Dim parameterForField As String = GetOptionValue(arguments, FieldOption)

        If parameterForField IsNot Nothing Then
            Dim columns As String() = parameterForField.Split(",")

            For Each column As String In columns
                fields.Add(New SAPFieldItem(column))
            Next

        End If

        Return fields

    End Function

    Private Function makeOptions(ByVal arguments As String()) As List(Of SAPFieldItem)
        Const OptionOption As String = "/O"
        Dim options As New List(Of SAPFieldItem)
        Dim parameterForOption As String = GetOptionValue(arguments, OptionOption)

        If parameterForOption IsNot Nothing Then
            Dim statements As String() = parameterForOption.Split(",")
            For i As Integer = 0 To statements.Count - 1
                'greater part of query fields are select-option 
                Dim f As SAPFieldItem = SAPFieldItem.createByStatement(statements(i), True)
                options.Add(f)
            Next
        End If

        Return options

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
        usages.Add("  from table : saploader <destination> <tablename> <filename> [/f <fields>] [/o <options>]")
        usages.Add("  from query : saploader <destination> <queryname/userGroup> <filename> /q [/o <options>] [/v <query variant>]")
        usages.Add("    if you want to set condition or order to query, create query variant in SAP at present.")
        usages.Add("  other options: ")
        usages.Add("    /h  : write header")
        usages.Add("    /hc : write header by caption text")
        usages.Add("examples:")
        usages.Add("  saploader MY_SAP T001 table.txt /t /f BUKRS,BUTXT /o BUKRS=1*")
        usages.Add("  saploader MY_SAP ZQUERY01/MYGROUP query.txt /q /v MY_VARIANT")
        usages.ForEach(Sub(u) Console.WriteLine(u))

    End Sub

End Module
