Imports Microsoft.VisualBasic

Imports SAP.Middleware.Connector
Imports System.Threading.Tasks
Imports System.Data
Imports System.Globalization

Namespace SAPExtractorDotNET

    ''' <summary>
    ''' Extract data from SAP Table
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SAPTableExtractor

        'RFC_READ_TABLE and BBP_RFC_READ_TABLE has some bug (mainly when extracting data from huge or including number fields table). 
        'SAP disabled /SDF/GET_GENERIC_APO_TABLE ...
        'Private Const RFC_TO_GET_TABLE As String = "/SDF/GET_GENERIC_APO_TABLE"

        Private Const RFC_TO_GET_TABLE As String = "RFC_READ_TABLE"

        Private Const TABLE_TABLE As String = "DD02T"
        Private Const COLUMN_DEFINE_TABLE As String = "DD03M"

        Private Const MaxRowByteLength As Integer = 512
        Private Const BasicEscapes As String = ";'\"

        Private _table As String = ""
        Public ReadOnly Property Table As String
            Get
                Return _table
            End Get
        End Property

        Public Property TableText As String

        Public Sub New(ByVal tableName As String)
            Me._table = tableName
        End Sub

        ''' <summary>
        ''' Find Table
        ''' </summary>
        ''' <param name="destination"></param>
        ''' <param name="tableName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Find(ByVal destination As RfcDestination, Optional ByVal tableName As String = "") As List(Of SAPTableExtractor)
            Dim result As New List(Of SAPTableExtractor)

            Dim fields As New List(Of SAPFieldItem)
            For Each columnName As String In {"TABNAME", "DDTEXT"}
                fields.Add(New SAPFieldItem(columnName))
            Next

            Dim options As New List(Of SAPFieldItem)
            options.Add(New SAPFieldItem("DDLANGUAGE").IsEqualTo(CultureInfo.CurrentCulture.TwoLetterISOLanguageName.Substring(0, 1).ToUpper))
            options.Add(New SAPFieldItem("AS4LOCAL").IsEqualTo("A"))
            If Not String.IsNullOrEmpty(tableName) Then
                If tableName.Contains("*") Then
                    options.Add(New SAPFieldItem("TABNAME").Matches(tableName))
                Else
                    options.Add(New SAPFieldItem("TABNAME").IsEqualTo(tableName))
                End If
            End If

            Dim tables As List(Of Dictionary(Of String, String)) = SAPTableExtractor.Invoke(destination, TABLE_TABLE, fields, options)
            tables = tables.OrderBy(Function(d) d("TABNAME")).ToList

            For Each table As Dictionary(Of String, String) In tables
                Dim t As New SAPTableExtractor(table("TABNAME"))
                t.TableText = table("DDTEXT")
                result.Add(t)
            Next

            Return result

        End Function

        ''' <summary>
        ''' Get table's columns definition. 
        ''' </summary>
        ''' <param name="destination"></param>
        ''' <returns></returns>
        ''' <remarks>Column definitions are taken from DD03M</remarks>
        Public Function GetColumnFields(ByVal destination As RfcDestination) As List(Of SAPFieldItem)

            Dim fields As New List(Of SAPFieldItem)
            Dim options As New List(Of SAPFieldItem)

            fields.Add(New SAPFieldItem("FIELDNAME"))
            fields.Add(New SAPFieldItem("ROLLNAME"))
            fields.Add(New SAPFieldItem("DATATYPE"))
            fields.Add(New SAPFieldItem("LENG"))
            fields.Add(New SAPFieldItem("POSITION"))
            fields.Add(New SAPFieldItem("KEYFLAG"))
            fields.Add(New SAPFieldItem("DDTEXT"))

            options.Add(New SAPFieldItem("TABNAME").IsEqualTo(Table))
            options.Add(New SAPFieldItem("FLDSTAT").IsEqualTo("A"))
            options.Add(New SAPFieldItem("DDLANGUAGE").IsEqualTo(CultureInfo.CurrentCulture.TwoLetterISOLanguageName.Substring(0, 1).ToUpper))

            Dim columnDefines As List(Of Dictionary(Of String, String)) = Invoke(destination, COLUMN_DEFINE_TABLE, fields, options)
            columnDefines = columnDefines.OrderBy(Function(d) d("POSITION")).ToList

            Dim result As New List(Of SAPFieldItem)
            For Each cDefine As Dictionary(Of String, String) In columnDefines
                Dim f As SAPFieldItem = SAPFieldItem.createByTableStructure(cDefine)
                result.Add(f)
            Next

            Return result

        End Function

        ''' <summary>
        ''' Get SAP Table data.<br/>
        ''' If condition is Nothing, then get all columns of table.
        ''' </summary>
        ''' <param name="destination"></param>
        ''' <param name="fields"></param>
        ''' <param name="options"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Invoke(ByVal destination As RfcDestination, ByVal fields As List(Of SAPFieldItem), ByVal options As List(Of SAPFieldItem)) As DataTable

            Dim result As New DataTable
            Dim rows As List(Of Dictionary(Of String, String)) = Nothing
            Dim localFields As List(Of SAPFieldItem) = fields
            Dim rowByteLength As Integer = 0
            If fields Is Nothing OrElse fields.Count = 0 Then
                'get all columns
                localFields = GetColumnFields(destination)
                localFields.ForEach(Sub(f) rowByteLength += f.FieldSize)
            End If

            'RFC has byte-length restriction of row. So if there are too many columns, splite it and merge result after.
            If rowByteLength > MaxRowByteLength Then

                Dim columnDefines As List(Of SAPFieldItem) = GetColumnFields(destination)
                Dim keys As List(Of SAPFieldItem) = columnDefines.Where(Function(c) c.isKey).ToList
                Dim keySize As Integer = 0
                keys.ForEach(Sub(k) k.isIgnore = True)
                keys.ForEach(Sub(k) keySize += k.FieldSize)

                'make row's primary key string
                Dim makeKey = Function(row As Dictionary(Of String, String)) As String
                                  Dim localKeies As New List(Of String)
                                  keys.ForEach(Sub(k) localKeies.Add(row(k.FieldId)))
                                  Return String.Join("__", localKeies)
                              End Function

                'split columns and merge it by primary after.
                Dim splitedFields As New List(Of List(Of SAPFieldItem))
                Dim limit As Integer = MaxRowByteLength - keySize
                Dim splitedTotal As Integer = 0
                Dim splited As New List(Of SAPFieldItem)

                For Each columns In columnDefines

                    If splitedTotal + columns.FieldSize > limit Then
                        Dim fieldSet As New List(Of SAPFieldItem)(keys)
                        fieldSet.AddRange(splited)
                        splitedFields.Add(fieldSet)
                        splited.Clear()
                        splitedTotal = 0
                    End If

                    If Not columns.isIgnore Then
                        splited.Add(columns)
                        splitedTotal += columns.FieldSize
                    End If

                Next

                If splited.Count > 0 Then
                    Dim fieldSet As New List(Of SAPFieldItem)(keys)
                    fieldSet.AddRange(splited)
                    splitedFields.Add(fieldSet)
                End If

                Dim tasks As New List(Of Task(Of Dictionary(Of String, Dictionary(Of String, String))))
                splitedFields.ForEach(Sub(sc)
                                          tasks.Add(New Task(Of Dictionary(Of String, Dictionary(Of String, String)))(
                                                             Function()
                                                                 Dim tResult As List(Of Dictionary(Of String, String)) = Invoke(destination, Table, sc, options)
                                                                 Return tResult.ToDictionary(Of String, Dictionary(Of String, String))(Function(tr) makeKey(tr), Function(tr) tr)
                                                             End Function
                                                             ))
                                      End Sub)

                Dim taskRuns = tasks.ToArray
                Array.ForEach(taskRuns, Sub(t) t.Start())
                task.WaitAll(taskRuns)

                Dim merged As New Dictionary(Of String, Dictionary(Of String, String))

                'merge by primary key
                For Each runed In taskRuns
                    If runed.IsCompleted Then
                        For Each tResult In runed.Result
                            If Not merged.ContainsKey(tResult.Key) Then
                                merged.Add(tResult.Key, tResult.Value)
                            Else
                                For Each item In tResult.Value
                                    If Not merged(tResult.Key).ContainsKey(item.Key) Then merged(tResult.Key).Add(item.Key, item.Value)
                                Next
                            End If
                        Next
                    End If
                Next

                rows = merged.Values.ToList

            Else
                rows = Invoke(destination, Table, localFields, options)

            End If

            For Each cond As SAPFieldItem In localFields
                Dim column As New DataColumn(cond.FieldId)
                column.Caption = cond.FieldText
                result.Columns.Add(column)
            Next

            For Each row In rows
                Dim addRow As DataRow = result.NewRow
                For Each column As DataColumn In result.Columns
                    If row.ContainsKey(column.ColumnName) Then
                        addRow(column.ColumnName) = row(column.ColumnName)
                    End If
                Next
                result.Rows.Add(addRow)
            Next

            Return result

        End Function

        Private Shared Function Invoke(ByVal destination As RfcDestination, ByVal table As String, ByVal fields As List(Of SAPFieldItem), ByVal options As List(Of SAPFieldItem))

            Dim func As IRfcFunction = destination.Repository.CreateFunction(RFC_TO_GET_TABLE)

            'set fields
            If fields IsNot Nothing Then
                Dim fs As IRfcTable = func.GetTable("FIELDS")
                For i As Integer = 0 To fields.Count - 1
                    fs.Append()
                    fs.CurrentIndex = i
                    fs.SetValue("FIELDNAME", fields(i).FieldId)
                Next
            End If

            'set options
            If options IsNot Nothing Then
                Dim opts As IRfcTable = func.GetTable("OPTIONS")
                For i As Integer = 0 To options.Count - 1
                    opts.Append()
                    opts.CurrentIndex = i
                    opts.SetValue("TEXT", If(i > 0, " AND ", "") + options(i).makeWhere)
                Next
            End If

            func.SetValue("QUERY_TABLE", SAPFieldItem.escape(table))

            func.Invoke(destination)

            Dim result As New List(Of Dictionary(Of String, String))

            Dim index As Integer = 0
            Dim struct As IRfcTable = func.GetTable("FIELDS")
            Dim columns As New Dictionary(Of String, Integer)

            For i As Integer = 0 To struct.Count - 1
                struct.CurrentIndex = i
                columns.Add(struct.GetString("FIELDNAME"), struct.GetInt("LENGTH"))
            Next

            Dim data As IRfcTable = func.GetTable("DATA")
            For i As Long = 0 To data.Count - 1
                data.CurrentIndex = i
                Dim line = data.GetString("WA")

                Dim row As New Dictionary(Of String, String)
                Dim position As Integer = 0
                For Each c In columns
                    If position + c.Value <= line.Length Then
                        row.Add(c.Key, line.Substring(position, c.Value).Trim)
                    ElseIf position <= line.Length Then
                        row.Add(c.Key, line.Substring(position).Trim)
                    Else
                        row.Add(c.Key, String.Empty)
                    End If
                    position += c.Value
                Next
                result.Add(row)
            Next

            Return result

        End Function

        Public Function Invoke(ByVal destination As RfcDestination) As DataTable
            Return Invoke(destination, Nothing, Nothing)
        End Function

        Public Function Invoke(ByVal destination As RfcDestination, fields As String()) As DataTable
            Return Invoke(destination, New List(Of String)(fields))
        End Function

        Public Function Invoke(ByVal destination As RfcDestination, fields As List(Of String)) As DataTable
            Dim fieldItems As New List(Of SAPFieldItem)

            For Each field As String In fields
                fieldItems.Add(New SAPFieldItem(field))
            Next

            Return Invoke(destination, fieldItems, Nothing)

        End Function

    End Class

End Namespace


