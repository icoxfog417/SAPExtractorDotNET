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
        'Private Const RFC_TO_GET_TABLE As String = "RFC_READ_TABLE"
        Private Const RFC_TO_GET_TABLE As String = "/SDF/GET_GENERIC_APO_TABLE"

        Private Const COLUMN_DEFINE_TABLE As String = "DD03M"

        Private Const MaxStatementCount As Integer = 12
        Private Const BasicEscapes As String = ";'\"

        Private _tableName As String = ""
        Public ReadOnly Property TableName As String
            Get
                Return _tableName
            End Get
        End Property

        Public Sub New(ByVal tableName As String)
            Me._tableName = tableName
        End Sub

        ''' <summary>
        ''' Get table's columns definition. 
        ''' </summary>
        ''' <param name="destination"></param>
        ''' <returns></returns>
        ''' <remarks>Column definitions are taken from DD03M</remarks>
        Public Function GetColumnFields(ByVal destination As RfcDestination) As List(Of SAPFieldItem)

            Dim conditions As New List(Of SAPFieldItem)
            Dim fields As New List(Of SAPFieldItem)

            conditions.Add(New SAPFieldItem("FIELDNAME"))
            conditions.Add(New SAPFieldItem("ROLLNAME"))
            conditions.Add(New SAPFieldItem("DATATYPE"))
            conditions.Add(New SAPFieldItem("LENG"))
            conditions.Add(New SAPFieldItem("POSITION"))
            conditions.Add(New SAPFieldItem("KEYFLAG"))
            conditions.Add(New SAPFieldItem("DDTEXT"))

            fields.Add(New SAPFieldItem("TABNAME").IsEqualTo(TableName))
            fields.Add(New SAPFieldItem("FLDSTAT").IsEqualTo("A"))
            fields.Add(New SAPFieldItem("DDLANGUAGE").IsEqualTo(CultureInfo.CurrentCulture.TwoLetterISOLanguageName.Substring(0, 1).ToUpper))

            Dim columnDefines As List(Of Dictionary(Of String, String)) = Invoke(destination, COLUMN_DEFINE_TABLE, conditions, fields, "POSITION")


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
        ''' <param name="conditions"></param>
        ''' <param name="fields"></param>
        ''' <param name="orders">usual order by string(MANDT , BUKRS DESC ... etc.)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Invoke(ByVal destination As RfcDestination, ByVal conditions As List(Of SAPFieldItem), ByVal fields As List(Of SAPFieldItem), Optional ByVal orders As String = "") As DataTable

            Dim result As New DataTable
            Dim rows As List(Of Dictionary(Of String, String)) = Nothing
            Dim localConditions As List(Of SAPFieldItem) = conditions

            If conditions Is Nothing Then
                'get all columns
                localConditions = GetColumnFields(destination)
            End If

            'RFC has byte-length restriction of row. So if there are too many columns, splite it and merge result after.
            If localConditions.Count > MaxStatementCount Then

                Dim columnDefines As List(Of SAPFieldItem) = GetColumnFields(destination)
                Dim keys As List(Of SAPFieldItem) = columnDefines.Where(Function(c) c.isKey).ToList
                keys.ForEach(Sub(k) k.isIgnore = True)
                Dim makeKey = Function(row As Dictionary(Of String, String)) As String
                                  Dim localKeies As New List(Of String)
                                  keys.ForEach(Sub(k) localKeies.Add(row(k.FieldId)))
                                  Return String.Join("__", localKeies)
                              End Function

                'split columns and merge it by primary after.
                Dim splitedCondition As New List(Of List(Of SAPFieldItem))
                Dim size As Integer = MaxStatementCount - keys.Count
                Dim limit As Integer = Math.Ceiling(localConditions.Count / size)

                For i As Integer = 0 To limit - 1
                    Dim splited As New List(Of SAPFieldItem)
                    keys.ForEach(Sub(k) splited.Add(k))
                    splited.AddRange(localConditions.Skip(i * size).Take(size).ToList)
                    splitedCondition.Add(splited)
                Next

                Dim tasks As New List(Of Task(Of Dictionary(Of String, Dictionary(Of String, String))))
                splitedCondition.ForEach(Sub(sc)
                                             tasks.Add(New Task(Of Dictionary(Of String, Dictionary(Of String, String)))(
                                                                Function()
                                                                    Dim tResult As List(Of Dictionary(Of String, String)) = Invoke(destination, TableName, sc, fields)
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
                rows = Invoke(destination, TableName, localConditions, fields, orders)

            End If

            For Each cond As SAPFieldItem In localConditions
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

        Private Function Invoke(ByVal destination As RfcDestination, ByVal table As String, ByVal conditions As List(Of SAPFieldItem), ByVal fields As List(Of SAPFieldItem), Optional ByVal orders As String = "") As List(Of Dictionary(Of String, String))

            If conditions Is Nothing OrElse conditions.Count = 0 Then
                Throw New Exception("You have to set more than 1 condition ")
            End If
            If conditions.Count > MaxStatementCount Or (fields IsNot Nothing AndAlso fields.Count > MaxStatementCount) Then
                Throw New Exception("The maximum count of condition / field is " + MaxStatementCount.ToString)
            End If

            Dim parameterLog As New Dictionary(Of String, String)
            Dim func As IRfcFunction = destination.Repository.CreateFunction(RFC_TO_GET_TABLE)
            For i As Integer = 1 To conditions.Count
                parameterLog.Add("SELECT_CONDITION" + i.ToString, If(i > 1, ",", "") + conditions(i - 1).FieldId)
                func.SetValue(parameterLog.Last.Key, parameterLog.Last.Value)
            Next

            If fields IsNot Nothing Then
                For i As Integer = 1 To fields.Count
                    parameterLog.Add("WHERE_CLAUSE" + i.ToString, If(i > 1, "AND ", "") + fields(i - 1).makeWhere)
                    func.SetValue(parameterLog.Last.Key, parameterLog.Last.Value)
                Next
            End If

            parameterLog.Add("FROM_TABLE", SAPFieldItem.escape(table))
            func.SetValue(parameterLog.Last.Key, parameterLog.Last.Value)

            If Not String.IsNullOrEmpty(orders) Then
                parameterLog.Add("ORDER_BY", SAPFieldItem.escape(orders))
                func.SetValue(parameterLog.Last.Key, parameterLog.Last.Value)
            End If

            func.Invoke(destination)

            Dim result As New List(Of Dictionary(Of String, String))

            Dim index As Integer = 0
            Dim lines As New List(Of String)
            Dim struct As IRfcTable = func.GetTable("LCSYSTAB")
            For i As Long = 0 To struct.Count - 1
                struct.CurrentIndex = i
                lines.Add(struct.GetString("ZEILE"))
                If lines.Count = conditions.Count Then
                    Dim row As New Dictionary(Of String, String)
                    For lc As Integer = 0 To lines.Count - 1
                        If Not row.ContainsKey(conditions(lc).FieldId) Then row.Add(conditions(lc).FieldId, lines(lc))
                    Next
                    result.Add(row)
                    lines.Clear()
                    index += 1
                End If
            Next

            If lines.Count > 0 Then ' add last line
                Dim row As New Dictionary(Of String, String)
                For lc As Integer = 0 To lines.Count - 1
                    If Not row.ContainsKey(conditions(lc).FieldId) Then row.Add(conditions(lc).FieldId, lines(lc))
                Next
                result.Add(row)
            End If

            Return result

        End Function

        Public Function Invoke(ByVal destination As RfcDestination, Optional ByVal orders As String = "") As DataTable
            Return Invoke(destination, Nothing, Nothing, orders)
        End Function

        Public Function Invoke(ByVal destination As RfcDestination, ByVal fields As List(Of SAPFieldItem), Optional ByVal orders As String = "") As DataTable
            Return Invoke(destination, Nothing, fields, orders)
        End Function

    End Class

End Namespace


