Imports Microsoft.VisualBasic

Imports SAP.Middleware.Connector
Imports System.Threading.Tasks
Imports System.Data

Namespace SAPExtractorDotNET

    ''' <summary>
    ''' Extract data from SAP Query
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SAPQueryExtractor
        Private Const RFC_TO_GET_QUERY_CATALOG As String = "RSAQ_REMOTE_QUERY_CATALOG"
        Private Const RFC_TO_GET_QUERY_FIELD As String = "RSAQ_REMOTE_QUERY_FIELDLIST"
        Private Const RFC_TO_EXECUTE_QUERY As String = "RSAQ_REMOTE_QUERY_CALL"
        Private Const DEFAULT_QUERY_AREA As String = "G"

        Private _queryArea As String = DEFAULT_QUERY_AREA
        Public ReadOnly Property QueryArea As String
            Get
                Return _queryArea
            End Get
        End Property

        Private _query As String = ""
        Public ReadOnly Property Query As String
            Get
                Return _query
            End Get
        End Property

        Private _userGroup As String = ""
        Public ReadOnly Property UserGroup As String
            Get
                Return _userGroup
            End Get
        End Property

        Public Property QueryVariant As String = ""
        Public Property QueryText As String = ""

        Public Sub New(ByVal queryArea As String, ByVal query As String, ByVal userGroup As String)
            _queryArea = queryArea
            _query = query
            _userGroup = userGroup
        End Sub

        Public Sub New(ByVal query As String, ByVal userGroup As String)
            _query = query
            _userGroup = userGroup
        End Sub

        ''' <summary>
        ''' Find SAP Query
        ''' </summary>
        ''' <param name="destination"></param>
        ''' <param name="queryName"></param>
        ''' <param name="userGroup"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Find(ByVal destination As RfcDestination, Optional ByVal queryName As String = "", Optional ByVal userGroup As String = "") As List(Of SAPQueryExtractor)
            Dim result As New List(Of SAPQueryExtractor)
            Dim func As IRfcFunction = destination.Repository.CreateFunction(RFC_TO_GET_QUERY_CATALOG)
            func.SetValue("WORKSPACE", DEFAULT_QUERY_AREA)
            func.SetValue("GENERIC_QUERYNAME", If(String.IsNullOrEmpty(queryName), "*", queryName))
            func.SetValue("GENERIC_USERGROUP", If(String.IsNullOrEmpty(userGroup), "*", userGroup))

            func.Invoke(destination)

            Dim rfcTable As IRfcTable = func.GetTable("QUERYCATALOG")

            For i As Integer = 0 To rfcTable.Count - 1
                rfcTable.CurrentIndex = i
                Dim q As New SAPQueryExtractor(rfcTable.GetValue("QUERY"), rfcTable.GetValue("NUM"))
                q.QueryText = rfcTable.GetValue("QTEXT")
                result.Add(q)
            Next

            Return result

        End Function

        Public Function GetSelectFields(ByVal destination As RfcDestination) As List(Of SAPFieldItem)
            Dim fields As New List(Of SAPFieldItem)

            Dim func As IRfcFunction = destination.Repository.CreateFunction(RFC_TO_GET_QUERY_FIELD)
            func.SetValue("WORKSPACE", QueryArea)
            func.SetValue("QUERY", Query)
            func.SetValue("USERGROUP", UserGroup)

            func.Invoke(destination)

            Dim rfcTable As IRfcTable = func.GetTable("SEL_FIELDS")

            For i As Integer = 0 To rfcTable.Count - 1
                rfcTable.CurrentIndex = i
                Dim f As SAPFieldItem = SAPFieldItem.createByQueryStructure(rfcTable)
                f.Order = i
                fields.Add(f)
            Next

            Return fields

        End Function

        ''' <summary>
        ''' Extract data from SAP Query
        ''' </summary>
        ''' <param name="destination"></param>
        ''' <param name="fields"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Invoke(ByVal destination As RfcDestination, Optional ByVal fields As List(Of SAPFieldItem) = Nothing) As DataTable
            Dim result As New DataTable
            Dim filters As List(Of SAPFieldItem) = fields

            'no-variant and no-filters cause NO_SELECTION exception. so if input parameter exist , set * to it.
            If fields Is Nothing OrElse fields.Count = 0 Then
                If String.IsNullOrEmpty(QueryVariant) Then
                    Dim param As SAPFieldItem = GetSelectFields(destination).Where(Function(p) Not p.isIgnore).FirstOrDefault
                    If param IsNot Nothing Then
                        filters = New List(Of SAPFieldItem) From {param.Matches("*")}
                    End If
                End If
            Else
                For i As Integer = 0 To filters.Count - 1
                    If String.IsNullOrEmpty(filters(i).FieldId) Then
                        filters(i).FieldId = "SP$" + (i + 1).ToString.PadLeft(5, "0") 'set default parameter name
                    End If
                Next
            End If

            Dim func As IRfcFunction = destination.Repository.CreateFunction(RFC_TO_EXECUTE_QUERY)
            func.SetValue("WORKSPACE", QueryArea)
            func.SetValue("QUERY", Query)
            func.SetValue("USERGROUP", UserGroup)
            func.SetValue("DATA_TO_MEMORY", "X")
            func.SetValue("EXTERNAL_PRESENTATION", "Z") 'convert sap currency to external (without comma format)

            If Not String.IsNullOrEmpty(QueryVariant) Then
                func.SetValue("VARIANT", QueryVariant)
            End If

            If filters IsNot Nothing AndAlso filters.Count > 0 Then
                Dim selections As IRfcTable = func.GetTable("SELECTION_TABLE")
                For Each field As SAPFieldItem In filters
                    selections.Append()
                    selections.CurrentIndex = selections.Count - 1
                    selections.SetValue("SELNAME", field.FieldId)
                    selections.SetValue("KIND", If(field.IsRangeField, "S", "P"))
                    selections.SetValue("SIGN", If(field.IsExclude, "E", "I"))
                    selections.SetValue("OPTION", field.Operand)
                    selections.SetValue("LOW", field.Value)
                    If field.IsRangeField Then
                        selections.SetValue("HIGH", field.MaxValue)
                    End If
                Next
            End If

            func.Invoke(destination)

            'read type and make columns
            Dim struct As IRfcTable = func.GetTable("LISTDESC")
            For i As Long = 0 To struct.Count - 1
                struct.CurrentIndex = i
                Dim column As New DataColumn(struct.GetString("FNAME"))
                If result.Columns.Contains(column.ColumnName) Then
                    column.ColumnName = column.ColumnName + "__" + struct.GetString("LID") ' for duplicate type
                End If

                column.Caption = struct.GetString("FDESC")

                result.Columns.Add(column)

            Next

            'read data
            Dim lines As New List(Of LineData)
            Dim ldata As IRfcTable = func.GetTable("LDATA")
            For i As Long = 0 To ldata.Count - 1
                ldata.CurrentIndex = i
                Dim line As New LineData(i, ldata.GetString("LINE"))
                lines.Add(line)
            Next

            'returned data is like csv data. so split it
            Parallel.ForEach(lines, Sub(r As LineData) r.Split())

            Dim index As Integer = 0
            Dim bound As Integer = result.Columns.Count
            Dim rowValues As New List(Of String)
            'read each splited element into datarow
            For Each item As LineData In lines
                For Each element As String In item.Elements
                    If index < bound Then
                        rowValues.Add(element)
                    Else
                        Dim row As DataRow = result.NewRow()
                        For i As Integer = 0 To rowValues.Count - 1
                            row(i) = rowValues(i)
                        Next
                        result.Rows.Add(row)
                        rowValues.Clear()
                        rowValues.Add(element)
                        index = 0
                    End If
                    index += 1

                Next
            Next

            'add last row
            If rowValues.Count > 0 Then
                Dim lastRow As DataRow = result.NewRow()
                For i As Integer = 0 To rowValues.Count - 1
                    lastRow(i) = rowValues(i)
                Next
                result.Rows.Add(lastRow)
            End If

            Return result

        End Function

        ''' <summary>
        ''' to read rfc returned string
        ''' </summary>
        ''' <remarks></remarks>
        Private Class LineData
            Public Property Index As Long = 0
            Public Property Line As String = ""
            Public Property Elements As New List(Of String)

            Public Const LineSeparator As String = ";"
            Public Const ItemSeparator As String = ","
            Public Const FieldSeparator As String = ":"

            Public Sub New(ByVal index As Long, ByVal line As String)
                Me.Index = index
                Me.Line = line
            End Sub

            Public Sub Split()
                Dim separated As String() = Line.Split((LineSeparator + ItemSeparator).ToCharArray)
                Dim els As New List(Of String)

                For i As Integer = 0 To UBound(separated)
                    Dim str As String = separated(i)
                    If String.IsNullOrEmpty(str) Then Continue For
                    If isElement(str) Then
                        els.Add(str)
                    ElseIf isElement(els.Last + str) Then 'for data include separator
                        Dim concated As String = els.Last + ItemSeparator + str
                        els.RemoveAt(els.Count - 1)
                        els.Add(concated)
                    End If

                Next

                For Each e As String In els
                    Dim es As String() = e.Split(FieldSeparator.ToCharArray, 2)
                    Elements.Add(es(1))
                Next

            End Sub

            Private Function isElement(ByVal str As String) As Boolean
                Dim itemLength As Integer = 0
                'each element is like 001:xxxxx
                If str.IndexOf(FieldSeparator) = 3 AndAlso Integer.TryParse(str.Substring(0, 3), itemLength) Then
                    Return True
                Else
                    Return False
                End If
            End Function

        End Class


    End Class

End Namespace
