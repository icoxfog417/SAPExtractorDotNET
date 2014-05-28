Imports Microsoft.VisualBasic
Imports SAP.Middleware.Connector

Namespace SAPExtractorDotNET

    ''' <summary>
    ''' Select parameter to execute query
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SAPFieldItem

        ''' <summary>Unique id for field</summary>
        Public Property FieldId As String

        ''' <summary>Structure type of field</summary>
        Public Property FieldStructure As String

        ''' <summary>is field for range selection</summary>
        Public Property IsRangeField As Boolean = False

        ''' <summary>is system field or not</summary>
        Public Property isIgnore As Boolean = False

        ''' <summary>Text for field</summary>
        Public Property FieldText As String

        ''' <summary>Order no of field</summary>
        Public Property Order As Integer

        ''' <summary>Field type</summary>
        Public Property FieldType As String = "C"

        ''' <summary>Field size</summary>
        Public Property FieldSize As Integer

        ''' <summary>the group of field</summary>
        Public Property FieldGroup As String

        ''' <summary>is key field or not</summary>
        Public Property isKey As Boolean = False

        Public Property Operand As String = "EQ"
        Public Property IsExclude As Boolean = False
        Public Property Value As String = ""
        Public Property MaxValue As String = ""

        Public Sub New()
        End Sub

        Public Sub New(ByVal fieldId As String, Optional ByVal isRangeField As Boolean = False)
            Me.FieldId = fieldId
            Me.IsRangeField = isRangeField
        End Sub

        Public Shared Function createByStatement(ByVal statement As String, Optional ByVal isRangeField As Boolean = False) As SAPFieldItem
            Dim operands As String() = {"<>", ">=", "<=", "><", ">", "<", "="}
            Dim f As SAPFieldItem = Nothing

            For Each opr As String In operands
                If statement.IndexOf(opr) > -1 Then
                    Dim keyValue As String() = statement.Split(opr)
                    If keyValue.Count = 2 Then
                        f = New SAPFieldItem(Trim(keyValue(0)), isRangeField)
                        f.ComparesBy(opr, Trim(keyValue(1)))
                    End If
                End If
            Next

            Return f

        End Function

        Public Shared Function createByQueryStructure(ByVal sapStructure As IRfcTable) As SAPFieldItem
            Dim f As New SAPFieldItem()
            f.FieldId = sapStructure.GetString("SPNAME")
            If f.FieldId.StartsWith("%") Then f.isIgnore = True

            f.FieldStructure = sapStructure.GetString("FNAME")
            f.IsRangeField = If(sapStructure.GetString("KIND") = "S", True, False)
            f.FieldText = sapStructure.GetString("FTEXT")

            f.FieldType = sapStructure.GetString("TYPE")
            f.FieldSize = CInt(sapStructure.GetString("LENGTH"))
            f.FieldGroup = sapStructure.GetString("RGROUP")

            Return f

        End Function

        Public Shared Function createByTableStructure(ByVal columns As Dictionary(Of String, String)) As SAPFieldItem
            Dim f As New SAPFieldItem()
            Dim getFieldValue = Function(name As String) As String
                                    If columns.ContainsKey(name) Then
                                        Return columns(name)
                                    Else
                                        Return Nothing
                                    End If
                                End Function

            f.FieldId = getFieldValue("FIELDNAME")
            f.FieldStructure = getFieldValue("ROLLNAME")
            f.FieldType = getFieldValue("DATATYPE")
            f.FieldSize = getFieldValue("LENG")
            f.Order = getFieldValue("POSITION")
            f.isKey = If(String.IsNullOrEmpty(getFieldValue("KEYFLAG")), False, True)
            f.FieldText = getFieldValue("DDTEXT")

            Return f

        End Function

        Public Function IsEqualTo(ByVal value As String) As SAPFieldItem
            Me.Operand = "EQ"
            Me.Value = value
            Return Me
        End Function

        Public Function Likes(ByVal value As String) As SAPFieldItem
            Me.Operand = "CP"
            Me.Value = "*" + value + "*"
            Return Me
        End Function

        Public Function StartsWith(ByVal value As String) As SAPFieldItem
            Me.Operand = "CP"
            Me.Value = value + "*"
            Return Me
        End Function

        Public Function EndsWith(ByVal value As String) As SAPFieldItem
            Me.Operand = "CP"
            Me.Value = "*" + value
            Return Me
        End Function

        Public Function Matches(ByVal value As String) As SAPFieldItem
            Me.Operand = "CP"
            Me.Value = value
            Return Me
        End Function

        Public Function IsNotEqualTo(ByVal value As String) As SAPFieldItem
            Me.Operand = "NE"
            Me.Value = value
            Return Me
        End Function

        Public Function GreaterThan(ByVal value As String) As SAPFieldItem
            Me.Operand = "GT"
            Me.Value = value
            Return Me
        End Function

        Public Function LowerThan(ByVal value As String) As SAPFieldItem
            Me.Operand = "LT"
            Me.Value = value
            Return Me
        End Function

        Public Function GreaterEqual(ByVal value As String) As SAPFieldItem
            Me.Operand = "GE"
            Me.Value = value
            Return Me
        End Function

        Public Function LowerEqual(ByVal value As String) As SAPFieldItem
            Me.Operand = "LE"
            Me.Value = value
            Return Me
        End Function

        Public Function Between(ByVal value As String, ByVal maxValue As String) As SAPFieldItem
            If Not IsRangeField Then Throw New Exception("You can use between only when field is RangeField.")
            Me.Operand = "BT"
            Me.Value = value
            Me.MaxValue = maxValue
            Return Me
        End Function

        Public Function ComparesBy(ByVal opr As String, value As String) As SAPFieldItem
            Select Case opr.Trim
                Case "="
                    If value.Contains("*") Then
                        Me.Matches(value)
                    Else
                        Me.IsEqualTo(value)
                    End If
                Case "<>"
                    Me.IsNotEqualTo(value)
                Case ">"
                    Me.GreaterThan(value)
                Case "<"
                    Me.LowerThan(value)
                Case ">="
                    Me.GreaterEqual(value)
                Case "<="
                    Me.LowerEqual(value)
                Case "><"
                    Dim values As String() = value.Split(",")
                    Me.Between(values(0), values(1))
            End Select

            Return Me

        End Function

        Public Function makeWhere() As String
            Dim where As String = FieldId
            Dim wValue As String = makeWhereValue(Value)
            Select Case Operand
                Case "EQ"
                    where += " = " + wValue
                Case "CP"
                    where += " LIKE " + wValue
                Case "NE"
                    where += " <> " + wValue
                Case "GT"
                    where += " > " + wValue
                Case "LT"
                    where += " < " + wValue
                Case "GE"
                    where += " >= " + wValue
                Case "LE"
                    where += " <= " + wValue
                Case "BT"
                    where += " BETWEEN " + wValue + " AND " + makeWhereValue(MaxValue)
            End Select

            Return where

        End Function

        Private Function makeWhereValue(ByVal value As String) As String
            Dim result As String = escape(value)
            If FieldType = "C" Or FieldType = "CHAR" Then
                result = "'" + result + "'" 'round by single quote
            End If
            Return result
        End Function

        Public Shared Function escape(ByVal value As String) As String
            Dim escaped As String = value
            escaped = escaped.Replace(";", "").Replace("\", "\\").Replace("'", "''")
            escaped = escaped.Replace("%", "").Replace("*", "%")
            Return escaped
        End Function

    End Class

End Namespace
