SAPExtractorDotNET
=============

This is connector for SAP system.  
You can extract data from Table or Query.  

## How to use
Below is code sample (see test project code).

### Extract from Table

```vbnet
Dim connector As New SAPConnector(TestDestination)

Try
    'Login to sap
    Dim connection As RfcDestination = connector.Login

    Dim tableExtractor As New SAPTableExtractor("T001")

    'Select columns
    Dim conditions As New List(Of SAPFieldItem)
    For Each column As String In {"BUKRS", "BUTXT"}
        conditions.Add(New SAPFieldItem(column))
    Next
    
    'Set parameters
    fields.Add(New SAPFieldItem("SPRAS").IsEqualTo(CultureInfo.CurrentCulture.TwoLetterISOLanguageName.Substring(0, 1).ToUpper))
    
    'Execute
    Dim table As DataTable = tableExtractor.Invoke(connection, conditions, fields)

Catch ex As Exception
    Console.WriteLine(ex.Message)
End Try

```

### Extract from SAP Query

```vbnet
Dim connector As New SAPConnector(TestDestination)

Try
    'Login to sap
    Dim connection As RfcDestination = connector.Login
    
    'Set query'name and usergroup
    Dim query As New SAPQueryExtractor(TestQuery, TestUserGroup)

    'Get SAP Query's input parameters
    Dim param As SAPFieldItem = query.GetSelectFields(connection).Where(Function(p) Not p.isIgnore).FirstOrDefault
    param.Likes("*") 'set parameter value
    
    'Execute Query
    Dim table As DataTable = query.Invoke(connection, New List(Of SAPFieldItem) From {param})

Catch ex As Exception
    Console.WriteLine(ex.Message)
End Try

```
