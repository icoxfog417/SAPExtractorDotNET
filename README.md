SAPExtractorDotNET
=============

This is connector for SAP system.  
You can extract data from Table or Query.  

[API Documentation](http://icoxfog417.github.io/SAPExtractorDotNET/Index.html)  
[Install from Nuget](https://www.nuget.org/packages/SAPExtractorDotNET)  

## How to use
Below is code sample (see test project code).  
(repository is builded by x86, if your system is x64 then build it by SAPExtractorDotNET/SAPNco(x64))

### Extract from Table

C#

```csharp
SAPConnector connector = new SAPConnector(TestDestination);

try {
	//Login to sap
	RfcDestination connection = connector.Login;

	SAPTableExtractor tableExtractor = new SAPTableExtractor("T001");

	//Select columns
	List<SAPFieldItem> conditions = new List<SAPFieldItem>();
	foreach (string column in {
		"BUKRS",
		"BUTXT"
	}) {
		conditions.Add(new SAPFieldItem(column));
	}

	//Set parameters
	fields.Add(new SAPFieldItem("SPRAS").IsEqualTo(CultureInfo.CurrentCulture.TwoLetterISOLanguageName.Substring(0, 1).ToUpper));

	//Execute
	DataTable table = tableExtractor.Invoke(connection, conditions, fields);

} catch (Exception ex) {
	Console.WriteLine(ex.Message);
}
```

VB.Net

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

C#

```csharp
SAPConnector connector = new SAPConnector(TestDestination);

try {
	//Login to sap
	RfcDestination connection = connector.Login;

	//Set query'name and usergroup
	SAPQueryExtractor query = new SAPQueryExtractor(TestQuery, TestUserGroup);

	//Get SAP Query's input parameters
	SAPFieldItem param = query.GetSelectFields(connection).Where(p => !p.isIgnore).FirstOrDefault;
	param.Likes("*");
	//set parameter value

	//Execute Query
	DataTable table = query.Invoke(connection, new List<SAPFieldItem> { param });

} catch (Exception ex) {
	Console.WriteLine(ex.Message);
}
```

VB.Net

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
