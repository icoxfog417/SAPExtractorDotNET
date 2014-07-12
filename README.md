SAPExtractorDotNET
=============

This is connector for SAP system.  
You can extract data from Table or Query.  

[API Documentation](http://icoxfog417.github.io/SAPExtractorDotNET/Index.html)  
[Install from Nuget](https://www.nuget.org/packages/SAPExtractorDotNET)  

## How to use
Below is code sample (see test project code).  
(repository is builded by x86, if your system is x64 then build it by SAPExtractorDotNET/SAPNco(x64))  
**[You can use it by command line! see saploader's README](https://github.com/icoxfog417/SAPExtractorDotNET/tree/master/saploader)**

### Extract from Table

C#

```csharp
SAPConnector connector = new SAPConnector(TestDestination);

try {
	RfcDestination connection = connector.Login;

	SAPTableExtractor tableExtractor = new SAPTableExtractor(TestTable);
	DataTable table = tableExtractor.Invoke(connection, {
		"BUKRS",
		"BUTXT"
	}, new SAPFieldItem("SPRAS").IsEqualTo("EN"));
	ResultWriter.Write(table);

} catch (Exception ex) {
	Console.WriteLine(ex.Message);
}
```

VB.Net

```vbnet
Dim connector As New SAPConnector(TestDestination)

Try
    Dim connection As RfcDestination = connector.Login

    Dim tableExtractor As New SAPTableExtractor(TestTable)
    Dim table As DataTable = tableExtractor.Invoke(connection, {"BUKRS", "BUTXT"}, New SAPFieldItem("SPRAS").IsEqualTo("EN"))
    ResultWriter.Write(table)

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
