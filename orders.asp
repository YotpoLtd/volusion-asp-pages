
<%
Option Explicit
Response.Buffer = True

'DO NOT MODIFY ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOUR DOING
Const CONFIGURATION_YOTPO_DESTINATION_PATH = "/v/vspfiles/schema/Generic"
Const CONFIGURATION_API_USE_SSL = False
Const VOLUSION_API_CALL_DEFAULT_USE_SSL = False
Const VOLUSION_API_CALL_RELATIVE_URL = "/net/WebService.aspx"
Const VOLUSION_API_CALL_QUERY_STRING_USER_NAME = "Login"
Const VOLUSION_API_CALL_QUERY_STRING_PASSWORD = "EncryptedPassword"
Const VOLUSION_API_CALL_QUERY_STRING_API_NAME = "API_Name"
Const VOLUSION_API_CALL_QUERY_STRING_SELECT = "SELECT_Columns"
Const VOLUSION_API_CALL_QUERY_STRING_WHERE_COLUMN = "WHERE_Column"
Const VOLUSION_API_CALL_QUERY_STRING_WHERE_VALUE = "WHERE_Value"
Const VOLUSION_API_CALL_QUERY_STRING_IMPORT = "Import"
Const VOLUSION_API_CALL_CONTENT_TYPE = "text/xml; charset=utf-8"
Const VOLUSION_API_CALL_CONTENT_ACTION = "Volusion_API"
Const VOLUSION_API_CALL_REQUEST_METHOD = "POST"

Private Function GetFile(ByVal Path) 'As Scripting.File
	Set GetFile = Nothing
	If FSO.FileExists(Path) Then
		Set GetFile = FSO.GetFile(Path)
	End If
End Function

Class VolusionAPICall
	Private LocalXMLHTTP 'As MSXML2.ServerXMLHTTP.3.0
	Private LocalDomainName 'As String
	Private LocalUserName 'As String
	Private LocalPassword 'As String
	Private LocalUseSSL 'As Boolean	
	Private LocalAPISchemaName 'As String
	Private LocalDestinationPath 'As String
	Private LocalFSO 'As Scripting.FileSystemObject	

	Private Sub Class_Initialize
		LocalUseSSL = VOLUSION_API_CALL_DEFAULT_USE_SSL
		LocalDestinationPath = Null
		LocalDomainName = Null
		LocalUserName = Null
		LocalPassword = Null
		LocalAPISchemaName = Null		
		Set LocalXMLHTTP = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
		Set LocalFSO = Server.CreateObject("Scripting.FileSystemObject")
	End Sub
	
	Private Sub Class_Terminate	
		Set LocalXMLHTTP = Nothing
		Set LocalFSO = Nothing
	End Sub

	Public Property Get DomainName()
		DomainName = LocalDomainName
	End Property

	Public Property Let DomainName(ByVal vDomainName)
		LocalDomainName = vDomainName
	End Property
	
	Public Property Let UserName(ByVal vUserName)
		LocalUserName = vUserName
	End Property
	
	Public Property Let Password(ByVal vPassword)
		LocalPassword = vPassword
	End Property
	
	Public Property Let UseSSL(ByVal vUseSSL)
		LocalUseSSL = vUseSSL
	End Property
	
	Public Property Let APISchemaName(ByVal vAPISchemaName)
		LocalAPISchemaName = vAPISchemaName
	End Property
	
	Public Property Let DestinationPath(ByVal vDestinationPath)
		LocalDestinationPath = vDestinationPath
	End Property
		
	Public Property Get ResponseText()
		ResponseText = LocalXMLHTTP.responseText
	End Property
	
	Public Property Get ResponseXML()
		Set ResponseXML = LocalXMLHTTP.responseXML
	End Property	
	
	Public Property Get XMLHTTP()
		Set XMLHTTP = LocalXMLHTTP
	End Property
	
	Public Property Get ResponseIsValid() 'As Boolean
		'Make sure the response is correct
		If Trim(LocalXMLHTTP.status) <> "200" Then
			ResponseIsValid = False
			Exit Property
		End If
	
		'Make sure at least some data is back
		If Not (0 < Len(Trim(LocalXMLHTTP.responseText))) Then
			ResponseIsValid = False
			Exit Property
		End If
		
		ResponseIsValid = True
	End Property
	
	Public Sub DoAPICall()
		Dim URL 'As String

		'Create the URL for the request
		If LocalUseSSL Then
			URL = "https://"
		Else
			URL = "http://"
		End If
		URL = URL & LocalDomainName
		URL = URL & VOLUSION_API_CALL_RELATIVE_URL
		URL = URL & "?"
		URL = URL & VOLUSION_API_CALL_QUERY_STRING_USER_NAME & "=" & Server.URLEncode(LocalUserName)
		URL = URL & "&" & VOLUSION_API_CALL_QUERY_STRING_PASSWORD & "=" & Server.URLEncode(LocalPassword)
		URL = URL & "&" & VOLUSION_API_CALL_QUERY_STRING_API_NAME & "=" & Server.URLEncode(LocalAPISchemaName)

		'NEXT LINE IS FOR TESTING
		'response.write URL
		
		'Open the object	
		Call LocalXMLHTTP.Open(VOLUSION_API_CALL_REQUEST_METHOD, URL, False)
		
		'Set some header values
		Call LocalXMLHTTP.setRequestHeader("Content-Type", VOLUSION_API_CALL_CONTENT_TYPE)
		Call LocalXMLHTTP.setRequestHeader("Content-Action", VOLUSION_API_CALL_CONTENT_ACTION)
		
		'Set the timeout variables in milliseconds
		Call LocalXMLHTTP.setTimeouts(0,60000,60000,60000)
		
		'Make the actual request
		Call LocalXMLHTTP.Send()
	End Sub

	Public Sub DoCustomAPICall(ByVal SQL, ByVal Schema)
		Dim TempAPIName 'As String
		Dim TempStream 'As Scripting.TextStream
		Dim SQLFileName 'As String
		Dim XSDFileName 'As String
		
		'Create the temporary schema name
		TempAPIName = LocalFSO.GetTempName()
		
		'Create XSD and SQL filesf
		SQLFileName = LocalFSO.BuildPath(LocalDestinationPath, TempAPIName & ".sql")
		XSDFileName = LocalFSO.BuildPath(LocalDestinationPath, TempAPIName & ".xsd")
		'Write out the SQL to the temporary file
		Set TempStream = LocalFSO.CreateTextFile(SQLFileName, True, False)
		Call TempStream.Write(SQL)
		Call TempStream.Close()
		Set TempStream = Nothing
		'Write out the SQL to the temporary file
		Set TempStream = LocalFSO.CreateTextFile(XSDFileName, True, False)
		Call TempStream.Write(Schema)
		Call TempStream.Close()
		Set TempStream = Nothing

		'Verify the files are available, if not exit
		If Not LocalFSO.FileExists(SQLFileName) Or Not LocalFSO.FileExists(XSDFileName) Then
			Exit Sub
		End If

		'Execute through API
		LocalAPISchemaName = "Generic/" & TempAPIName
		
		'Keep going if the call errors out so any temp files are deleted.
		On Error Resume Next
			Call DoAPICall()
		
			'Delete XSD and SQL files
		Call LocalFSO.DeleteFile(SQLFileName, True)
		Call LocalFSO.DeleteFile(XSDFileName, True)
		On Error GoTo 0
		
		If Err.number <> 0 Then
			Err.Raise Err.number, Err.Source, Err.description
		End If
		
	End Sub
	
	Public Function ReadFile(ByVal FileName) 'As String
		Dim Stream 'As Scripting.TextStream
	
		Set Stream = LocalFSO.OpenTextFile(FileName, 1, False, 0) 
		'1 = ForReading, 0 = Opens the file as ASCII
		ReadFile = Stream.ReadAll()
		Call Stream.Close()
		Set Stream = Nothing
	End Function
End Class


Class YQuery
	Private LocalVolusionAPICallObject 'As VolusionAPICall
	Private LocalYotpoQueryTemplateSQL 'As String
	Private LocalYotpoQueryTemplateXSD 'As String	
	Private LocalInstallPath 'As String
	Private LocalDestinationPath 'As String
	Private LocalFSO 'As Scripting.FileSystemObject
	
	Private Sub Class_Initialize()
		Set LocalFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		Set LocalVolusionAPICallObject = New VolusionAPICall
		LocalVolusionAPICallObject.UseSSL = False
		
		LocalInstallPath = Null

		LocalYotpoQueryTemplateSQL = "SELECT Orders.OrderID," _
				& " Orders.CustomerID," _
				& " Orders.OrderDate," _
				& " Orders.ShipDate," _
				& " Orders.LastModified," _
				& " Orders.OrderStatus," _
				& " OrderDetails.ProductCode," _
				& " OrderDetails.ProductID" _
				& " FROM Orders" _
				& " LEFT JOIN OrderDetails ON Orders.OrderID = OrderDetails.OrderID" _
				& " WHERE Orders.LastModified between '{StartDate}' and '{EndDate}'" _
				& " ORDER BY Orders.OrderDate DESC"		

		LocalYotpoQueryTemplateXSD = "<?xml version=""1.0"" encoding=""utf-8"" ?>" _
				& "<xs:schema id=""Order"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"" xmlns:msdata=""urn:schemas-microsoft-com:xml-msdata"" xmlns:msprop=""urn:schemas-microsoft-com:xml-msprop"">" _
				& "<xs:element name=""Order"" msdata:IsDataSet=""true"" msdata:UseCurrentLocale=""true"">" _
				& "<xs:complexType>" _
				& "<xs:choice minOccurs=""0"" maxOccurs=""unbounded"">" _
				& "<xs:element name=""Table"" msdata:IsDataSet=""true"" msdata:UseCurrentLocale=""true"">" _
				& "<xs:complexType>" _
				& "<xs:choice minOccurs=""0"" maxOccurs=""unbounded"">" _
				& "<xs:sequence>" _
				& "<xs:element name=""OrderID"" msprop:SqlDbType=""Int"" minOccurs=""1"" />" _
				& "<xs:element name=""CustomerID"" msprop:SqlDbType=""Int"" minOccurs=""1"" />" _
				& "<xs:element name=""OrderDate"" msprop:SqlDbType=""SmallDateTime"" minOccurs=""0"" />" _
				& "<xs:element name=""ShipDate"" msprop:SqlDbType=""SmallDateTime"" minOccurs=""0"" />" _
				& "<xs:element name=""LastModified"" msprop:SqlDbType=""SmallDateTime"" minOccurs=""0"" />" _
				& "<xs:element name=""OrderStatus"" msprop:maxLength=""255"" msprop:SqlDbType=""VarChar"" minOccurs=""1"" />" _
				& "<xs:element name=""ProductCode"" msprop:maxLength=""255"" msprop:SqlDbType=""VarChar"" minOccurs=""1"" />" _
				& "<xs:element name=""ProductID"" msprop:SqlDbType=""Int"" minOccurs=""1"" />" _
				& "</xs:sequence></xs:choice></xs:complexType></xs:element></xs:choice></xs:complexType></xs:element></xs:schema>" 			
	End Sub

	Private Sub Class_Terminate()
		Set LocalFSO = Nothing
		Set LocalVolusionAPICallObject = Nothing
	End Sub
	
	Public Property Let DestinationPath(ByVal vDestinationPath)
		LocalVolusionAPICallObject.DestinationPath = vDestinationPath
	End Property
		
	Public Property Let DomainName(ByVal vDomainName)
		LocalVolusionAPICallObject.DomainName = vDomainName
	End Property
	
	Public Property Let UserName(ByVal vUserName)
		LocalVolusionAPICallObject.UserName = vUserName
	End Property
	
	Public Property Let Password(ByVal vPassword)
		LocalVolusionAPICallObject.Password = vPassword
	End Property
	
	Public Property Let UseSSL(ByVal vUseSSL)
		LocalVolusionAPICallObject.UseSSL = vUseSSL
	End Property
	
	Public Function Retrieve(ByVal iso, ByVal Sdate, ByVal Edate) 'As YotpoQueryItem()
		Dim SQL 'As String
		Dim XSD 'As String		
		Dim TempSQLFileName 'As String
		Dim TempXSDFileName 'As String
		Dim TempSQLStream 'As Scripting.TextStream
		Dim SQLFileName 'As String
		Dim XSDFileName 'As String
		Dim Document 'As MSXML.Document
		
		'Set default return value
		Retrieve = Array()

		'Setup SQL using template		
		SQL = LocalYotpoQueryTemplateSQL
		SQL =  replace(SQL, "{DomainName}", LocalVolusionAPICallObject.DomainName)
		SQL =  replace(SQL, "{CurrencyISO}", iso)		
		SQL =  replace(SQL, "{StartDate}", Sdate)
		SQL =  replace(SQL, "{EndDate}", Edate)

		'Setup XSD using template
		XSD = LocalYotpoQueryTemplateXSD	
		
	    'return sql string with replaced text
	    'Response.Write SQL 
		
		'Make custom API call
		Call LocalVolusionAPICallObject.DoCustomAPICall( SQL , XSD )				

		'Parse results
		If Not LocalVolusionAPICallObject.ResponseIsValid Then
			Exit Function
		End If
		Set Document = LocalVolusionAPICallObject.ResponseXML

		'This is where we return the XML data via Document.xml
		'Remember to set Response.ContentType = "text/xml"
		Response.AddHeader "Content-Type", "text/xml;charset=UTF-8"
		Response.CodePage = 65001
		Response.CharSet = "UTF-8"
		Response.ContentType = "text/xml"
		Response.Write Document.xml
		Set Document = Nothing
	End Function
End Class


Dim SearchQueryItems
Dim YotpoQuery
Dim Login
Dim EncryptedPassword
Dim CUR
Dim StartDate
Dim EndDate

CUR = Request.QueryString("Currency")
StartDate = Request.QueryString("StartDate")
EndDate = Request.QueryString("EndDate")
Login = Request.QueryString("Login")
EncryptedPassword = Request.QueryString("EncryptedPassword")

'Clear the buffer
Response.Clear()

'Set the content type, this is the default usage
Response.ContentType = "application/json; charset=windows-1252"

'Cache headers, none for this
Response.CacheControl = "private, no-cache, no-cache=Set-Cookie, proxy-revalidate"
Call Response.AddHeader("Pragma", "no-cache")

'Check for query string error, if so just output empty error message
If StartDate = "" OR EndDate = "" Then
	Response.Write("NO DATA TO RETURN BASED ON QUERY STRING SUPPLIED")
	Response.End()
End If

If  LEN(CUR) <> 3 OR CUR = "" Then
	CUR = "usd"
End If

Set YotpoQuery = New YQuery
YotpoQuery.DestinationPath = Server.MapPath(CONFIGURATION_YOTPO_DESTINATION_PATH)
YotpoQuery.DomainName = Request.ServerVariables("HTTP_HOST")
YotpoQuery.UserName = Login
YotpoQuery.Password = EncryptedPassword
YotpoQuery.UseSSL = CONFIGURATION_API_USE_SSL

Call YotpoQuery.Retrieve( CUR , StartDate , EndDate )

'Reset some objects
Set YotpoQuery = Nothing

Response.Flush()
Response.End()
%>