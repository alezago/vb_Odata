<div id="top"></div>

<br />
<div align="center">

<h3 align="center">VB-OData</h3>


  <p align="center">
    A VBA Generic OData Client
  </p>

</div>

## About The Project

This library was created to provide a simple way to integrate Excel as an OData client for any web service supporting the OData Protocol.

It supports all HTTP Methods (GET, PUT, POST, DELETE, PATCH, HEAD), and allows for easy integration with any API supporting OData. <br>
The library supports the OAuth2 Client-Credentials Authentication flow and takes care of all steps necessary for authenticating the client to the Service Provider, with no configuration necessary from the user.

The package also includes the following resources which can be used as standalone:
<ls>
<li><b>libHTTP:</b> a generic HTTP request implementation, can be used directly to perform HTTP requests for services which do not use the OData protocol (REST, etc)</li>
<li><b>libJSON:</b> a simple JSON parsing library, useful to process the values returned from the HTTP requests.</li>
</ls>

<p align="right">(<a href="#top">back to top</a>)</p>

### Installation

1. Import the `OAuthClient.cls`, `libHTTP.bas`, `libJSON.bas`, `utilities.bas` modules into the VBA Project.<br>
2. Enable the Reference to `Microsoft XML, v6.0` and `Microsoft Visual Basic for Applications Extensibility 5.3` from the Tools>References menu.

<p align="right">(<a href="#top">back to top</a>)</p>

### Performing OData Requests

To perform an (Authenticated) request to an OData service, you need to:
<ls>
<li>Instance an object of the OAuthClient class</li>
<li>Set the client credentials (ClientID, Client Secret and Token Endpoint URL) for the object with the setClientCredentials method.</li>
<li>Perform the request through the sendGenericApiRequest method</li>
</ls>

<br />
The method signature is the following:

```vbnet
Public Function sendGenericApiRequest(
    url As String,
    method As httpMethod,
    body As String,
    ByVal headers As Scripting.Dictionary,
    queryParams As Scripting.Dictionary,
    Optional userID As String = "",
    Optional password As String = "",
    Optional forceTokenRefresh As Boolean = False
) As httpResponse
```

The arguments to this function are:
<ls>
<li><b>url:</b> a String value, representing the endpoint to send the request to.</li>
<li><b>method:</b> the HTTP method of the request (GET, POST, ...)</li>
<li><b>body:</b> a String value, representing the body of the request. Can be left blank, for request without body.</li>
<li><b>headers:</b> a Scripting.Dictionary object, representing Request Headers as key-value pairs</li>
<li><b>queryParams:</b> a Scripting.Dictionary object, representing Query Parameters as key-value pairs</li>
<li><b>userID:</b> (Optional) a String value, to be used in case username:password have to be sent explicitly within the request. Can be left blank for most cases.</li>
<li><b>password:</b> (Optional) a String value, to be used in case username:password have to be sent explicitly within the request. Can be left blank for most cases.</li>
<li><b>forceTokenRefresh:</b> (Optional) a Boolean value. If set to true, forces the client to request a new auth. token even if a valid previous one is available. Defaults to false.</li>
</ls>

<br />
<br />

The function returns a value of the  `httpResponse` type, defined as follows:

```vbnet
Public Type httpResponse
    status As httpResponseStatus    'Enum representing the overall status (1XX, 2XX, 3XX, 4XX, 5XX)
    statusCode As Integer           'Status Code of the Response
    statusText As String            'Status Text of the Response
    headers As String               'Headers of the Response
    text As String                  'Content of the Response, as raw text
End Type
```

<p align="right">(<a href="#top">back to top</a>)</p>

### Parsing JSON Responses

For responses with JSON payloads, the `libJSON` library provides a simple way of parsing the contents.<br>
The function used to parse a JSON string for a specific value is the following:

```vbnet
Public Function getJSONFieldValue(
    jsonString As String,
    fieldName As String,
    Optional valueIfNotFound As Variant = "",
    Optional valueIfNull As Variant = "null",
    Optional valueIfNaN As Variant = "NaN
) As Variant
```
The arguments to this function are:
<ls>
<li><b>jsonString:</b> The JSON string representing the object/array to parse </li>
<li><b>fieldName:</b> the name of the field to find in the JSON. For nested objects, can be specified as a path, like "nestedField1/nestedField2/field".</li>
<li><b>valueIfNotFound:</b> (Optional) the value to be returned in case the requested field is not found in the JSON object.</li>
<li><b>valueIfNull:</b> (Optional) the value to be returned in case the requested field is null.</li>
<li><b>valueIfNaN:</b> (Optional) the value to be returned in case the requested field is NaN.</li>
</ls>

The type of value returned from this function depends on the type of the field in the JSON Object.
|Value in the JSON Object      | Returned Value Type |
|:--------------:| :-------------------|
|String|String value|
|Numeric (no decimals)|LongLong value|
|Numeric (with decimals)|Double value|
|Boolean|Boolean value|
|null|value defined by the valueIfNotFound parameter|
|NaN|value defined by the valueIfNaN parameter|
|Object|String value representing a JSON encoded Object. If neccessary, it can be further parsed the same way as the original object|
|Array|Array of String Values. In case the array contains JSON Object, each element of the returned array can be further parsed for its fields.|

The function can correctly parse numbers expressed in scientific notation.

<p align="right">(<a href="#top">back to top</a>)</p>

### Sample
```vbnet
Sub sendTestRequest()
  
  Dim url As String
  Dim headers As New Scripting.Dictionary
  Dim queryParams As New Scripting.Dictionary

  'Declare and set a new OAuth client
  Dim client As odataClient
  client.CLIENTID = "xxxx"
  client.CLIENTSECRET = "xxxx"
  client.TOKENENDPOINT = "www.xxxx.com/odata/token"
  
  'Add Headers and Query Parameters to the request
  headers.Add "Sample-Header", "value"

  queryParams.Add "$top", 50
  queryParams.Add "$count", "true"

  'Send the request and get the Response
  Dim response As httpResponse
  response = libOData.sendODataGenericApiRequest(client, url, httpMethod_get, "", headers, queryParams)

  'Check the Response Status
  if response.status <> httpResponseOk Then
    Debug.Print "Error while processing request. Status code: " & response.statusCode & ", text: " & response.text
    Exit Sub
  End If
  
  'Process the response JSON content to extract a specific field
  Dim requestedValue As String
  requestedValue = getJSONFieldValue(response.text, "field1/field2")
  
  Debug.Print "The requested value is " & requestedValue

End Sub
```
