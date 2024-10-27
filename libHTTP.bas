Attribute VB_Name = "libHTTP"
Option Explicit

Public Enum httpResponseStatus
    httpResponseUnknown = 0
    httpResponseInfo = 1            '1xx
    httpResponseOk = 2              '2xx
    httpResponseRedirect = 3        '3xx
    httpResponseErrorClient = 4     '4xx
    httpResponseErrorServer = 5     '5xx
End Enum

Public Enum httpMethod
    httpMethod_get = 1
    httpMethod_post = 2
    httpMethod_put = 3
    httpMethod_patch = 4
    httpMethod_delete = 5
    httpMethod_head = 6
    httpMethod_options = 7
End Enum

Public Type httpResponse
    status As httpResponseStatus
    statusCode As Integer
    statusText As String
    headers As String 'TODO: put headers in a dictionary or an array
    text As String
End Type


Private Function getStatus(statusCode As Integer) As httpResponseStatus

If statusCode >= 100 And statusCode < 200 Then
    getStatus = httpResponseInfo
ElseIf statusCode >= 200 And statusCode < 300 Then
    getStatus = httpResponseOk
ElseIf statusCode >= 300 And statusCode < 400 Then
    getStatus = httpResponseRedirect
ElseIf statusCode >= 400 And statusCode < 500 Then
    getStatus = httpResponseErrorClient
ElseIf statusCode >= 500 And statusCode < 600 Then
    getStatus = httpResponseErrorServer
Else
    getStatus = httpResponseUnknown
End If

End Function

Private Function getHTTPMethodAsEnum(methodStr As String) As httpMethod

Select Case methodStr
    Case "GET"
        getHTTPMethodAsEnum = httpMethod_get
    Case "POST"
        getHTTPMethodAsEnum = httpMethod_post
    Case "PUT"
        getHTTPMethodAsEnum = httpMethod_put
    Case "PATCH"
        getHTTPMethodAsEnum = httpMethod_patch
    Case "DELETE"
        getHTTPMethodAsEnum = httpMethod_delete
    Case "HEAD"
        getHTTPMethodAsEnum = httpMethod_head
    Case "OPTIONS"
        getHTTPMethodAsEnum = httpMethod_options
End Select

End Function

Private Function getHTTPMethodAsString(methodID As httpMethod) As String

Select Case methodID
    Case httpMethod_get
        getHTTPMethodAsString = "GET"
    Case httpMethod_post
        getHTTPMethodAsString = "POST"
    Case httpMethod_put
        getHTTPMethodAsString = "PUT"
    Case httpMethod_patch
        getHTTPMethodAsString = "PATCH"
    Case httpMethod_delete
        getHTTPMethodAsString = "DELETE"
    Case httpMethod_head
        getHTTPMethodAsString = "HEAD"
    Case httpMethod_options
        getHTTPMethodAsString = "OPTIONS"
End Select

End Function


Public Function sendHTTPRequest(url As String, method As httpMethod, body As String, headers As Dictionary, queryParams As Scripting.Dictionary, Optional userID As String = "", Optional password As String = "") As httpResponse

Dim request As New MSXML2.XMLHTTP60
Dim methodStr As String
Dim h As Variant
Dim finalURL As String

methodStr = getHTTPMethodAsString(method)

finalURL = url

If Not queryParams Is Nothing Then
    If queryParams.count > 0 Then
        finalURL = finalURL & "?"
        For Each h In queryParams.Keys
            finalURL = finalURL & Application.WorksheetFunction.EncodeURL(CStr(h)) & "=" & Application.WorksheetFunction.EncodeURL(CStr(queryParams(h))) & "&"
        Next h
        finalURL = Left(finalURL, Len(finalURL) - 1)    'Remove trailing "&"
    End If
End If

Debug.Print "Sending request: " & finalURL

'Open Request
If userID <> "" And password <> "" Then
    request.Open methodStr, finalURL, False, userID, password
Else
    request.Open methodStr, finalURL, False
End If

'Add Headers
If Not headers Is Nothing Then
    If headers.count > 0 Then
        For Each h In headers.Keys
            request.setRequestHeader CStr(h), CStr(headers(h))
        Next h
    End If
End If

request.send body

sendHTTPRequest.statusCode = request.status
sendHTTPRequest.status = getStatus(request.status)
sendHTTPRequest.statusText = request.statusText
sendHTTPRequest.headers = request.getAllResponseHeaders
sendHTTPRequest.text = request.responseText

End Function
