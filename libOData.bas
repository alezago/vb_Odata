Attribute VB_Name = "libOData"
Option Explicit

Private Const DEFAULT_OAUTH_CLIENT_FILENAME As String = "oauth.json"

Private Const OAUTH_TOKEN_VALIDITY_THRESHOLD_RATE As Double = 0.9    'After 90% of the expected validity time is elapsed, assume that the token is expired
Private Const SECONDS_IN_A_DAY As Long = 86400

Private oauthToken As String
Private oauthTokenExpiration As Date

Public Type ODataClient
    TOKENENDPOINT As String
    CLIENTID As String
    CLIENTSECRET As String
End Type



Public Function hasValidOauthToken() As Boolean
If oauthToken <> "" And Now() < oauthTokenExpiration Then
    hasValidOauthToken = True
Else
    hasValidOauthToken = False
End If
End Function

Private Function requestNewToken(client As ODataClient) As Boolean

If client.CLIENTID = "" Or client.CLIENTSECRET = "" Or client.TOKENENDPOINT = "" Then
    Debug.Print "Invalid credentials provided."
    requestNewToken = False
    Exit Function
End If

'clear previous token (if available)
oauthToken = ""
oauthTokenExpiration = Now()

Dim authStr As String

authStr = client.CLIENTID & ":" & client.CLIENTSECRET
authStr = "Basic " & EncodeBase64(authStr)

Dim reqHeaders As New Scripting.Dictionary
Dim queryParams As New Scripting.Dictionary

Dim response As httpResponse
Dim expiration As Long

reqHeaders.Add "Authorization", authStr
reqHeaders.Add "Content-Type", "application/x-www-form-urlencoded"

response = libHTTP.sendHTTPRequest(client.TOKENENDPOINT, httpMethod_post, "grant_type=client_credentials", reqHeaders, queryParams, client.CLIENTID, client.CLIENTSECRET)

If response.status <> httpResponseOk Then
    Debug.Print "Could not get a valid token for the provided client. Token endpoint response: " & response.text
    requestNewToken = False
    Exit Function
End If

oauthToken = libJSON.getJSONFieldValue(response.text, "access_token")
expiration = libJSON.getJSONFieldValue(response.text, "expires_in", 1)
oauthTokenExpiration = Now() + (expiration * OAUTH_TOKEN_VALIDITY_THRESHOLD_RATE) / SECONDS_IN_A_DAY

If oauthToken <> "" And expiration > 0 Then
    requestNewToken = True
    Debug.Print "oAuth Token retrieved correctly."
Else
    requestNewToken = False
    Debug.Print "Something went wrong while retrieving the oAuth token. Please check your ClientID and ClientSecret."
End If

End Function

'Wrapper function for libHTTP.sendHTTPRequest that automatically handles authentication/authorization
'also includes the forceTokenRefresh parameter which, when set, forces the client to retrieve a new valid oAuth token
Public Function sendODataGenericApiRequest(oauthClient As ODataClient, url As String, method As httpMethod, body As String, ByVal headers As Dictionary, queryParams As Scripting.Dictionary, Optional userID As String = "", Optional password As String = "", Optional forceTokenRefresh As Boolean = False) As httpResponse

If hasValidOauthToken = False Or forceTokenRefresh Then
    If requestNewToken(oauthClient) <> True Then
        MsgBox "Something went wrong while retrieving a valid Authentication token from the Odata Token Endpoint, please check the log in the Debug window."
        Exit Function
    End If
End If

'add authentication header
If Not headers.Exists("Authorization") Then
    headers.Add "Authorization", "Bearer " & oauthToken
End If

sendODataGenericApiRequest = libHTTP.sendHTTPRequest(url, method, body, headers, queryParams, userID, password)

End Function
