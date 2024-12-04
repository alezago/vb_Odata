VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OAuthClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const OAUTH_TOKEN_VALIDITY_THRESHOLD_RATE As Double = 0.9    'After 90% of the expected validity time is elapsed, assume that the token is expired
Private Const SECONDS_IN_A_DAY As Long = 86400

Private clientId As String
Private clientSecret As String
Private tokenEndpoint As String

Private oauthToken As String
Private oauthTokenExpiration As Date

Private isInitialized As Boolean

'Class Constructor
Private Sub Class_Initialize()
isInitialized = False
End Sub

Public Function setClientCredentials(inputClientId As String, inputClientSecret As String, inputTokenEndpoint As String)

clientId = inputClientId
clientSecret = inputClientSecret
tokenEndpoint = inputTokenEndpoint

'reset current token
oauthToken = ""
oauthTokenExpiration = Now()

isInitialized = True

End Function

Public Function hasValidOauthToken() As Boolean

If isInitialized And oauthToken <> "" And Now() < oauthTokenExpiration Then
    hasValidOauthToken = True
Else
    hasValidOauthToken = False
End If
End Function

Private Function requestNewToken() As Boolean

If clientId = "" Or clientSecret = "" Or tokenEndpoint = "" Then
    Debug.Print "Invalid credentials provided."
    requestNewToken = False
    Exit Function
End If

'clear previous token (if available)
oauthToken = ""
oauthTokenExpiration = Now()

Dim authStr As String

authStr = clientId & ":" & clientSecret
authStr = "Basic " & EncodeBase64(authStr)

Dim reqHeaders As New Scripting.Dictionary
Dim queryParams As New Scripting.Dictionary

Dim response As httpResponse
Dim expiration As Long

reqHeaders.Add "Authorization", authStr
reqHeaders.Add "Content-Type", "application/x-www-form-urlencoded"

response = libHTTP.sendHTTPRequest(tokenEndpoint, httpMethod_post, "grant_type=client_credentials", reqHeaders, queryParams, clientId, clientSecret)

If response.status <> httpResponseOk Then
    Debug.Print "No credentials file found in the required path/file malformed."
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
Public Function sendGenericApiRequest(url As String, method As httpMethod, body As String, ByVal headers As Scripting.Dictionary, queryParams As Scripting.Dictionary, Optional userID As String = "", Optional password As String = "", Optional forceTokenRefresh As Boolean = False) As httpResponse

If Not isInitialized Then
    MsgBox "The OAuth client is not initialized"
    Exit Function
End If

If hasValidOauthToken = False Or forceTokenRefresh Then
    If requestNewToken <> True Then
        MsgBox "Something went wrong while retrieving a valid Authentication token from the OAuth Token Endpoint, please check the log in the Debug window."
        Exit Function
    End If
End If

'add authentication header
If Not headers.Exists("Authorization") Then
    headers.Add "Authorization", "Bearer " & oauthToken
End If

sendGenericApiRequest = libHTTP.sendHTTPRequest(url, method, body, headers, queryParams, userID, password)

End Function


