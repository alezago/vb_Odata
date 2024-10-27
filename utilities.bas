Attribute VB_Name = "utilities"
Option Explicit

Public Function EncodeBase64(text As String) As String

Dim arrData() As Byte
arrData = StrConv(text, vbFromUnicode)

Dim objXML As Object
Dim objNode As Object

Set objXML = CreateObject("MSXML2.DOMDocument")
Set objNode = objXML.createElement("b64")

objNode.DataType = "bin.base64"
objNode.nodeTypedValue = arrData

'Remove newlines
EncodeBase64 = Replace(objNode.text, vbLf, "")

Set objNode = Nothing
Set objXML = Nothing

End Function

Public Function convertTimestampToDate(ts As String) As Date
'Input Format: YYYY-MM-DDThh:mm:ss.SSSZ"
Dim d As String
Dim t As String

d = Left(ts, 10)
t = Mid(ts, 12, 8)

convertTimestampToDate = DateValue(d) + TimeValue(t)

End Function

Public Function isArrayInitialized(arr() As String) As Boolean

Dim l As Long
On Error GoTo notInitialized
l = UBound(arr)
isArrayInitialized = True
On Error GoTo 0

Exit Function

notInitialized:
isArrayInitialized = False

End Function


Public Function arrCount(arr() As String) As Long

arrCount = 0

On Error GoTo notInitialized

arrCount = UBound(arr) - LBound(arr) + 1

notInitialized:

End Function

Public Function getElementRow(value As String, rng As Range) As Long

On Error GoTo notFound

getElementRow = Application.WorksheetFunction.Match(value, rng, 0)
Exit Function
notFound:
getElementRow = 0

End Function
