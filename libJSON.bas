Attribute VB_Name = "libJSON"
'**************************************************************
'* libJSON: Library for Parsing JSON in VBA
'**************************************************************
'* Author: Alessandro Zago
'* Last Modified: 15/01/2025
'* Version: 2.0.1
'**************************************************************
'* This library allows the retrieval of specific fields from an Object in JSON notation
'* It supports:
'*     - all elementary field types (string, number, boolean, null, NaN)
'*     - complex field types (object, array)
'*     - direct retrieval from fields in nested objects by providing a direct path parameter (i.e: parentObj/nestedObj1/nestedObj2/fieldName)
'*     - automatic conversion of the return value to the correct (VBA) Data Type, based on the field type in the JSON Object (see table below)
'*     - possibility to specify a custom return value in case of null / NaN fields, or for fields not found in the object
'*     - robust handling of escaped quotes in strings
'*     - direct retrieval of subFields from Nested Objects (with the synthax field/subField/subSubField)
'*     - parsing of JSON Arrays
'*
'* Important Notes:
'*     -  This library is NOT a JSON validator. It should be used only on complete and valid JSON objects. Using it on incomplete/invalid JSON objects can lead to unexpected results/errors.
'*     -  Arrays of objects are returned in VBA as arrays of Strings. Each element can then be parsed individually/in a loop for their respective field.
'*     -  Sub-Objects are returned as strings. These can be further parsed for their fields with the same method.
'*
'* Provided (Public) Functions and parameters:
'*     - getJSONFieldValue: parses a JSON Object provided as a String, and returns the value for a specific field.
'*           JsonString [String]: a String containing a (valid) JSON Object.
'*           fieldName [String]: a String containing the Path to the requested field. In case of nested objects, subfields can be retrieved with the "a/b/c/d" synthax
'*           valueIfNotFound [Variant][Optional]: the value to be returned in case the requested field is not found in the object. Defaults to the "" (empty) String
'*           valueIfNull [Variant][Optional]: the value to be returned in case the requested field is null. Defaults to the "null" String.
'*           valueIfNaN [Variant][Optional]: the value to be returned in case the requested field is NaN. Defaults to the "NaN" String.
'*     - getJSONArrayValue: parses a JSON Array provided as string, and returns an array of Strings.
'*           JsonArray: the JSON Array to be parsed, provided as a String.
'*
'* Return Types
'* String fields are returned as String values.
'* Numeric fields are returned as Integer or Double values (based on the original value).
'* Boolean fields are returned as Boolean values.
'* null fields return type can be customized on demand through an optional parameter provided to the parsing function
'* NaN fields return type can be customized on demand through an optional parameter provided to the parsing function
'* Object fields are returned as String values.
'* Array fields are returned as an Array of String values.
'*
'**************************************************************
Option Explicit

'Control Characters
Private Const CHAR_COLON As String = ":"
Private Const CHAR_COMMA As String = ","
Private Const CHAR_QUOTE As String = """"
Private Const CHAR_OPEN_OBJECT As String = "{"
Private Const CHAR_CLOSE_OBJECT As String = "}"
Private Const CHAR_OPEN_ARRAY As String = "["
Private Const CHAR_CLOSE_ARRAY As String = "]"

'Escaping
Private Const CHAR_ESCAPED_BACKSLASH As String = "\\"
Private Const CHAR_ESCAPED_QUOTE As String = "\"""
Private Const CHAR_ESCAPED_BACKSLASH_SANITIZED As String = "%ESCAPEDBACKSLASH%"
Private Const CHAR_ESCAPED_QUOTE_SANITIZED As String = "%ESCAPEDQUOTESANITIZED%"

'Constant Values/Strings
Private Const VALUE_NULL As String = "null"
Private Const VALUE_NAN As String = "NaN"
Private Const VALUE_TRUE As String = "true"
Private Const VALUE_FALSE As String = "false"

Public Enum JSONFieldType
    JSONFieldType_literal = 1
    JSONFieldType_number = 2
    JSONFieldType_array = 3    'lists [] in JSON are returned as arrays of string(s). Each element will have to be then parsed separately
    JSONFieldType_object = 5
    JSONFieldType_bool = 6
    JSONFieldType_null = 7
    JSONFieldType_nan = 8
    
    'Other (Supported) Types?
    JSONFieldType_unknown = -1     'Shouldn't occur, errors should be catched sooner
    JSONFieldType_No_Field = -2    'Used internally, when there are no more fields to parse
End Enum

'Key-Value pair. Used in parsing to detect and store information about a k-v pair in a JSON object. All indexes are 1-Based
Public Type JSONKeyValuePair
    key As String
    keyStartIndex As Long
    keyEndIndex As Long
    value As String
    valueStartIndex As Long
    valueEndIndex As Long
    valueType As JSONFieldType
End Type

Private Type searchResult
    term As String
    position As Long
End Type

Public Const JSON_PATH_DELIMITER As String = "/"

Private Function sanitizeJSON(JSONStr As String) As String

sanitizeJSON = Replace(Replace(Trim(JSONStr), CHAR_ESCAPED_BACKSLASH, CHAR_ESCAPED_BACKSLASH_SANITIZED), CHAR_ESCAPED_QUOTE, CHAR_ESCAPED_QUOTE_SANITIZED)

End Function

Private Function restoreJSON(JSONStr As String) As String

restoreJSON = Replace(Replace(JSONStr, CHAR_ESCAPED_QUOTE_SANITIZED, CHAR_ESCAPED_QUOTE), CHAR_ESCAPED_BACKSLASH_SANITIZED, CHAR_ESCAPED_BACKSLASH)

End Function

Private Function getFirstOccurringCharIndex(text As String, startPosition As Long, listOfControlChars() As Variant) As searchResult

Dim pos As Long
Dim currentMin As Long
Dim currentCharIndex As Integer
Dim i As Integer

For i = LBound(listOfControlChars) To UBound(listOfControlChars)
    pos = InStr(startPosition, text, listOfControlChars(i))
    If currentMin = 0 Or (pos > 0 And pos < currentMin) Then
        currentMin = pos
        currentCharIndex = i
    End If
Next i

Dim ret As searchResult

ret.position = currentMin
ret.term = listOfControlChars(currentCharIndex)

getFirstOccurringCharIndex = ret

End Function

Private Function getArrayEnd(JsonString As String, arrayOpenIndex As Long) As Long

If Mid(JsonString, arrayOpenIndex, 1) <> CHAR_OPEN_ARRAY Then
    Err.Raise 500, , "libJSON: getArrayEnd: start character for array is not ["
End If

Dim inString As Boolean
Dim currentPos As Long

Dim searchTerms() As Variant
Dim searchResult As searchResult

Dim arrayLevel As Integer

'Initialization
inString = False
currentPos = arrayOpenIndex + 1
arrayLevel = 1

Do While currentPos < Len(JsonString) And arrayLevel >= 1
    
    If inString Then
        searchTerms = Array(CHAR_QUOTE)
    Else
        searchTerms = Array(CHAR_QUOTE, CHAR_OPEN_ARRAY, CHAR_CLOSE_ARRAY)
    End If
    searchResult = getFirstOccurringCharIndex(JsonString, currentPos, searchTerms)
    
    'Validation
    If searchResult.position <= 0 Then
        Err.Raise 500, , "libJSON: getArrayEnd: unexpected end of string, malformed json."
    End If
    
    Select Case searchResult.term
        Case CHAR_QUOTE
            inString = Not inString
        Case CHAR_OPEN_ARRAY
            arrayLevel = arrayLevel + 1
        Case CHAR_CLOSE_ARRAY
            arrayLevel = arrayLevel - 1
    End Select
    
    currentPos = searchResult.position + 1

Loop

getArrayEnd = currentPos - 1

End Function

Private Function parseObject(JsonString As String, objectOpenIndex As Long) As String

If Mid(JsonString, objectOpenIndex, 1) <> CHAR_OPEN_OBJECT Then
    Err.Raise 500, , "libJSON: parseObject: start character for object is not {"
End If

Dim objStr As String

objStr = Mid(JsonString, objectOpenIndex, getObjectEnd(JsonString, objectOpenIndex) - objectOpenIndex + 1)
parseObject = restoreJSON(objStr)

End Function

Private Function parseArray(JsonString As String, arrayOpenIndex As Long) As String()

If Mid(JsonString, arrayOpenIndex, 1) <> CHAR_OPEN_ARRAY Then
    Err.Raise 500, , "libJSON: parseArray: start character for array is not ["
End If

Dim inString As Boolean
Dim currentPos As Long
Dim currentElementStartIndex As Long

Dim searchTerms() As Variant
Dim searchResult As searchResult

Dim inArray As Boolean

Dim output() As String
Dim elemCount As Integer
Dim tempElement As String

'Initialization
inString = False

inArray = True

elemCount = 0
currentPos = arrayOpenIndex + 1
currentElementStartIndex = currentPos

Do While currentPos <= Len(JsonString) And inArray
    
    If inString Then
        searchTerms = Array(CHAR_QUOTE)
    Else
        searchTerms = Array(CHAR_QUOTE, CHAR_OPEN_OBJECT, CHAR_OPEN_ARRAY, CHAR_COMMA, CHAR_CLOSE_ARRAY)
    End If
    searchResult = getFirstOccurringCharIndex(JsonString, currentPos, searchTerms)
    
    If searchResult.position <= 0 Then
        Err.Raise 500, , "libJSON: getArrayEnd: unexpected end of string, malformed json."
    End If
    
    Select Case searchResult.term
        Case CHAR_QUOTE
            inString = Not inString
            currentPos = searchResult.position + 1
        
        Case CHAR_OPEN_OBJECT
            currentPos = getObjectEnd(JsonString, searchResult.position) + 1
            
        Case CHAR_OPEN_ARRAY
            'In case of nested arrays, skip all the nested objects
            currentPos = getArrayEnd(JsonString, searchResult.position) + 1
            
        Case CHAR_COMMA
            'In case a comma is reached, add the element up to the character to the output
            If utilities.isArrayInitialized(output) Then
                ReDim Preserve output(0 To elemCount)
            Else
                ReDim output(0 To 0)
            End If
            
            'Add the new element to the output
            tempElement = Trim(Mid(JsonString, currentElementStartIndex, searchResult.position - currentElementStartIndex))
            output(elemCount) = restoreJSON(tempElement)
            
            'Update internal state for next element
            elemCount = elemCount + 1
            currentElementStartIndex = searchResult.position + 1
            currentPos = searchResult.position + 1
        
        Case CHAR_CLOSE_ARRAY
            'Add the last element to the array before exiting the loop
            'In this case, verify that the array is not an empty array
            tempElement = Trim(Mid(JsonString, currentElementStartIndex, searchResult.position - currentElementStartIndex))
            If tempElement <> "" Then
                If utilities.isArrayInitialized(output) Then
                    ReDim Preserve output(0 To elemCount)
                Else
                    ReDim output(0 To 0)
                End If
            
                'Add the new element to the output
                output(elemCount) = restoreJSON(tempElement)
            
            Else
                If utilities.isArrayInitialized(output) = False Then
                    'Initialize with empty array
                    output = Split(Empty)
                    
                End If
            End If
            
            inArray = False
    
    End Select
    
Loop

parseArray = output

End Function

Private Function parseLiteral(str As String) As String
parseLiteral = restoreJSON(str)
End Function

Private Function parseInteger(numString As String) As LongLong

If InStr(1, numString, "E") > 0 Then
    Dim expSplit() As String
    expSplit = Split(numString, "E")
    parseInteger = CLngLng(CDbl(Replace(expSplit(0), ".", Application.International(xlDecimalSeparator))) * (10 ^ CInt(expSplit(1))))
Else
    parseInteger = CLngLng(numString)
End If

End Function

Sub testInt()

Debug.Print parseInteger("1.163168865416E10")

End Sub

Private Function parseDecimal(numString As String) As Double

If InStr(1, numString, "E") > 0 Then
    Dim expSplit() As String
    expSplit = Split(numString, "E")
    parseDecimal = CDbl(Replace(expSplit(0), ".", Application.International(xlDecimalSeparator))) * (10 ^ CInt(expSplit(1)))
Else
    parseDecimal = CDbl(Replace(numString, ".", Application.International(xlDecimalSeparator)))
End If

End Function

Private Function parseNumber(numString) As Variant

If InStr(1, numString, "E") > 0 Then
    'Exponential (can be either integer or decimal)
    Dim expSplit() As String
    expSplit = Split(numString, "E")
    
    If CInt(expSplit(1)) > 0 And Len(expSplit(0)) - 2 <= CInt(expSplit(1)) Then
        'integer
        parseNumber = CLngLng(CDbl(Replace(expSplit(0), ".", Application.International(xlDecimalSeparator))) * (10 ^ CInt(expSplit(1))))
    Else
        'decimal
        parseNumber = CDbl(Replace(expSplit(0), ".", Application.International(xlDecimalSeparator))) * (10 ^ CInt(expSplit(1)))
    End If
    
ElseIf InStr(1, numString, ".") > 0 Then
    'Decimal
    parseNumber = CDbl(Replace(numString, ".", Application.International(xlDecimalSeparator)))

Else
    'Integer
    parseNumber = CLngLng(numString)

End If

End Function

Private Function parseBool(boolString As String) As Boolean

If boolString = VALUE_TRUE Then
    parseBool = True
ElseIf boolString = VALUE_FALSE Then
    parseBool = False
Else
    Err.Raise 500, , "libJSON: parseBool: provided value is not a Boolean value (true/false): " & boolString
End If

End Function

Private Function getObjectEnd(JsonString As String, objectOpenIndex As Long) As Long

If Mid(JsonString, objectOpenIndex, 1) <> CHAR_OPEN_OBJECT Then
    Err.Raise 500, , "libJSON: getObjectEnd: start character for object is not {"
End If

Dim inString As Boolean
Dim currentPos As Long

Dim searchTerms() As Variant
Dim searchResult As searchResult

Dim nestingLevel As Integer

'Initialization
inString = False
currentPos = objectOpenIndex + 1
nestingLevel = 1

Do While currentPos < Len(JsonString) And nestingLevel >= 1
    
    If inString Then
        searchTerms = Array(CHAR_QUOTE)
    Else
        searchTerms = Array(CHAR_QUOTE, CHAR_OPEN_OBJECT, CHAR_CLOSE_OBJECT)
    End If
    searchResult = getFirstOccurringCharIndex(JsonString, currentPos, searchTerms)
    
    'Validation
    If searchResult.position <= 0 Then
        Err.Raise 500, , "libJSON: getObjectEnd: unexpected end of string, malformed json."
    End If
    
    Select Case searchResult.term
        Case CHAR_QUOTE
            inString = Not inString
        Case CHAR_OPEN_OBJECT
            nestingLevel = nestingLevel + 1
        Case CHAR_CLOSE_OBJECT
            nestingLevel = nestingLevel - 1
    End Select
    
    currentPos = searchResult.position + 1

Loop

getObjectEnd = currentPos - 1

End Function

'Traverses a JSON Object horizontally, getting the next available keyValue pair
'ObjectStart should be always {
'ObjectEnd should be always }
'all indexes are 1-based (as returned from Excel Instr function)
Private Function getNextKeyOnLevel(JsonString As String, objectOpenIndex As Long, startFromIndex As Long) As JSONKeyValuePair

'Basic Validation
If Mid(JsonString, objectOpenIndex, 1) <> CHAR_OPEN_OBJECT Then
    Err.Raise 500, , "libJSON: ERROR: getNextKeyOnLevel: the substring provided starting at " & objectOpenIndex & " is not a valid JSON Object"
End If

If startFromIndex < objectOpenIndex Then
    Err.Raise 500, , "libJSON: ERROR: getNextKeyOnLevel: startFromIndex parameter is out of bounds (value provided: " & startFromIndex & ", objects starts at " & objectOpenIndex & ")"
End If

Dim searchTerms() As Variant
Dim searchResult As searchResult

Dim posKeyOpen As Long
Dim posKeyClose As Long
Dim posCurrent As Long

Dim tempChar As String

Dim returnValue As JSONKeyValuePair

'At the beginning, check if there is actually some more key to parse
searchTerms = Array(CHAR_CLOSE_OBJECT, CHAR_QUOTE)
searchResult = getFirstOccurringCharIndex(JsonString, startFromIndex, searchTerms)

If searchResult.term = CHAR_CLOSE_OBJECT Then
    returnValue.valueType = JSONFieldType_No_Field
    getNextKeyOnLevel = returnValue
    Exit Function
End If

posKeyOpen = searchResult.position

'do not exceed the object boundaries
If posKeyOpen <= 0 Or posKeyOpen > Len(JsonString) Then
    Err.Raise 500, , "libJson: getNextKeyOnLevel: something went wrong (DEBUGGING)"
End If

posKeyClose = InStr(posKeyOpen + 1, JsonString, CHAR_QUOTE)

'do not exceed the object boundaries
If posKeyClose <= 0 Or posKeyClose > Len(JsonString) Then
    Err.Raise 500, , "libJSON: ERROR: getNextKeyOnLevel: detected opening quotes at char. " & posKeyOpen & " without corresponding closing."
End If

returnValue.key = restoreJSON(Mid(JsonString, posKeyOpen + 1, posKeyClose - posKeyOpen - 1))
returnValue.keyStartIndex = posKeyOpen
returnValue.keyEndIndex = posKeyClose

posCurrent = posKeyClose + 1

'skip any whitespace present between the end of the key and the colon
Do While Mid(JsonString, posCurrent, 1) = " " And posCurrent <= Len(JsonString)
    posCurrent = posCurrent + 1
Loop

'Validation: key without value
If posCurrent = Len(JsonString) Then
    Err.Raise 500, , "libJSON: ERROR: getNextKeyOnLevel: key " & returnValue.key & " does not have an associated value within the object starting at " & objectOpenIndex
End If

'Validation: missing colon after key
If Mid(JsonString, posCurrent, 1) <> CHAR_COLON Then
    Err.Raise 500, , "libJSON: ERROR: getNextKeyOnLevel: unexpected character after key " & returnValue.key & " in object starting at " & objectOpenIndex
End If

posCurrent = posCurrent + 1
'skip any whitespace present between the colon and the value
Do While Mid(JsonString, posCurrent, 1) = " " And posCurrent <= Len(JsonString)
    posCurrent = posCurrent + 1
Loop

'attempt to detect the output type based on the next character
'For some types, also directly parses the value
'Possible Types:
'    Literal              (PARSED)
'    Numeric              (PARSED)
'    Boolean (true/false) (PARSED)
'    Null                 (PARSED)
'    NaN                  (PARSED)
'    Array                (parsing deferred)
'    Object               (parsing deferred)

If Mid(JsonString, posCurrent, 1) = CHAR_QUOTE Then    'Literal
    returnValue.valueType = JSONFieldType_literal
    returnValue.valueStartIndex = posCurrent
    posCurrent = InStr(posCurrent + 1, JsonString, CHAR_QUOTE)
    returnValue.valueEndIndex = posCurrent
    returnValue.value = Mid(JsonString, returnValue.valueStartIndex + 1, returnValue.valueEndIndex - returnValue.valueStartIndex - 1)
    returnValue.value = restoreJSON(returnValue.value)
    
ElseIf IsNumeric(Mid(JsonString, posCurrent, 1)) Or Mid(JsonString, posCurrent, 1) = "-" Then    'Numeric
    
    returnValue.valueStartIndex = posCurrent
    returnValue.value = Mid(JsonString, posCurrent, 1)
    
    Dim hasDecimalSep As Boolean
    Dim hasExponent As Boolean
    
    tempChar = Mid(JsonString, posCurrent + 1, 1)
    
    Do While (IsNumeric(tempChar) Or (hasDecimalSep = False And tempChar = ".") Or (hasExponent = False And tempChar = "E")) And posCurrent + 1 < Len(JsonString)
        returnValue.value = returnValue.value & tempChar
        If tempChar = "." Then
            hasDecimalSep = True
        End If
        
        If tempChar = "E" Then
            hasExponent = True
            hasDecimalSep = False
        End If
            
        posCurrent = posCurrent + 1
        tempChar = Mid(JsonString, posCurrent + 1, 1)
    Loop

    returnValue.valueType = JSONFieldType_number
    returnValue.valueEndIndex = posCurrent
    
ElseIf Mid(JsonString, posCurrent, 4) = VALUE_TRUE Then    'Boolean
    
    returnValue.valueType = JSONFieldType_bool
    returnValue.valueStartIndex = posCurrent
    returnValue.valueEndIndex = posCurrent + 3
    returnValue.value = VALUE_TRUE
    
    posCurrent = posCurrent + 4

ElseIf Mid(JsonString, posCurrent, 5) = VALUE_FALSE Then    'Boolean
        
    returnValue.valueType = JSONFieldType_bool
    returnValue.valueStartIndex = posCurrent
    returnValue.valueEndIndex = posCurrent + 4
    returnValue.value = VALUE_FALSE
    
    posCurrent = posCurrent + 5

ElseIf Mid(JsonString, posCurrent, 4) = VALUE_NULL Then    'null

    returnValue.valueType = JSONFieldType_null
    returnValue.valueStartIndex = posCurrent
    returnValue.valueEndIndex = posCurrent + 3
    returnValue.value = VALUE_NULL
    
    posCurrent = posCurrent + 4

ElseIf Mid(JsonString, posCurrent, 3) = VALUE_NAN Then    'NaN
    
    returnValue.valueType = JSONFieldType_nan
    returnValue.valueStartIndex = posCurrent
    returnValue.valueEndIndex = posCurrent + 2
    returnValue.value = VALUE_NAN
    
    posCurrent = posCurrent + 3

ElseIf Mid(JsonString, posCurrent, 1) = CHAR_OPEN_ARRAY Then    'Array
    returnValue.valueType = JSONFieldType_array
    returnValue.valueStartIndex = posCurrent
    'For Array and Objects, contents will be parsed outside only if necessary
    

ElseIf Mid(JsonString, posCurrent, 1) = CHAR_OPEN_OBJECT Then    'Object
    returnValue.valueType = JSONFieldType_object
    returnValue.valueStartIndex = posCurrent
    'For Arrays and Objects, contents will be parsed outside only if necessary

Else
    Err.Raise 500, , "libJSON: ERROR: getNextKeyOnLevel: cannot resolve value type for field " & returnValue.key & " in object starting at: " & objectOpenIndex
End If

getNextKeyOnLevel = returnValue

End Function

Public Function getJSONFieldValue(JsonString As String, fieldName As String, Optional valueIfNotFound As Variant = "", Optional valueIfNull As Variant = VALUE_NULL, Optional valueIfNaN As Variant = VALUE_NAN) As Variant

Dim tempString As String

Dim currentPath() As String
Dim targetPath() As String
Dim currentDepth As Integer    'Zero-Based
Dim targetDepth As Integer     'Zero-Based
Dim kvPair As JSONKeyValuePair

Dim referencePoint() As Long

'String Parsing
Dim currentPos As Long

tempString = sanitizeJSON(JsonString)
targetPath = Split(fieldName, JSON_PATH_DELIMITER)
targetDepth = UBound(targetPath) - LBound(targetPath)

'Initialization
currentPos = 1
currentDepth = 0
ReDim referencePoint(0 To 0)
referencePoint(0) = 1

Do While currentPos <= Len(tempString)
    
    kvPair = getNextKeyOnLevel(tempString, referencePoint(currentDepth), currentPos)
    
    'if all fields are exausted, return
    If kvPair.valueType = JSONFieldType_No_Field Then
        getJSONFieldValue = valueIfNotFound
        Debug.Print Now() & " - libJSON.getJSONFieldValue: Field '" & fieldName & "' does not exist in the provided JSON."
        Exit Function
    End If
    
    If targetPath(currentDepth) = kvPair.key Then
        
        If currentDepth < targetDepth Then
            'expected a branch but found a leaf
            If kvPair.valueType <> JSONFieldType_object Then
                getJSONFieldValue = valueIfNotFound
                Debug.Print Now() & " - libJSON.getJSONFieldValue: Field '" & fieldName & "' does not exist in the provided JSON. Search is interrupted at level '" & kvPair.key & "' (current depth: " & currentDepth + 1 & "/" & targetDepth + 1 & ")"
                Exit Function
            Else
                'navigate the tree to the next level
                currentDepth = currentDepth + 1
                ReDim Preserve referencePoint(0 To currentDepth)
                referencePoint(currentDepth) = kvPair.valueStartIndex
                currentPos = referencePoint(currentDepth)
            End If
        Else
            'found the target
            If kvPair.valueType = JSONFieldType_array Then
                getJSONFieldValue = parseArray(tempString, kvPair.valueStartIndex)
                
            ElseIf kvPair.valueType = JSONFieldType_object Then
                getJSONFieldValue = parseObject(tempString, kvPair.valueStartIndex)
                
            ElseIf kvPair.valueType = JSONFieldType_bool Then
                getJSONFieldValue = parseBool(kvPair.value)
            
            ElseIf kvPair.valueType = JSONFieldType_literal Then
                getJSONFieldValue = parseLiteral(kvPair.value)
                
            ElseIf kvPair.valueType = JSONFieldType_nan Then
                getJSONFieldValue = valueIfNaN
            
            ElseIf kvPair.valueType = JSONFieldType_null Then
                getJSONFieldValue = valueIfNull
                
            ElseIf kvPair.valueType = JSONFieldType_number Then
                getJSONFieldValue = parseNumber(kvPair.value)
            
            Else
                Debug.Print Now() & " - libJSON.getJSONFieldValue: Unexpected Value Type (" & kvPair.valueType & ")"
                Err.Raise 500, , "libJSON: getJSONFieldValue: field type not supported for field " & kvPair.key
            End If
            
            'Return the value
            Exit Function
            
        End If
    Else
    
        'Key is not the one we are looking for, set up for the next key search
        If kvPair.valueType = JSONFieldType_array Then
            'Array -> Skip the entire array
            currentPos = getArrayEnd(tempString, kvPair.valueStartIndex) + 1
        ElseIf kvPair.valueType = JSONFieldType_object Then
            'Object: Skip the entire object
            currentPos = getObjectEnd(tempString, kvPair.valueStartIndex) + 1
        Else
            currentPos = kvPair.valueEndIndex + 1
        End If
        
    End If

Loop

End Function


Public Function getJSONArrayValue(JsonArray As String) As String()

Dim tempString As String

tempString = sanitizeJSON(JsonArray)

getJSONArrayValue = parseArray(tempString, 1)

End Function

'returns a Formatted String
'naive implementation, does not work if the JSON contains characters to be escaped
Public Function prettify(inputText As String, Optional useTabs As Boolean = False) As String

Dim curlyOpen As Long, curlyClose As Long, squareOpen As Long, squareClose As Long, comma As Long
Dim currentPos As Long
Dim nextPos As Long
Dim length As Long
Dim maxLen As Long
Dim tabulation As String
Dim currentLevel As Integer
Dim newLine As Boolean
Dim i As Integer
Dim charAfterNewLine As String

Dim outputString As String

If useTabs Then
    tabulation = Chr(9)
Else
    tabulation = "    "
End If

Dim jsonText As String

jsonText = inputText

maxLen = Len(jsonText)

currentPos = 1
currentLevel = 0

While currentPos <= maxLen
    
    newLine = False
    charAfterNewLine = ""
    
    curlyOpen = InStr(currentPos, jsonText, "{")
    If curlyOpen < currentPos Then
        curlyOpen = maxLen
    End If
    
    curlyClose = InStr(currentPos, jsonText, "}")
    If curlyClose < currentPos Then
        curlyClose = maxLen
    End If
    
    squareOpen = InStr(currentPos, jsonText, "[")
    If squareOpen < currentPos Then
        squareOpen = maxLen
    End If
    
    squareClose = InStr(currentPos, jsonText, "]")
    If squareClose < currentPos Then
        squareClose = maxLen
    End If
    
    comma = InStr(currentPos, jsonText, ",")
    If comma < currentPos Then
        comma = maxLen
    End If
    
    nextPos = Application.WorksheetFunction.Min(curlyOpen, curlyClose, squareOpen, squareClose, comma)
    length = nextPos - currentPos + 1
    
    Select Case Mid(jsonText, nextPos, 1)
        Case "{"
            outputString = outputString & Mid(jsonText, currentPos, length)
            currentLevel = currentLevel + 1
            newLine = True    'Newline after {
        Case "}"
    
            If length > 1 Then
                outputString = outputString & Mid(jsonText, currentPos, length - 1)
            End If
            
            currentLevel = currentLevel - 1
            charAfterNewLine = "}"
            
            If currentPos < maxLen Then
                If Mid(jsonText, currentPos + 1, 1) = "," Then
                    newLine = True    'No Newline after } if followed by a comma
                Else
                    newLine = True
                End If
            Else
                newLine = True
            End If
            
        Case "["
            outputString = outputString & Mid(jsonText, currentPos, length)
            currentLevel = currentLevel + 1
            newLine = True    'Newline after {
        Case "]"
            
            If length > 1 Then
                outputString = outputString & Mid(jsonText, currentPos, length - 1)
            End If
            
            currentLevel = currentLevel - 1
            charAfterNewLine = "]"
            
            If currentPos < maxLen Then
                If Mid(jsonText, currentPos + 1, 1) = "," Then
                    newLine = True    'No Newline after ] if followed by a comma
                Else
                    newLine = True
                End If
            Else
                newLine = True
            End If
        Case ","
            outputString = outputString & Mid(jsonText, currentPos, length)
            newLine = True
    End Select
    
    If newLine Then
        outputString = outputString & vbCrLf
        If currentLevel > 0 Then
            For i = 1 To currentLevel
                outputString = outputString & tabulation
            Next i
        End If
    End If
    
    If charAfterNewLine <> "" Then
        outputString = outputString & charAfterNewLine
    End If
    
    currentPos = nextPos + 1
    
Wend

prettify = outputString

End Function
