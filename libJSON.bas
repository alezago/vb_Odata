Attribute VB_Name = "libJSON"
Option Explicit

Private Const version_major As Integer = 1
Private Const version_minor As Integer = 0

Public Enum JSONFieldType
    JSONFieldType_literal = 1
    JSONFieldType_number = 2
    JSONFieldType_array = 3    'lists [] in JSON are returned as arrays of string(s). Each element will have to be then parsed separately
    JSONFieldType_float = 4
    JSONFieldType_object = 5
    JSONFieldType_bool = 6
    'Other (Supported) Types?
    JSONFieldType_unknown = -1
End Enum

Public Const JSON_PATH_DELIMITER As String = "/"
Public Const JSON_EMPTY_OBJECT = "{}"

'Fast implementation for getJSONFieldValue
Private Function getJSONFieldValue_f(jsonString As String, fieldName As String, valueType As JSONFieldType, Optional startFrom As Long = 1, Optional valueIfNotFound As Variant = "") As Variant

Dim tempJSON As String
Dim pathArr() As String
Dim nodeSearchString As String
Dim count As String
Dim startPos As Long, endPos As Long, endPos2 As Long

'remove unnecessary characters from start/end, and newlines, and in case start position is not 1 then also ignore all characters until the start position
tempJSON = Trim(Replace(Replace(Right(jsonString, Len(jsonString) - startFrom + 1), vbCrLf, ""), vbCr, ""))

'workaround to parse JSON arrays
'creates a normal JSON with an array field, named JSONARRAY
If fieldName = "" And valueType = JSONFieldType_array Then
    tempJSON = "{""JSONARRAY"": " & tempJSON & "}"
End If

'check that the first non blank character is a {
If Left(tempJSON, 1) <> "{" Then
    Debug.Print "getJSONFieldValue: the provided JSON String is not a valid JSON object, missing '{' at the start."
    getJSONFieldValue_f = valueIfNotFound
    Exit Function
End If

'split the path into an array, so it can be easily navigated from root to leaves
If fieldName = "" Or fieldName = "." Then
    pathArr = Split("JSONARRAY", JSON_PATH_DELIMITER)
ElseIf Left(fieldName, 2) = "./" Then
    pathArr = Split(Right(fieldName, Len(fieldName) - 2), JSON_PATH_DELIMITER)
Else
    pathArr = Split(fieldName, JSON_PATH_DELIMITER)
End If

'for the fast implementation, we only consider the leaf element
nodeSearchString = """" & pathArr(UBound(pathArr)) & """:"

startPos = InStr(startFrom, tempJSON, nodeSearchString)

If startPos < 1 Then
    getJSONFieldValue_f = valueIfNotFound
    Exit Function
End If

startPos = InStr(startPos, tempJSON, ":")

If valueType = JSONFieldType_literal Then
            
                startPos = InStr(startPos, tempJSON, """", vbTextCompare) + 1
                endPos = InStr(startPos, tempJSON, """", vbTextCompare)
                
                getJSONFieldValue_f = Mid(tempJSON, startPos, endPos - startPos)
                
            ElseIf valueType = JSONFieldType_number Then
                
                endPos = InStr(startPos, tempJSON, ",", vbTextCompare)
                endPos2 = InStr(startPos, tempJSON, "}", vbTextCompare)
                
                If endPos < endPos2 And endPos <> 0 Then
                    getJSONFieldValue_f = CLng(Trim(Mid(tempJSON, startPos, endPos - startPos)))
                Else
                    getJSONFieldValue_f = CLng(Trim(Mid(tempJSON, startPos, endPos2 - startPos)))
                End If
                
            ElseIf valueType = JSONFieldType_float Then
                
                endPos = InStr(startPos, tempJSON, ",", vbTextCompare)
                endPos2 = InStr(startPos, tempJSON, "}", vbTextCompare)
                
                If endPos < endPos2 And endPos <> 0 Then
                    getJSONFieldValue_f = val(Replace(Trim(Mid(tempJSON, startPos, endPos - startPos)), ".", ","))
                Else
                    getJSONFieldValue_f = val(Replace(Trim(Mid(tempJSON, startPos, endPos2 - startPos)), ".", ","))
                End If
            ElseIf valueType = JSONFieldType_array Then
                
                Dim returnArr() As String
                Dim arrLevel As Integer
                Dim inArray As Boolean
                Dim arrayPos As Long
                Dim elemStart As Long
                Dim elemCount As Integer
                
                inArray = True
                arrLevel = 0
                elemCount = 0
                
                arrayPos = InStr(startPos, tempJSON, "[", vbTextCompare) + 1
                
                While inArray
                    If arrLevel = 0 Then
                        'valid continuations:
                        '    ]  - close array
                        '    {  - new object (only first object in list
                        '    ,{ - new object (from second onwards)
                        'SORRY FOR THE NAMING OF VARIABLES
                        
                        startPos = InStr(arrayPos, tempJSON, "]", vbTextCompare)
                        If startPos <= 0 Then
                            Debug.Print "getJSONFieldValue: list at path " & fieldName & " doesn't have a matching closing bracket (])."
                            getJSONFieldValue_f = valueIfNotFound
                            Exit Function
                        End If
                        
                        endPos = InStr(arrayPos, tempJSON, "{", vbTextCompare)
                        
                        '
                        If startPos < endPos Or endPos = 0 Then
                            getJSONFieldValue_f = returnArr
                            Exit Function
                        Else
                            
                            arrayPos = endPos + 1
                            elemStart = endPos
                            arrLevel = arrLevel + 1
                            
                        End If
                        
                    ElseIf arrLevel > 0 Then
                        
                        'valid continuations:
                        '    { - increase level by one (may be missing)
                        '    } - decrease level by one (must be present)
                        startPos = InStr(arrayPos, tempJSON, "{", vbTextCompare)
                        endPos = InStr(arrayPos, tempJSON, "}", vbTextCompare)
                        
                        If startPos = 0 Or endPos < startPos Then
                            arrLevel = arrLevel - 1
                            
                            'if we returned to level 0, we isolated a new array element
                            If arrLevel = 0 Then
                                ReDim Preserve returnArr(0 To elemCount)
                                returnArr(elemCount) = Mid(tempJSON, elemStart, endPos - elemStart + 1)
                                elemCount = elemCount + 1
                            End If
                            
                            arrayPos = endPos + 1
                            
                        Else
                            
                            arrLevel = arrLevel + 1
                            arrayPos = startPos + 1
                            
                        End If
                        
                    End If
                Wend
            ElseIf valueType = JSONFieldType_bool Then
                
                endPos = InStr(startPos, tempJSON, ",", vbTextCompare)
                endPos2 = InStr(startPos, tempJSON, "}", vbTextCompare)
                
                If endPos < endPos2 And endPos <> 0 Then
                    getJSONFieldValue_f = CBool(Trim(Mid(tempJSON, startPos, endPos - startPos)))
                Else
                    getJSONFieldValue_f = CBool(Trim(Mid(tempJSON, startPos, endPos2 - startPos)))
                End If
            
            ElseIf valueType = JSONFieldType_object Then
                
                Dim objectStart As Long, currentPosition As Long
                Dim relativeLevel As Integer
                
                objectStart = InStr(startPos, tempJSON, "{", vbTextCompare)
                currentPosition = objectStart + 1
                
                relativeLevel = 1
                
                While currentPosition <= Len(tempJSON) And currentPosition > 0
                
                    startPos = InStr(currentPosition, tempJSON, "{", vbTextCompare)
                    endPos = InStr(currentPosition, tempJSON, "}", vbTextCompare) 'this should never be 0, as there must be a closing bracket at the very end of the object
                    
                    If startPos > 0 And startPos < endPos Then    'increase level
                        relativeLevel = relativeLevel + 1
                        currentPosition = startPos + 1
                        
                    Else    'decrease level
                        relativeLevel = relativeLevel - 1
                        currentPosition = endPos + 1
                    End If
                    
                    'check if the character was the closing character
                    If relativeLevel = 0 Then
                        getJSONFieldValue_f = Trim(Mid(tempJSON, objectStart, endPos - objectStart + 1))
                        Exit Function
                    End If
                Wend
            ElseIf valueType = JSONFieldType_unknown Then
            
                endPos = InStr(startPos, tempJSON, ",", vbTextCompare)
                endPos2 = InStr(startPos, tempJSON, "}", vbTextCompare)
                
                If endPos < endPos2 And endPos <> 0 Then
                    getJSONFieldValue_f = Trim(Mid(tempJSON, startPos, endPos - startPos))
                Else
                    getJSONFieldValue_f = Trim(Mid(tempJSON, startPos, endPos2 - startPos))
                End If
            End If
End Function



' JSONString: the full JSON object that need to be parsed. Must start with the { character, and ends when the matching } is found. Nothing is parsed beyond that
' fieldName: the full path of the field, starting from the JSON object root. ex: "./level1/level2/level3/fieldname". The initial "./" can be omitted
' valueType: the return type expected for the field (numeric or literal)
' startFrom: the starting point for the JSON object to be parsed inside the string, in case only a substring need to be parsed. The first character from the starting point must still be the { character
Public Function getJSONFieldValue(jsonString As String, fieldName As String, valueType As JSONFieldType, Optional startFrom As Long = 1, Optional valueIfNotFound As Variant = "", Optional valueIfNull As Variant = "", Optional useFastImpl As Boolean = False) As Variant

Dim tempJSON As String

Dim pathArr() As String
Dim i As Integer
Dim currentPath As String

Dim currentPosition As Long
Dim searchLevel As Integer
Dim currentLevel As Integer

Dim instrOpen As Long
Dim instrClose As Long
Dim instrSearch As Long

Dim fieldLookup As Long
Dim nestUpLookup As Long
Dim nestDownLookup As Long

'remove unnecessary characters from start/end, and newlines, and in case start position is not 1 then also ignore all characters until the start position
tempJSON = Trim(Replace(Replace(Right(jsonString, Len(jsonString) - startFrom + 1), vbCrLf, ""), vbCr, ""))

'workaround to parse JSON arrays
'creates a normal JSON with an array field, named JSONARRAY
If fieldName = "" And valueType = JSONFieldType_array Then
    tempJSON = "{""JSONARRAY"": " & tempJSON & "}"
End If

'check that the first non blank character is a {
If Left(tempJSON, 1) <> "{" Then
    Debug.Print "getJSONFieldValue: the provided JSON String is not a valid JSON object, missing '{' at the start."
    getJSONFieldValue = valueIfNotFound
    Exit Function
End If

'split the path into an array, so it can be easily navigated from root to leaves
If fieldName = "" Or fieldName = "." Then
    pathArr = Split("JSONARRAY", JSON_PATH_DELIMITER)
ElseIf Left(fieldName, 2) = "./" Then
    pathArr = Split(Right(fieldName, Len(fieldName) - 2), JSON_PATH_DELIMITER)
Else
    pathArr = Split(fieldName, JSON_PATH_DELIMITER)
End If


'Optimization:
'check if all the levels needed are present in the string somewhere, and in the correct order at least
currentPosition = 1
For searchLevel = LBound(pathArr) To UBound(pathArr)
    currentPosition = InStr(currentPosition, tempJSON, pathArr(searchLevel), vbTextCompare)
    If currentPosition = 0 Then
        'one of the levels needed to construct the path is missing, so the field cannot be found
        If valueType = JSONFieldType_array Then
            getJSONFieldValue = Split(Empty)
        Else
            getJSONFieldValue = valueIfNotFound
        End If
        
        Debug.Print "getJSONFieldValue: cannot construct path " & fieldName & ", the field cannot be found in the JSON Object."
        Exit Function
    Else
        currentPosition = currentPosition + 1
    End If
Next searchLevel

'reset current Position
currentPosition = 2  'ignore the first bracket, since we know the first character is always a "{"

searchLevel = LBound(pathArr)    'start looking for the initial term
currentLevel = searchLevel       'initially inside the first parenthesis

While currentPosition < Len(tempJSON)
    
    instrOpen = InStr(currentPosition, tempJSON, "{", vbTextCompare)
    
    If instrOpen = 0 Then
        instrOpen = Len(tempJSON) + 1
    End If
    
    instrClose = InStr(currentPosition, tempJSON, "}", vbTextCompare)
    
    If currentLevel = searchLevel Then
        instrSearch = InStr(currentPosition, tempJSON, """" & pathArr(searchLevel) & """:", vbTextCompare)
    Else
        instrSearch = Len(tempJSON) + 1
    End If
    
    If instrSearch <> 0 And instrSearch < instrOpen And instrSearch < instrClose And currentLevel = searchLevel Then
        'search term found
        If searchLevel = UBound(pathArr) Then
            'get the value and return
            Dim instr1 As Long, instr2 As Long, instr3 As Long
            
            instr1 = InStr(instrSearch, tempJSON, ":", vbTextCompare) + 1
            
            'check that the requested value is not null
            If LCase(Mid(tempJSON, instr1, 4)) = "null" Then
                getJSONFieldValue = valueIfNull
            
            ElseIf valueType = JSONFieldType_literal Then
            
                instr1 = InStr(instr1, tempJSON, """", vbTextCompare) + 1
                instr2 = InStr(instr1, tempJSON, """", vbTextCompare)
                
                getJSONFieldValue = Mid(tempJSON, instr1, instr2 - instr1)
                
            ElseIf valueType = JSONFieldType_number Then
                
                instr2 = InStr(instr1, tempJSON, ",", vbTextCompare)
                instr3 = InStr(instr1, tempJSON, "}", vbTextCompare)
                
                If instr2 < instr3 And instr2 <> 0 Then
                    getJSONFieldValue = CLng(Trim(Mid(tempJSON, instr1, instr2 - instr1)))
                Else
                    getJSONFieldValue = CLng(Trim(Mid(tempJSON, instr1, instr3 - instr1)))
                End If
                
            ElseIf valueType = JSONFieldType_float Then
                
                instr2 = InStr(instr1, tempJSON, ",", vbTextCompare)
                instr3 = InStr(instr1, tempJSON, "}", vbTextCompare)
                
                Dim sep As String
                
                If Application.UseSystemSeparators Then
                    sep = Application.International(xlDecimalSeparator)
                Else
                    sep = Application.DecimalSeparator
                End If
                
                If instr2 < instr3 And instr2 <> 0 Then
                    getJSONFieldValue = CDbl(Replace(Trim(Mid(tempJSON, instr1, instr2 - instr1)), ".", sep))
                Else
                    getJSONFieldValue = CDbl(Replace(Trim(Mid(tempJSON, instr1, instr3 - instr1)), ".", sep))
                End If
            ElseIf valueType = JSONFieldType_array Then
            
                Dim returnArr() As String
                Dim arrLevel As Integer
                Dim inArray As Boolean
                Dim arrayPos As Long
                Dim elemStart As Long
                Dim elemCount As Integer
                inArray = True
                arrLevel = 0
                elemCount = 0
                
                arrayPos = InStr(instr1, tempJSON, "[", vbTextCompare) + 1
                
                While inArray
                    If arrLevel = 0 Then
                        'valid continuations:
                        '    ]  - close array
                        '    {  - new object (only first object in list
                        '    ,{ - new object (from second onwards)
                        instr1 = InStr(arrayPos, tempJSON, "]", vbTextCompare)
                        If instr1 <= 0 Then
                            Debug.Print "getJSONFieldValue: list at path " & fieldName & " doesn't have a matching closing bracket (])."
                            getJSONFieldValue = valueIfNotFound
                            Exit Function
                        End If
                        
                        instr2 = InStr(arrayPos, tempJSON, "{", vbTextCompare)
                        
                        '
                        If instr1 < instr2 Or instr2 = 0 Then
                            getJSONFieldValue = returnArr
                            Exit Function
                        Else
                            
                            arrayPos = instr2 + 1
                            elemStart = instr2
                            arrLevel = arrLevel + 1
                            
                        End If
                        
                    ElseIf arrLevel > 0 Then
                        
                        'valid continuations:
                        '    { - increase level by one (may be missing)
                        '    } - decrease level by one (must be present)
                        instr1 = InStr(arrayPos, tempJSON, "{", vbTextCompare)
                        instr2 = InStr(arrayPos, tempJSON, "}", vbTextCompare)
                        
                        If instr1 = 0 Or instr2 < instr1 Then
                            arrLevel = arrLevel - 1
                            
                            'if we returned to level 0, we isolated a new array element
                            If arrLevel = 0 Then
                                ReDim Preserve returnArr(0 To elemCount)
                                returnArr(elemCount) = Mid(tempJSON, elemStart, instr2 - elemStart + 1)
                                elemCount = elemCount + 1
                            End If
                            
                            arrayPos = instr2 + 1
                            
                        Else
                            
                            arrLevel = arrLevel + 1
                            arrayPos = instr1 + 1
                            
                        End If
                        
                    End If
                Wend
            ElseIf valueType = JSONFieldType_bool Then
                
                instr2 = InStr(instr1, tempJSON, ",", vbTextCompare)
                instr3 = InStr(instr1, tempJSON, "}", vbTextCompare)
                
                If instr2 < instr3 And instr2 <> 0 Then
                    getJSONFieldValue = CBool(Trim(Mid(tempJSON, instr1, instr2 - instr1)))
                Else
                    getJSONFieldValue = CBool(Trim(Mid(tempJSON, instr1, instr3 - instr1)))
                End If
            
            ElseIf valueType = JSONFieldType_object Then
                
                Dim objectStart As Long
                Dim relativeLevel As Integer
                
                objectStart = InStr(instr1, tempJSON, "{", vbTextCompare)
                currentPosition = objectStart + 1
                
                relativeLevel = 1
                
                While currentPosition <= Len(tempJSON) And currentPosition > 0
                
                    instr1 = InStr(currentPosition, tempJSON, "{", vbTextCompare)
                    instr2 = InStr(currentPosition, tempJSON, "}", vbTextCompare) 'this should never be 0, as there must be a closing bracket at the very end of the object
                    
                    If instr1 > 0 And instr1 < instr2 Then    'increase level
                        relativeLevel = relativeLevel + 1
                        currentPosition = instr1 + 1
                        
                    Else    'decrease level
                        relativeLevel = relativeLevel - 1
                        currentPosition = instr2 + 1
                    End If
                    
                    'check if the character was the closing character
                    If relativeLevel = 0 Then
                        getJSONFieldValue = Trim(Mid(tempJSON, objectStart, instr2 - objectStart + 1))
                        Exit Function
                    End If

                Wend
                
                
            ElseIf valueType = JSONFieldType_unknown Then
            
                instr2 = InStr(instr1, tempJSON, ",", vbTextCompare)
                instr3 = InStr(instr1, tempJSON, "}", vbTextCompare)
                
                If instr2 < instr3 And instr2 <> 0 Then
                    getJSONFieldValue = Trim(Mid(tempJSON, instr1, instr2 - instr1))
                Else
                    getJSONFieldValue = Trim(Mid(tempJSON, instr1, instr3 - instr1))
                End If
            End If
            
            Exit Function
            
        Else
            'keep navigating inside
            currentPosition = InStr(instrSearch, tempJSON, "{", vbTextCompare) + 1
            searchLevel = searchLevel + 1
            currentLevel = currentLevel + 1
            
        End If
        
        
    ElseIf instrOpen < instrClose And instrOpen <> 0 Then
        currentLevel = currentLevel + 1
        currentPosition = instrOpen + 1
        
    Else  'instrClose  < instrOpen
        currentLevel = currentLevel - 1
        currentPosition = instrClose + 1
        
    End If
    
    
    'check if we fell out of the loop
    If currentLevel < searchLevel Then
        Dim exPath As String
        
        exPath = "."
        
        For i = LBound(pathArr) To searchLevel
            exPath = exPath & JSON_PATH_DELIMITER & pathArr(i)
        Next i
        
        Debug.Print "getJSONFieldValue: cannot construct path " & exPath & "."
        getJSONFieldValue = valueIfNotFound
        Exit Function
        
    End If
Wend

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

'If Left(inputText, 1) = "[" Then
'    jsonText = "{""ARRAY"":" & inputText & "}"
'Else
    jsonText = inputText
'End If

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
