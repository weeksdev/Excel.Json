Sub Example()
    Dim j
    'clear current range
    Range("A2:A1000").ClearContents
    'create ajax object
    Set j = New Json
    'make yql request for json
    j.request "https://query.yahooapis.com/v1/public/yql?q=show%20tables&format=json&callback=&diagnostics=true"
    'Debug.Print j.ResponseText
    'set root of data
    Set obj = j.setJsonRoot("query.results.table")
    Dim index
    'determine the total number of records returned
    index = j.getJsonObjectCount
    'if you need a field value from the object that is not in the array
    'tempValue = j.getJsonObjectValue("query.created")
    Dim x As Long
    x = 2
    If index > 0 Then
        For i = 0 To index - 1
            'set cell to the value of content field
            Range("A" & x).value = j.getJsonArrayValue(i, "content")
            x = x + 1
        Next
    Else
        MsgBox "No items found."
    End If
End Sub
Sub OceanExample()
    Dim j
    'clear current range
    Range("A2:A1000").ClearContents
    'create ajax object
    Set j = New Json
    'make yql request for json
    j.request "https://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20geo.oceans&format=json&diagnostics=true&callback="
    'Debug.Print j.ResponseText
    'set root of data
    Set obj = j.setJsonRoot("query.results.place")
    Dim index
    'determine the total number of records returned
    index = j.getJsonObjectCount
    'if you need a field value from the object that is not in the array
    'tempValue = j.getJsonObjectValue("query.created")
    Dim x As Long
    x = 2
    If index > 0 Then
        For i = 0 To index - 1
            'set cell to the value of content field
            Range("A" & x).value = j.getJsonArrayValue(i, "name")
            Range("B" & x).value = j.getJsonArrayValue(i, "woeid")
            Range("C" & x).value = j.getJsonArrayValue(i, "placeTypeName.content")
            Range("D" & x).value = j.getJsonArrayValue(i, "uri")
            x = x + 1
        Next
    Else
        MsgBox "No items found."
    End If
End Sub
