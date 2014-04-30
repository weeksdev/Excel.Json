Public Enum ResponseFormat
    Text
    Json
End Enum
Private pResponseText As String
Private pResponseJson
Private pScriptControl As Object
'Request method returns the responsetext and optionally will fill out json or xml objects
Public Function request(url As String, Optional postParameters As String = "", Optional format As ResponseFormat = ResponseFormat.Json) As String
    Dim xml
    Dim requestType As String
    If postParameters <> "" Then
        requestType = "POST"
    Else
        requestType = "GET"
    End If
    
    Set xml = CreateObject("MSXML2.XMLHTTP")
    xml.Open requestType, url, False
    If postParameters <> "" Then
        xml.send postParameters
    Else
        xml.send
    End If
    pResponseText = xml.ResponseText
    request = pResponseText
    Select Case format
        Case Json
            SetJson
    End Select
End Function
Private Sub SetJson()
    Dim qt As String
    qt = """"
    Set pScriptControl = CreateObject("scriptcontrol")
    pScriptControl.Language = "JScript"
    pScriptControl.eval "var obj=(" & pResponseText & ")"
    'pScriptControl.ExecuteStatement "var rootObj = null"
    pScriptControl.AddCode "function getObject(){return obj;}"
    'pScriptControl.eval "var rootObj=obj[" & qt & "query" & qt & "]"
    pScriptControl.AddCode "function getRootObject(){return rootObj;}"
    pScriptControl.AddCode "function getCount(){ return rootObj.length;}"
    pScriptControl.AddCode "function getBaseValue(){return baseValue;}"
    pScriptControl.AddCode "function getValue(){ return arrayValue;}"
    Set pResponseJson = pScriptControl.Run("getObject")
End Sub
Public Function setJsonRoot(rootPath As String)
    pScriptControl.ExecuteStatement "rootObj = obj." & rootPath
    Set setJsonRoot = pScriptControl.Run("getRootObject")
End Function
Public Function getJsonObjectCount()
    getJsonObjectCount = pScriptControl.Run("getCount")
End Function
Public Function getJsonObjectValue(path As String)
    pScriptControl.ExecuteStatement "baseValue = obj." & path
    getJsonObjectValue = pScriptControl.Run("getBaseValue")
End Function
Public Function getJsonArrayValue(index, key As String)
    Dim qt As String
    qt = """"
    If InStr(key, ".") > 0 Then
        arr = Split(key, ".")
        key = ""
        For Each cKey In arr
            key = key + "[" & qt & cKey & qt & "]"
        Next
    Else
        key = "[" & qt & key & qt & "]"
    End If
    Dim statement As String
    statement = "arrayValue = rootObj[" & index & "]" & key
    
    pScriptControl.ExecuteStatement statement
    getJsonArrayValue = pScriptControl.Run("getValue", index, key)
End Function
Public Property Get ResponseText() As String
    ResponseText = pResponseText
End Property
Public Property Get ResponseJson()
    ResponseJson = pResponseJson
End Property
Public Property Get ScriptControl() As Object
    ScriptControl = pScriptControl
End Property