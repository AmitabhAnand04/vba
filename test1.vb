Option Explicit

Sub SaveVbaScriptToGitHub()
    
    'Declare our variables related to our URL
    Dim base_url As String
    Dim username As String
    Dim repo_name As String
    Dim file_name As String
    Dim access_token As String
    Dim payload As String
    Dim full_url As String
    
    'Declare variables related to the HTTP Request.
    Dim xml_obj As MSXML2.XMLHTTP60
    
    'Declare variables related to the Visual Basic Editor
    Dim VBAEditor As VBIDE.VBE
    Dim VBProj As VBIDE.VBProject
    Dim VBCodeMod As VBIDE.CodeModule
    Dim VBRawCode As String
    Dim RawCodeEncoded As String
    
    'Create a reference to the VB Editor, TURN OFF MACRO SECURITY!!!
    Set VBAEditor = Application.VBE
    
    'Grab the Visual Basic Project which will be commited
    Set VBProj = VBAEditor.VBProjects(1)
    
    'Reference a single component in our Project and then grab the code module.
    Set VBCodeMod = VBProj.VBComponents.Item("Module1").CodeModule
    
    'Grab the raw code in the code module
    VBRawCode = VBCodeMod.Lines(StartLine:=1, Count:=VBCodeMod.CountOfLines)
    
    'Debug.Print VBRawCode
    
    'Base64 Encode the string
    RawCodeEncoded = EncodeBase64(text:=VBRawCode)
    
    'Print out the code
    Debug.Print "Here is the encoded content: " + RawCodeEncoded
    
    'Define our XML HTTP Object
    Set xml_obj = New MSXML2.XMLHTTP60
    
        'Define our URL Components.
        base_url = "https://api.github.com/repos/"
        repo_name = "vba/"
        username = "AmitabhAnand04/"
        file_name = "test1.vb"
        access_token = "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"
        
        'Build the Full Url
        full_url = base_url + username + repo_name + "contents/" + file_name + "?ref=master"
        
        'Open a new request
        xml_obj.Open bstrMethod:="PUT", bstrUrl:=full_url, varAsync:=True
        
        'Set the headers
        xml_obj.setRequestHeader "Accept", "application/vnd.github.v3+json"
        xml_obj.setRequestHeader "Authorization", "token " + access_token
        
        'Define the payload
        payload = "{""message"": ""This is my Test1.vb"", ""content"":"""
        payload = payload + Application.Clean(RawCodeEncoded)
        payload = payload + """}"
        
        'Send the request.
        xml_obj.send varBody:=payload
        
        'Wait till it is finished.
        While xml_obj.readyState <> 4
            DoEvents
        Wend
        
        Debug.Print "RESPONSE: " + full_url
        'Print out some info
        'Debug.Print "FULL URL: " + full_url
        Debug.Print "STATUS TEXT: " + xml_obj.statusText
        'Debug.Print "PAYLOAD: " + payload
    
End Sub

Function EncodeBase64(text As String) As String
    'Define our variables.
    Dim arrData() As Byte
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    
    'Convert our string to a Unicode String
    arrData = StrConv(text, vbFromUnicode)
    
    'Define our Dom Objects.
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
    
    'Define the data Type.
    objNode.DataType = "bin.base64"
    
    'Assign the node value.
    objNode.nodeTypedValue = arrData
    
    'Return the Encoded Text.
    EncodeBase64 = Replace(objNode.text, vbLf, "")
    
    'Memory Cleanup
    Set objNode = Nothing
    Set objXML = Nothing
    
End Function