Option Explicit
Public gMcp As Object      'WshScriptExec – stays alive for the whole session

Private mNextId As Long

'--- helper to get the next unique request id -------------
Private Function NextReqId() As Long
    mNextId = mNextId + 1
    NextReqId = mNextId
End Function

'Launch the Dataverse MCP server and return the WshScriptExec handle
Function StartMcpServer() As Object
    Dim cmd As String, wsh As Object
    cmd = """" & Environ$("USERPROFILE") & _
          "\.dotnet\tools\Microsoft.PowerPlatform.Dataverse.MCP""" & _
          " --ConnectionUrl [Url to the connection that was set up in Power Automate]" & _
          " --MCPServerName DataverseMCPServer" & _
          " --TenantId [Your tenant ID]" & _
          " --EnableHttpLogging true --EnableMsalLogging false --Debug false --BackendProtocol HTTP"

    Set wsh = CreateObject("WScript.Shell")
    Set gMcp = wsh.Exec(cmd)
End Function

Sub StartAndInitMcp()

    If Not gMcp Is Nothing Then Exit Sub      'already running
    StartMcpServer

    Dim reqId As Long
    reqId = NextReqId()
    '----- 1.  SEND INITIALIZE -------------------------------------------------
    Dim init As String
    init = _
      "{""jsonrpc"":""2.0"",""id"":" & reqId & ",""method"":""initialize"",""params"":{" & _
      """protocolVersion"":""2025-03-26"",""capabilities"":{}," & _
      """clientInfo"":{""name"":""Excel VBA Client"",""version"":""0.1""}}}"

    gMcp.StdIn.WriteLine init        'newline terminates the message
    Debug.Print "?", init

    '----- 2.  WAIT FOR InitializeResult --------------------------------------
    Dim line As String, gotResult As Boolean
    Do
        DoEvents
        If Not gMcp.StdOut.AtEndOfStream Then
            line = gMcp.StdOut.ReadLine
            Debug.Print "?", line        'log to Immediate window
            If InStr(line, """result""") > 0 Or InStr(line, """error""") > 0 Then
                gotResult = True
                Exit Do
            End If
        End If
    Loop While gMcp.Status = 0 And Not gotResult

    '----- 3.  SEND notifications/initialized ---------------------------------
    gMcp.StdIn.WriteLine "{""jsonrpc"":""2.0"",""method"":""notifications/initialized""}"
    
End Sub

Sub StopMcpServer()
    On Error Resume Next
    If gMcp Is Nothing Then Exit Sub

    gMcp.StdIn.Close          'tell the server we’re done (EOF)
    Application.Wait Now + TimeSerial(0, 0, 2)
    If gMcp.Status = 0 Then gMcp.Terminate
    Set gMcp = Nothing
End Sub


Function read_query(querytext As String)
    If gMcp Is Nothing Then
        MsgBox "MCP server is not running – start it first.", vbExclamation
        Exit Function
    End If

    Dim reqId  As Long

    '----- 1. build & send request --------------------
    reqId = NextReqId()
    Dim jsonReq As String
     
     jsonReq = "{""jsonrpc"":""2.0"",""id"":" & reqId & _
      ",""method"":""tools/call"",""params"":{" & _
      """name"":""read_query"",""arguments"":{" & _
      """querytext"":" & querytext & "}}}"

    gMcp.StdIn.WriteLine jsonReq

    '----- 2. wait for reply --------------------------
    Dim line As String, resp As Dictionary
    Do
        DoEvents
        If gMcp.StdOut.AtEndOfStream Then
            If gMcp.Status <> 0 Then
                MsgBox "MCP process exited unexpectedly.", vbCritical
                Exit Function
            End If
        Else
            line = gMcp.StdOut.ReadLine
            Debug.Print "?", line
            Set resp = ParseJson(line)
            If resp("id") = reqId Then Exit Do
        End If
    Loop

    '----- 3. handle errors ---------------------------
    If resp.Exists("error") Then
        MsgBox "tools/call error: " & resp("error")("message"), vbCritical
        Exit Function
    End If
    
    Set read_query = resp("result")
End Function

Function update_record(tablename As String, recordId As String, item As String)
    If gMcp Is Nothing Then
        MsgBox "MCP server is not running – start it first.", vbExclamation
        Exit Function
    End If

    Dim reqId  As Long

    '----- 1. build & send request --------------------
    reqId = NextReqId()
    Dim jsonReq As String
     
     jsonReq = "{""jsonrpc"":""2.0"",""id"":" & reqId & _
      ",""method"":""tools/call"",""params"":{" & _
      """name"":""update_record"",""arguments"":{" & _
      """tablename"":" & tablename & ", " & _
      """recordId"":" & recordId & ", " & _
      """item"":" & item & _
      "}}}"
    
    Debug.Print "?", jsonReq

    gMcp.StdIn.WriteLine jsonReq

    '----- 2. wait for reply --------------------------
    Dim line As String, resp As Dictionary
    Do
        DoEvents
        If gMcp.StdOut.AtEndOfStream Then
            If gMcp.Status <> 0 Then
                MsgBox "MCP process exited unexpectedly.", vbCritical
                Exit Function
            End If
        Else
            line = gMcp.StdOut.ReadLine
            Debug.Print "?", line
            Set resp = ParseJson(line)
            If resp("id") = reqId Then Exit Do
        End If
    Loop

    '----- 3. handle errors ---------------------------
    If resp.Exists("error") Then
        MsgBox "tools/call error: " & resp("error")("message"), vbCritical
        Exit Function
    End If
    
    Set update_record = resp("result")
End Function

Public Function DvMcpUpdateRecord(tablename As String, recordId As String, item As String) As String
    StartAndInitMcp

    Dim resp As Dictionary
    tablename = """" & tablename & """"
    recordId = """" & recordId & """"
    item = """" & item & """"
    
    Set resp = update_record(tablename, recordId, item)
    
    'The text payload is JSON in result("content")(1)("text")
    Dim outerJson As String
    outerJson = resp("content")(1)("text")
    
    Debug.Print "?", outerJson
    
    DvMcpUpdateRecord = "Successfully updated!"

End Function

Function create_record(tablename As String, item As String)
    If gMcp Is Nothing Then
        MsgBox "MCP server is not running – start it first.", vbExclamation
        Exit Function
    End If

    Dim reqId  As Long

    '----- 1. build & send request --------------------
    reqId = NextReqId()
    Dim jsonReq As String
     
     jsonReq = "{""jsonrpc"":""2.0"",""id"":" & reqId & _
      ",""method"":""tools/call"",""params"":{" & _
      """name"":""create_record"",""arguments"":{" & _
      """tablename"":" & tablename & ", " & _
      """item"":" & item & _
      "}}}"
    
    Debug.Print "?", jsonReq

    gMcp.StdIn.WriteLine jsonReq

    '----- 2. wait for reply --------------------------
    Dim line As String, resp As Dictionary
    Do
        DoEvents
        If gMcp.StdOut.AtEndOfStream Then
            If gMcp.Status <> 0 Then
                MsgBox "MCP process exited unexpectedly.", vbCritical
                Exit Function
            End If
        Else
            line = gMcp.StdOut.ReadLine
            Debug.Print "?", line
            Set resp = ParseJson(line)
            If resp("id") = reqId Then Exit Do
        End If
    Loop

    '----- 3. handle errors ---------------------------
    If resp.Exists("error") Then
        MsgBox "tools/call error: " & resp("error")("message"), vbCritical
        Exit Function
    End If
    
    Set create_record = resp("result")
End Function

Public Function DvMcpCreateRecord(tablename As String, item As String) As String
    StartAndInitMcp

    Dim resp As Dictionary
    tablename = """" & tablename & """"
    item = """" & item & """"
    
    Set resp = create_record(tablename, item)
    
    'The text payload is JSON in result("content")(1)("text")
    Dim outerJson As String
    outerJson = resp("content")(1)("text")
    
    Debug.Print "?", outerJson
    
    DvMcpCreateRecord = "Successfully created!"

End Function


Public Function DvMcpReadQuery(querytext As String) As Variant
    '--- 1. make sure the MCP CLI is running -------------------------------
    StartAndInitMcp             'no UI, so legal in a UDF
    
    '--- 2. call the tool ---------------------------------------------------
    Dim resp As Dictionary
    querytext = """" & querytext & """"
    
    Set resp = read_query(querytext)
    
    'The text payload is JSON in result("content")(1)("text")
    Dim outerJson As String
    outerJson = resp("content")(1)("text")
    
    Dim outerObj As Dictionary
    Set outerObj = ParseJson(outerJson)
    
    'outerObj("queryresult") is itself a JSON string -> parse again
    Dim rows As Collection
    Set rows = ParseJson(outerObj("queryresult"))
    If rows.Count = 0 Then
        DvMcpReadQuery = CVErr(xlErrNA)
        Exit Function
    End If
    
    '--- 3. build the 2-D array --------------------------------------------
    Dim headers As Variant, firstRec As Dictionary, r As Long, c As Long
    Set firstRec = rows(1)
    headers = firstRec.Keys        'Variant/array
    
    Dim outArr() As Variant
    ReDim outArr(0 To rows.Count, 0 To UBound(headers))
    
    'header row
    For c = 0 To UBound(headers)
        outArr(0, c) = headers(c)
    Next c
    
    'data rows
    For r = 1 To rows.Count
        Dim rec As Dictionary: Set rec = rows(r)
        For c = 0 To UBound(headers)
            outArr(r, c) = rec(headers(c))
        Next c
    Next r
    
    '--- 4. return – Excel spills it automatically -------------------------
    DvMcpReadQuery = outArr                 'Variant(1:nRows, 1:nCols)
End Function


'-------------------------------------------------------------
'  Function McpListTools()
'  Returns a 2-D array that Excel spills automatically.
'  It DOES NOT write to the sheet directly.
'-------------------------------------------------------------
Public Function DvMcpListTools() As Variant
    '1. Ensure the MCP server is running
    StartAndInitMcp
    
    '2. Pull the full tool list into a collection ----------------
    Dim cursor As String, reqId As Long, resp As Dictionary, tools As Collection
    Dim allTools As Collection: Set allTools = New Collection
    Do
        reqId = NextReqId()
        Dim jsonReq As String
        jsonReq = "{""jsonrpc"":""2.0"",""id"":" & reqId & _
                  ",""method"":""tools/list"",""params"":{"
        If Len(cursor) > 0 Then jsonReq = jsonReq & _
                  """cursor"":""" & Replace(cursor, """", "\""") & """" & ","
        jsonReq = jsonReq & "}}"
        gMcp.StdIn.WriteLine jsonReq
        
        Dim line As String
        Do
            DoEvents
            If Not gMcp.StdOut.AtEndOfStream Then
                line = gMcp.StdOut.ReadLine
                Set resp = ParseJson(line)
                If resp("id") = reqId Then Exit Do
            End If
        Loop
        
        If resp.Exists("error") Then
            DvMcpListTools = CVErr(xlErrNA)
            Exit Function
        End If
        
        Set tools = resp("result")("tools")
        Dim t As Variant: For Each t In tools: allTools.Add t: Next t
        
        If resp("result").Exists("nextCursor") Then
            cursor = resp("result")("nextCursor")
        Else
            cursor = ""
        End If
    Loop While Len(cursor) > 0
    
    '3. Copy into a 2-D variant array ---------------------------
    Dim r As Long, arr() As Variant
    ReDim arr(0 To allTools.Count, 0 To 2)      'header + n rows
    
    'header row
    arr(0, 0) = "Name": arr(0, 1) = "Title": arr(0, 2) = "Description"
    
    'data rows
    For r = 1 To allTools.Count
        arr(r, 0) = allTools(r)("name")
        arr(r, 1) = allTools(r)("title")
        arr(r, 2) = allTools(r)("description")
    Next r
    
    '4. Return—Excel will spill it
    DvMcpListTools = arr

End Function





