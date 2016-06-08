Class Mock
    Private Action
    Private STATE
    Private HTML
    Private VB
    Private domStr
    Private vbStr
    Private filePath
    Private Parameters
    Private fileName
    Private inputRequest
    'Constructor
    Private Sub Class_Initialize( )
        HTML = 0
        VB = 1
        STATE = HTML
        filePath = "C:\Users\User\My Coding\MockASP\"
        inputRequest = 0
    End Sub

    Private Sub Class_Terminate(  )
        'On Nothingd
    End Sub
    
    ' - GET/POST
    Function GetActionName
        If Parameters.Exists("_METHOD") Then : GetActionName = Parameters("_METHOD") : Else : GetActionName = "None"
    End Function
    
    ' - Set the ASP Page to use
    Function SetPage(fn)
        fileName = fn
    End Function
    
    ' - Get Useable Value From User, For Form Action
    Function ToMethodVal(V)
        If V=vbYes Then : ToMethodVal="GET" : Else : ToMethodVal="POST"
    End Function
    
    ' - Gather User Input As Web Form ("WEB"|"MSG"|Default:Execution Params)
    Function InitRequest(method)
        Dim Key, fSys, fStr
        Select Case method
            Case "WEB"  ' - Write a WebPage to gather the values into a request db file
                Dim HTML, reqFileName
                reqFileName = filePath & Replace(fileName, ".", "_") & "_Request.html"
                HTML = "<!DOCTYPE html>" & VbCrLf &_
                    "<html>" & VbCrLf &_
                        "<head>" & VbCrLf &_
                        "<title>Request Parameters Builder</title>" & VbCrLf &_
                        "<style>"  & "div{display:block;}" & VbCrLf &_
                        "</style></head>"  & VbCrLf &_
                        "<body>"
                    For Each Key In Parameters
                        HTML =HTML& "<div><label for=""" & Key & """ >"&Key&"</label><input type=""text"" name="""& Key & """ ></div>"  
                    Next
                HTML = HTML& "<label for""method"">Method: </label><input type=""radio"" name=""method"" value=""get"">GET</input><input type=""radio"" name=""method"" value=""post"" >POST</input>" &_
                            "<input type=""submit"" name=""submit"">Submit</input>"
                HTML =HTML& "</body>" & VbCrLf &_
                    "</html>"
                Set fSys = CreateObject("Scripting.FileSystemObject")
                Set fStr = fSys.CreateTextFile(reqFileName, 2)
                    fStr.Write HTML
                fStr.Close
                Set fStr = Nothing
                Set fSys = Nothing

                'Open It
                MsgBox reqFileName
                CreateObject("WScript.Shell").Run Chr(34) & reqFileName & Chr(34)

                ' Block for user input/save
                while inputRequest <> vbYes
                    inputRequest = MsgBox("Please fill out the form on "& fileName& "_Request.html, Then Press Yes to continue", vbYesNo)
                WEnd
            Case "MSG"
                For Each Key in Parameters
                    Parameters(Key) = InputBox("Enter value for " & Key & ": ", "Request Builder", Parameters(Key))
                Next
                Parameters.add "_METHOD", ToMethodVal(MsgBox("IS METHOD TYPE GET(Yes) OR POST(No): ", vbYesNo))
            Case Else

        End Select
    End Function 

    ' - Process User Input State Parameters
    Function ProcessInput()
        Dim line    : line = ""
        Dim vbPos   : vbPos = 0
        Dim vbNeg   : vbNeg = 0
        Dim devLimit : devLimit = 5
        Dim ParamI  : ParamI = 0
        Dim QStr(2) 
            QStr(0) = "Request.Querystring("""
            QStr(1) = "Request.Form("""
            QStr(2) = "Session("""
        vbStr = ""
        domStr = ""        
        Set Parameters = CreateObject("Scripting.Dictionary")
        Dim Key, K
        ' Read File Parse QS/F/Session Keys
        Set fSys = CreateObject("Scripting.FileSystemObject")
        Set fHdl = fSys.OpenTextFile(filePath & fileName)
        Do While fHdl.AtEndOfStream <> True
            line = fHdl.ReadLine
            While ParamI < UBound(QStr)
                vbPos = InStr(1,line,QStr(ParamI),1)
                If vbPos > 0 Then
                    vbNeg = InStr(vbPos + Len(QStr(ParamI)), line, """", vbTextCompare)
                    K = Mid(line, vbPos+len(QStr(ParamI)), vbNeg-vbPos-len(QStr(ParamI)))
                    If Parameters.Exists(K)=False Then Parameters.add K, "Empty"
                End If
                ParamI = ParamI + 1
            Wend
            ParamI = 0
        Loop

    End Function
    
    ' Separate Frontend from backend
    ' Insert variables into inline <%={var}%>
    ' Move Response.Writes to the Dom
    ' 
    Function LoadFile()
    Dim inputFlag
    If inputRequest=0 Then

    End If
        Dim line            : line = ""
        Dim vbPos           : vbPos = 0
        Dim vbNeg           : vbNeg = 0
        Dim vbPos2          : vbPos2 = 0
        Dim vbNeg2          : vbNeg2 = 0
        Dim devLimit        : devLimit = 5
        Dim Action          : Action = GetActionName
        Dim paramReplace    : paramReplace = ""
        Dim ParamI          : ParamI = 0
        Dim QStr(2) 
            QStr(0) = "Request.Querystring("""
            QStr(1) = "Request.Form("""
            QStr(2) = "Session("""
        If Action = "GET" Then : paramReplace = "Request.QueryString(""" : Else : paramReplace = "Request.Form"
        vbStr = ""
        domStr = ""
        ' Mock it
        Set fSys = CreateObject("Scripting.FileSystemObject")
        Set fHdl = fSys.OpenTextFile(filePath & fileName)
        
        ' Processing Steps: 
                            ' 1 Read Line
                            ' 2 Replace Request/Session Variables
                                '(EmptyString or Nothing to the unset Vars)
                                ' Go more than 1 per line deep
                            ' 3 Parse Writes
                            ' 4 Execute VBs inline 
        Dim val : val = """"
        Do While fHdl.AtEndOfStream <> True
            line = fHdl.ReadLine
          '  MsgBox "State: " & STATE &VbCrLf &_
           '         "Line: " & line 
           ' - Inline Replace Params
            While ParamI < UBound(QStr)
                If QStr(ParamI) = paramReplace Then : val = Parameters(K) : Else : val = """"
                vbPos2 = InStr(1,line,QStr(ParamI),1)
                If vbPos2>0 Then
                    vbNeg2 = InStr(vbPos2 + Len(QStr(ParamI)), line, """", vbTextCompare)
                    K = Mid(line, vbPos2+len(QStr(ParamI)), vbNeg2-vbPos2-len(QStr(ParamI)))
                    If Parameters.Exists(K)=True Then
                        line = Left(line, vbPos2-1) & """" & Parameters(K) & """" & Right(line, len(line) - vbNeg2-1)
                        'MsgBox line
                    End If
                End If
                ParamI = ParamI + 1
            Wend
            ParamI = 0

           ' - Parse 
            If STATE = HTML Then
                vbPos = InStr(line, "<%")
                If vbPos > 0 Then  ' Theres an open VB Tag to parse...TODO: ensure the surrounding html is grabbed to domStr
                    STATE = VB
                    vbNeg = InStr(line, "%>")-3
                    If vbNeg > 0 Then  ' The Script tag was closed
                       ' MsgBox "44 vbNeg: " & vbNeg & " vbPos: " & vbPos
                        vbStr = vbStr & Mid(line, vbPos+2, Len(line) - (Len(line)-vbNeg)) & VbCrLf
                        STATE = HTML
                    Else 
                       ' MsgBox "48 vbNeg: " & vbNeg & " vbPos: " & vbPos
                        vbStr = vbStr & Right(line, Len(line)-vbPos+2) & VbCrLf
                    End If
                Else
                    domStr = domStr & line & VbCrLf
                End If
            ElseIf STATE = VB Then
                vbPos = InStr(line, "%>")
                If vbPos > 0 Then
                   ' MsgBox "57 vbNeg: " & vbNeg & " vbPos: " & vbPos
                    STATE = HTML
                    vbStr = vbStr & Right(line, Len(line)-vbPos) & VbCrLf
                Else
                    vbStr = vbStr & line & VbCrLf
                End If
                
            End If
            devLimit = devLimit-1
        Loop
    MsgBox "VB: " & vbStr
    MsgBox "Dom: " & domStr
    Set fSys = Nothing : Set fHdl = Nothing
    End Function
    
    Sub WriteToFile
        Set fSys = CreateObject("Scripting.FileSystemObject")
        Set fStr = fSys.CreateTextFile(filePath&"MockFile.html", 2)
            fStr.Write domStr
        fStr.Close
        Set fStr = Nothing
        
        Set fStr = fSys.CreateTextFile(filePath&"MockFilex.vbs", 2)
            fStr.Write vbStr
        fStr.Close
        Set fStr = Nothing
        
        Set fSys = Nothing
    End Sub
End Class
' --==============END Mock Class==============-- 


' -
Class ResponseMock
    Private fileName
    Private filePath
    private domStr
    Private Sub Class_Initialize
        fileName = "mock.html"
        filePath = "C:\Users\User\My Coding\MockASP"
        domStr = ""
    End Sub

    Function Msg (data)
        MsgBox data
    End Function

' - Mock Writes to Virtual DOM
    Function Write(data)
        domStr = domStr & data
    End Function

' - Writes The Virtual MockDOM to MockFile
    Function Flush
        Set fSys = CreateObject("Scripting.FileSystemObject")
        Set fStr = fSys.CreateTextFile(filePath&fileName, 2)
            fStr.Write domStr
        fStr.Close
        Set fStr = Nothing
        Set fSys = Nothing
    End Function

' - Opens MockFile in default browser
    Function Open
        CreateObject("WScript.Shell").Run filePath&fileName
    End Function

' - JS Style Console log
    Function Log(var)
        Write("<script>console.log(" & var & ");</script>")
    End Function

' - Starts Mocking... (Writes Response & Opens)
    Function Mock
        Flush
        Open
    End Function
End Class
' ----- END Response CLASS -----'
' TODO: Request FNC
    ' -> QueryString
    ' -> Form
Class Request
    Private Sub Class_Initialize( )
        'Constructor

    End Sub

    Private Sub Class_Terminate(  )
        'On Nothingd
    End Sub

' - Mock Query String Actions
    Function QueryString(key)
    End Function
' - Mock Form INPUT/Reading'
    Function Form(key)

        'body
    End Function
    
        Private Requests
    Function ToMethodVal(V)
        If V=vbYes Then : ToMethodVal="GET" : Else : ToMethodVal="POST"
    End Function
    
    Function GetParameters(Ps, P)
        Set Requests = CreateObject("Scripting.Dictionary")
        Dim Key, K
        ' Santitize Fill into Dictionary
        For Each K IN Ps
            If Requests.Exists(K)=False Then Requests.add K, "Empty"
        Next

        For Each Key in Requests
            Requests(Key) = InputBox("Enter value for " & Key & ": ", "Request Builder", Requests(Key))
        Next
        Requests.add "_METHOD", ToMethodVal(MsgBox("IS METHOD TYPE GET(Yes) OR POST(No): ", vbYesNo))
        ' Print That Back
        Dim reqStr : reqStr = "" 
        For Each Key in Requests
            reqStr = reqStr & "[" & Key & " := " & Requests(Key) & "]" & vbCrLf
        Next
        MsgBox reqStr, " Request Parameters "
        Set P = Requests
        GetParameters = reqStr
    End Function
End Class

' ----- END Request CLASS -----'
MsgBox " Mocking..."
Set Response = new ResponseMock
' Play Data '


' Sample Action'
'Response.Write HTML
'Response.Mock
' ======================================================'
' ======================== MAIN ========================'
' ======================================================'
' Processing Steps: 
            ' [x] 1 Load ASP File
            ' [x] 2 Read Through For Form/QueryString/Session Variables Expected
            ' [x] 3 Prompt User For values or use file(save to file after prompt for next time)
            ' [] 4 Parse File Replace Input Params with last steps
            ' [] 5 Get Mock DB Input
            ' [] 6 Generate HTML+Execute VB
            ' [x] 7 Write HTML
            ' [x] 8 OPEN in Browser

' Wrap ASP
Set MockASP = new Mock                          '| - Initialize a Mock ASP Object
MockASP.SetPage("Page.ASP")                     '| - Set the Working ASP File to Mock
MockASP.ProcessInput                            '| - Process ASP State Parameters
MockASP.initRequest("WEB")                      '| - Gather User Input As Web Form ("WEB"|"MSG"|Default:Execution Params)
MockASP.LoadFile                                '| - Load the ASP File(TODO: And Dependency Includes)
MockASP.WriteToFile                             '| - Write Reulting HTML

