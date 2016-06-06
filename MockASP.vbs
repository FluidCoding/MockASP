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
    'Constructor
    Private Sub Class_Initialize( )
        HTML = 0
        VB = 1
        STATE = HTML
        filePath = "C:\Users\User\My Coding\MockASP\"
    End Sub

    Private Sub Class_Terminate(  )
        'On Nothingd
    End Sub
    
    ' - Set the ASP Page to use
    Function SetPage(fn)
        fileName = fn
    End Function
    
    ' - Get Useable Value From User, For Form Action
    Function ToMethodVal(V)
        If V=vbYes Then : ToMethodVal="GET" : Else : ToMethodVal="POST"
    End Function
    
    ' - Process User Input State Parameters
    Function ProcessInput()
        Dim line : line = ""
        Dim vbPos : vbPos = 0
        Dim vbNeg : vbNeg = 0
        Dim devLimit : devLimit = 5
        Dim ParamI : ParamI = 0
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
        Do While fHdl.AtEndOfStream <> True 'AND devLimit > 1
            line = fHdl.ReadLine
            While ParamI < UBound(QStr)-1
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
        For Each Key in Parameters
            Parameters(Key) = InputBox("Enter value for " & Key & ": ", "Request Builder", Parameters(Key))
        Next
        Parameters.add "_METHOD", ToMethodVal(MsgBox("IS METHOD TYPE GET(Yes) OR POST(No): ", vbYesNo))
    End Function
    
    ' Separate Frontend from backend
    ' Insert variables into inline <%={var}%>
    ' Move Response.Writes to the Dom
    ' 
    Function LoadFile()
        Dim line : line = ""
        Dim vbPos : vbPos = 0
        Dim vbNeg : vbNeg = 0
        Dim devLimit : devLimit = 5
        vbStr = ""
        domStr = ""
        ' Mock it
        Set fSys = CreateObject("Scripting.FileSystemObject")
        Set fHdl = fSys.OpenTextFile(filePath & fileName)
        Do While fHdl.AtEndOfStream <> True 'AND devLimit > 1
            line = fHdl.ReadLine
          '  MsgBox "State: " & STATE &VbCrLf &_
           '         "Line: " & line 
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
' - 


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
Dim HTML
 HTML = "<!DOCTYPE html>" & VbCrLf &_
    "<html>" & VbCrLf &_
        "<head>" & VbCrLf &_
        "<title>test</title>" & VbCrLf &_
        "</head>"  & VbCrLf &_
        "<body>MokMokMok</body>" & VbCrLf &_
    "</html>"

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
MockASP.ProcessInput                            '| - Process User Input State Parameters
MockASP.LoadFile                                '| - Load the ASP File(TODO: And Dependency Includes)
MockASP.WriteToFile                             '| - Write Reulting HTML

