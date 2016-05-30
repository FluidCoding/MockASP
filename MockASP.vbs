Class Mock
    Private Action
    Private STATE
    Private HTML
    Private VB
    Private domStr
    Private vbStr
    Private filePath

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
    
    '
    ' Separate Frontend from backend
    ' Insert variables into inline <%={var}%>
    ' Move Response.Writes to the Dom
    ' 
    Function LoadFile(fileName)
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
End Class

' ----- END Request CLASS -----'
MsgBox " Mocking..."
Set Response = new ResponseMock
' Play Data '
' TODO: Load via asp File'
Dim HTML
 HTML = "<!DOCTYPE html>" & VbCrLf &_
    "<html>" & VbCrLf &_
        "<head>" & VbCrLf &_
        "<title>test</title>" & VbCrLf &_
        "</head>"  & VbCrLf &_
        "<body>MokMokMok</body>" & VbCrLf &_
    "</html>"

' The Action'
'Response.Write HTML
'Response.Mock

' Wrap ASP
Set MockASP = new Mock
MockASP.LoadFile("Page.ASP")
MockASP.WriteToFile