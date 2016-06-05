' Select Cases can be used to short - circuit to prevent errors like 4/0 
OPTION EXPLICIT
DIM V : V = TRUE
SELECT CASE V 
    CASE TRUE
    CASE Val
        MsgBox "False"
END SELECT
Dim x
Redim y(1)
Dim tStart, tStop1
Dim d : Set d=CreateObject("Scripting.Dictionary")
' ==== Basic For Loop
' tStart = Timer
' For x=0 To 100000
' Next
' tStop1 = Timer
' MsgBox "tStart: " & tStart & " tStop: " & tStop1 & " Diff: " & Round((tStop1 - tStart),4)
' TIME = .006

' ===== Addingg 100,000 Elements to a dictionary with mult
' tStart = Timer
' For x=0 To 100000
'     d.add "F("&x&")", x*2
' Next
' tStop1 = Timer
' MsgBox "tStart: " & tStart & " tStop: " & tStop1 & " Diff: " & Round((tStop1 - tStart),4)
' msgBox d.count  ' About 0.7 Seconds

' ==== Adding 100,001 Items to an array with continuous redim
' tStart = Timer
' For x=0 To 100000
'     Redim y(x)
'     y(x) = "F("&x&")=" & x*2
' Next
' tStop1 = Timer
' MsgBox "tStart: " & tStart & " tStop: " & tStop1 & " Diff: " & Round((tStop1 - tStart),4)
' ' 101 Seconrds
' msgBox UBound(y) 
' ==== Calling UBound
' tStart = Timer
' For x=0 To 100000
'      UBound(y)
' Next
' tStop1 = Timer
' MsgBox "tStart: " & tStart & " tStop: " & tStop1 & " Diff: " & Round((tStop1 - tStart),4)
' Time = .04
' ==== Incremental Redim
' tStart = Timer
' For x=0 To 100000
'      If x = UBound(y) Then Redim y(x+1)
'      y(x) = x
' Next 
' tStop1 = Timer
' MsgBox "tStart: " & tStart & " tStop: " & tStop1 & " Diff: " & Round((tStop1 - tStart),4)
' Time = 101.063

' ==== Incremental Redim*2
' tStart = Timer
' For x=0 To 100000
'      If x = UBound(y) Then Redim y(x*2)
'      y(x) = x
' Next 
' tStop1 = Timer
' MsgBox "tStart: " & tStart & " tStop: " & tStop1 & " Diff: " & Round((tStop1 - tStart),4)
' MsgBox "New UBound " & UBound(y)
' Time = .08

' ==== Incremental Redim*2 With Preserve
' tStart = Timer
' For x=0 To 100000
'      If x = UBound(y) Then Redim Preserve y(x+1)
'      y(x) = x
' Next 
' tStop1 = Timer
' MsgBox "tStart: " & tStart & " tStop: " & tStop1 & " Diff: " & Round((tStop1 - tStart),4)
' Time = .53

' ==== Incremental Redim With Preserve\
' tStart = Timer
' For x=0 To 100000
'      If x = UBound(y) Then Redim Preserve y(x*2)
'      y(x) = x
' Next 
' tStop1 = Timer
' MsgBox "???tStart: " & tStart & " tStop: " & tStop1 & " Diff: " & Round((tStop1 - tStart),4)
' Time = .53?!


' ==== 100 Redim
' tStart = Timer
' For x=0 To 100
'     Redim y(10000000)
'     y(x) = x    
' Next
' tStop1 = Timer
' MsgBox "???tStart: " & tStart & " tStop: " & tStop1 & " Diff: " & Round((tStop1 - tStart),4)
' MsgBox y(4)
' Time = 22
' ==== 100 Redim Preserve
' tStart = Timer
' Redim y(0,0)
' For x=0 To 1000
'     Redim y(x,x)
'     y(x,x) = x
' Next
' tStop1 = Timer
' MsgBox "???tStart: " & tStart & " tStop: " & tStop1 & " Diff: " & Round((tStop1 - tStart),4)
' Time = 7.5


Sub psh(byRef a, byVal v)
    Redim Preserve a(UBound(a)+1)
    a(UBound(a))=v
End Sub

Sub psh2(byRef a, byVal v) 
    Dim enda : enda = UBound(a)
    Redim Preserve a(UBound(a)*2)
    a(enda) = v
End Sub


'=========== Loading a Dictionary and Making Copies
Function FormIn(O)
    Dim formi : Set formi=CreateObject("Scripting.Dictionary")
    For Each x In O
        formi.add x, O(x)
    Next
    Set FormIn = formi
End Function

d.add "v", 2
d.add "2v", 22
Dim F : Set F = FormIn(d)
Dim F2 : Set F2 = FormIn(d)
Set d = Nothing

For Each x in F2.keys
    MsgBox F2(x)
Next

For Each x in F2.keys
    F2(x) = F2(x)+4
Next

For Each x in F2.keys
    MsgBox F2(x)
Next

For Each x in F.keys
    MsgBox F(x)
Next
MsgBox TypeName(d)

' ReDim testManDim(10)
' testManDim(1) = 1
' testManDim(2) = 2
' testManDim(5) = 5
' testManDim(8) = 8
' testManDim(10) = 10
' psh testManDim, 99
' psh testManDim, 11
' Dim sOut : sOut = ""
' For Each v IN testManDim : sout=sout&v :  Next

' MsgBox sout
' sout =""
' MsgBox UBound(testManDim)
' psh2 testManDim, 1
' MsgBox UBound(testManDim)
' Time = 0.11

' tStart = Timer
' For x=0 To 100000
'     IF UBound(y) < x Then 
'         Redim y(x)
'     End If
'     y(x) = "F("&x&")=" & x*2
' Next
' tStop1 = Timer
' MsgBox "tStart: " & tStart & " tStop: " & tStop1 & " Diff: " & Round((tStop1 - tStart),4)
' ' 101.241 Seconrds
' msgBox UBound(y) 
