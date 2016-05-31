'======================================================'
'===============Lame VBScript things==================='
'======================================================'
OPTION EXPLICIT


' - # AND wont stop evaluating if 1st element is False
Sub AndExample
    Dim v1 : v1 = True
    Dim v2 : v2 = False
    Dim v3(10)
    If v2 AND v3(13)=1 Then
        MsgBox "True"
    Else
        MsgBox "False"
    End If
End Sub

' - > [AND wont stop evaluating if 1st element is False]
' - So you have to nest your Ifs in case their could be a possible error
Sub AndExample2
    Dim v1 : v1 = True
    Dim v2 : v2 = False
    Dim v3(10)
    If v2 Then 
        If (v3(13)=1) Then
            MsgBox "True"
        Else
            MsgBox "False"
        End If
    End If
End Sub

' - Concat a string multi line for readability without doing var = var &... 
Sub MultiLineConcatString
    Dim oStr : oStr = "" &_
        " 1 Line" &_
        " 2 Line"
    MsgBox oStr
End Sub

' - Using AND/OR to store multiple values limits
' - TBK most things you can store the exponent n = 30 before you get Overflow error
Sub LogicalOperatorLimits
    Dim v3 : v3 = 2^30          ' A Double Here
'            : v3 = 2^31         ' 31 will break(when AND applied below) VarType = Double
'            : v3 = 1           ' VarType = Integer
'            : v3 = 2^1         ' VarType = Double
'            : v3 = 2*2         ' VarType = Integer
'            : v3 = 2/2         ' VarType = Double
'            : v3 = SQR(3)      ' Vartype = Double
'            : v3 = v3 AND 2    ' VarType = Long
'            : v3 = v3 OR 2     ' VarType = Long
'            : v3 = v3 XOR 2    ' VarType = Long
'            : v3 = v3 * 2      ' VarType = Double
'            : v3 = SQR(-1)     ' Invalid argument Error
    MsgBox "Type = " & TypeName(v3) & ": " & v3
End Sub

' -1 is True as is the String
' 0 is False as is the String
Sub NegativeForTrue 
    If -1 = True THen 
        MsgBox "True"
    End If
    MsgBox "-1 = " & (CBool(-1))        ' Prints True
    MsgBox "0 = " & (CBool(0))          ' Prints False
    MsgBox "1 = " & (CBool(1))          ' Prints True
    MsgBox """1""= " & (CBool("1"))     ' Prints True
    MsgBox """0""= " & (CBool("0"))     ' Prints False
    MsgBox """-1""= " & (CBool("-1"))   ' Prints True
    'MsgBox """= " & (CBool(""))        ' TypeMissmatch Error 
    MsgBox "True = " & CInt(True)       ' Prints -1 
    MsgBox "False = " & CInt(False)     ' Prints 0
End Sub

' -- Auto Resiizng Aray
Sub sets(byRef arr ,byVal n , byVal v)
    If UBound(arr)<n Then
        Redim Preserve arr(n+100)
        arr(n) = v
    Else 
        arr(n) = v
    End If
End Sub

Sub puts(byRef arr, byVal n)
    Redim Preserve arr(UBound(arr)+1)
    arr(UBound(arr)) = n
End Sub

Sub C(byRef str, byVal str2)
    str = str & str2
End Sub


ReDim v(10)
v(1) = 1
v(2) = 3
v(3) = 5
Call sets(v,55,5)
MsgBox "55:" & v(55) & " 1:" & v(1) & " 2:" & v(2) & " 3:" & v(3)

Redim v2(5)
v2(0) = 10
v2(1) = 12
v2(2) = 10
v2(3) = 20
v2(4) = 30
v2(5) = 40
Dim str
C str , "Hi"
C str  , " Hello"
C str , " Hi"
C str , " Hiya"
MsgBox str

' Would Break v2(6) = 60
Call puts(v2,60)
Dim iV2
For iV2=0 To UBound(v2)
    MsgBox "v("&iV2&"): " & v2(iV2) 
Next

'###############################################'
' ############### MAIN TEST ################### '
'###############################################'
