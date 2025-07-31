Attribute VB_Name = "Eval"
' Math Interpreter 2.0
' Shunting-Yard Math Algorithm
' Module by Arfur (31/07/2025)
' Github: https://github.com/therealarfu

Option Explicit

Private Const MATH_CHARS As String = "0123456789.+-*/^()"
Private Const DIGITS As String = "0123456789"
Private Const OPERATORS As String = "+-*/^()"

Private Function Lexer(ByVal expr As String) As String()
    Dim newexpr As String, length As Long, i As Long, arr() As String, Char As String, tempnumber As String, ArrIndex As Long, lbar As Long, rbar As Long
    newexpr = Replace(expr, " ", "")
    length = Len(newexpr)
    ReDim arr(0 To length * 2) As String

    If InStr(1, DIGITS & "(+-.", Left$(newexpr, 1)) = 0 Then RaiseError "Invalid start: """ & Left$(newexpr, 1) & """.", 1
    If InStr(1, DIGITS & ").", right$(newexpr, 1)) = 0 Then RaiseError "Invalid end: """ & right$(newexpr, 1) & """.", length

    For i = 1 To length
        Char = Mid$(newexpr, i, 1)
        If InStr(1, MATH_CHARS, Char) = 0 Then RaiseError "Invalid character: """ & Char & """.", i
        If ArrIndex > UBound(arr) Then ReDim Preserve arr(UBound(arr) * 2 + 1) As String
        Select Case Char
            Case "(": lbar = lbar + 1
            Case ")": rbar = rbar + 1
        End Select
        
        If IsNumeric(Char) Then
            tempnumber = tempnumber & Char
        ElseIf Char = "." Then
            If InStr(1, tempnumber, ".") <> 0 Then RaiseError "Float with multiple dots", i
            tempnumber = tempnumber & "."
        Else
            If tempnumber <> "" Then
                arr(ArrIndex) = tempnumber
                ArrIndex = ArrIndex + 1
                tempnumber = ""
            End If
            arr(ArrIndex) = Char
            If ArrIndex = 0 Then
                Select Case Char
                    Case "-"
                        Char = "~"
                        arr(ArrIndex) = Char
                    Case "+": ArrIndex = ArrIndex - 1
                End Select
            ElseIf InStr(1, "+-~*/^(", arr(ArrIndex - 1)) <> 0 Then
                Select Case Char
                    Case "-"
                        Char = "~"
                        arr(ArrIndex) = Char
                    Case "+": ArrIndex = ArrIndex - 1
                End Select
            End If
            ArrIndex = ArrIndex + 1
        End If
    Next
    
    If lbar <> rbar Then RaiseError "Invalid brackets."
    
    If tempnumber <> "" Then
        If ArrIndex > UBound(arr) Then ReDim Preserve arr(0 To ArrIndex)
        arr(ArrIndex) = tempnumber
        ArrIndex = ArrIndex + 1
    End If

    ReDim Preserve arr(0 To ArrIndex - 1)
    Lexer = arr
End Function

Private Function Precedence(ByVal Char As String) As Byte
    Select Case Char
        Case "+", "-": Precedence = 1
        Case "*", "/": Precedence = 2
        Case "^": Precedence = 3
        Case "~": Precedence = 4
    End Select
End Function

Private Function IsRightAssociative(ByVal Char As String) As Boolean
    IsRightAssociative = (Char = "^" Or Char = "~")
End Function

Private Function Parser(expr() As String) As String()
    Dim OperatorList() As String, Exitlist() As String, i As Long, OpIndex As Long, ExIndex As Long, Char As String, j As Long
    
    ReDim OperatorList(0 To 8) As String
    ReDim Exitlist(0 To 8) As String
    
    For i = 0 To UBound(expr)
        Char = expr(i)
        If OpIndex > UBound(OperatorList) Then ReDim Preserve OperatorList(UBound(OperatorList) * 2 + 1) As String
        If ExIndex > UBound(Exitlist) Then ReDim Preserve Exitlist(UBound(Exitlist) * 2 + 1) As String
        
        If IsNumeric(Char) Then
            Exitlist(ExIndex) = Char
            ExIndex = ExIndex + 1
        Else
            If OpIndex = 0 Or Char = "(" Then
                OperatorList(OpIndex) = Char
                OpIndex = OpIndex + 1
            Else
                If Char = ")" Then
                    For j = OpIndex - 1 To 0 Step -1
                        If OperatorList(j) = "(" Then
                            OpIndex = OpIndex - 1
                            Exit For
                        End If
                        Exitlist(ExIndex) = OperatorList(j)
                        OpIndex = OpIndex - 1
                        ExIndex = ExIndex + 1
                    Next
                ElseIf OperatorList(OpIndex - 1) = "(" Or Precedence(Char) > Precedence(OperatorList(OpIndex - 1)) Or (Precedence(Char) = Precedence(OperatorList(OpIndex - 1)) And IsRightAssociative(Char)) Then
                    OperatorList(OpIndex) = Char
                    OpIndex = OpIndex + 1
                Else
                    Do While OpIndex > 0
                        If OperatorList(OpIndex - 1) <> "(" And Precedence(Char) <= Precedence(OperatorList(OpIndex - 1)) And Not IsRightAssociative(Char) Then
                            Exitlist(ExIndex) = OperatorList(OpIndex - 1)
                            ExIndex = ExIndex + 1
                            OpIndex = OpIndex - 1
                        Else
                            Exit Do
                        End If
                    Loop
                    OperatorList(OpIndex) = Char
                    OpIndex = OpIndex + 1
                End If
            End If
        End If
    Next
    
    For i = OpIndex - 1 To 0 Step -1
        If ExIndex > UBound(Exitlist) Then ReDim Preserve Exitlist(UBound(Exitlist) * 2 + 1) As String
        Exitlist(ExIndex) = OperatorList(i)
        ExIndex = ExIndex + 1
    Next
    
    ReDim Preserve Exitlist(0 To ExIndex - 1)
    Parser = Exitlist
End Function

Private Function Calc(ByVal x As String, ByVal OPERATOR As String, ByVal y As String) As String
    Dim num1 As Double, num2 As Double
    num1 = CDbl(Replace(x, ".", ","))
    num2 = CDbl(Replace(y, ".", ","))
    
    Select Case OPERATOR
        Case "+"
            Calc = CStr(num1 + num2)
        Case "-"
            Calc = CStr(num1 - num2)
        Case "*"
            Calc = CStr(num1 * num2)
        Case "/"
            If num2 = 0 Then RaiseError "Division by zero."
            Calc = CStr(num1 / num2)
        Case "^"
            If num1 < 0 And num2 < 1 And num2 > 0 Then RaiseError "Root of negative number."
            If num1 = num2 And num1 = 0 Then RaiseError "Undefined: 0 ^ 0."
            Calc = CStr(num1 ^ num2)
    End Select
    
    Calc = Replace(Calc, ",", ".")
End Function

Private Function Interpreter(expr() As String) As String
    Dim ValueList() As String, i As Long, Char As String, ValueIndex As Long
    
    ReDim ValueList(0 To 8) As String
    
    For i = 0 To UBound(expr)
        Char = expr(i)
        If ValueIndex > UBound(ValueList) Then ReDim Preserve ValueList(UBound(ValueList) * 2 + 1) As String
        
        If IsNumeric(Char) Then
            ValueList(ValueIndex) = Char
            ValueIndex = ValueIndex + 1
        Else
            If Char = "~" Then
                ValueList(ValueIndex - 1) = Calc(ValueList(ValueIndex - 1), "*", -1)
            Else
                ValueList(ValueIndex - 2) = Calc(ValueList(ValueIndex - 2), Char, ValueList(ValueIndex - 1))
                ValueIndex = ValueIndex - 1
            End If
        End If
    Next
    Interpreter = ValueList(0)
End Function

Public Function Evaluate(ByVal Expression As String) As String
    Evaluate = Interpreter(Parser(Lexer(Expression)))
End Function

Private Function RaiseError(ByVal Message As String, Optional ByVal Index As Long = -1)
    If Index <> -1 Then
        Err.Raise vbObjectError, "Eval", "Error at index " & Index & ": " & Message
    Else
        Err.Raise vbObjectError, "Eval", Message
    End If
End Function
