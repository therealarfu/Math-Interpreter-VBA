Attribute VB_Name = "Eval"
' Math Interpreter 1.0
' Module by Arfur (12/12/2024)
' Github: https://github.com/therealarfu

Option Explicit

Private Const MATH_CHARS As String = "0123456789+-/*^()."
Private Const DIGITS As String = "0123456789"
Private Const OPERATORS As String = "+-/*()^"

Private Sub AddItem(Arr As Variant, Item As Variant)
    ReDim Preserve Arr(UBound(Arr) + 1)
    Arr(UBound(Arr)) = Item
End Sub

Private Function IndexOf(List As Variant, Item As String) As Long
    Dim i As Long
    For i = 0 To UBound(List)
        If List(i)(0) = Item Then
            IndexOf = i
            Exit Function
        Else
            IndexOf = -1
        End If
    Next
End Function
 
Private Sub Pop(List As Variant, Optional ByVal Index As Long)
    Dim i As Long
    If IsMissing(Index) Or Index > UBound(List) Or UBound(List) = Index And UBound(List) = 1 Then
        ReDim Preserve List(UBound(List) - 1)
    ElseIf UBound(List) > Index Then
        For i = Index To UBound(List)
            If i <> UBound(List) Then
                List(i) = List(i + 1)
            End If
        Next
        ReDim Preserve List(UBound(List) - 1)
    ElseIf UBound(List) >= Index And UBound(List) = 0 Then
        ReDim List(UBound(List) - UBound(List))
    Else
        Exit Sub
    End If
End Sub

Private Function FindBrackets(Arr As Variant) As Variant
    Dim i As Long, lb As Long, rb As Long
    
    For i = UBound(Arr) To 0 Step -1
        If Arr(i)(0) = ")" Then
            rb = i
        ElseIf Arr(i)(0) = "(" Then
            lb = i
            Exit For
        End If
    Next
    If rb = lb Then
        FindBrackets = Array(-1, -1)
    Else
        FindBrackets = Array(lb, rb)
    End If
End Function

Private Function CheckString(ByVal Expr As String) As String
    Dim newexpr As String, length As Long, i As Long
    newexpr = Replace(Expr, " ", "")
    length = Len(newexpr)

    For i = 1 To length
        If InStr(1, MATH_CHARS, Mid$(newexpr, i, 1)) = 0 Then RaiseError "Invalid character: """ & Mid$(newexpr, i, 1) & """.", i
    Next
    
    If InStr(1, DIGITS & "(+-.", Left$(newexpr, 1)) = 0 Then RaiseError "Invalid start: """ & Left$(newexpr, 1) & """.", 1
    If InStr(1, DIGITS & ").", Right$(newexpr, 1)) = 0 Then RaiseError "Invalid end: """ & Right$(newexpr, 1) & """.", length
    If Len(Replace(newexpr, "(", "")) <> Len(Replace(newexpr, ")", "")) Then RaiseError "Invalid brackets."
    
    CheckString = newexpr
End Function

Private Function Tokenize(ByVal Expr As String) As Variant
    Dim temparr() As Variant, char As String, i As Long, tempnumber As String, newexpr As String
    newexpr = CheckString(Expr)
    temparr = Array()
    
    For i = 1 To Len(newexpr)
        char = Mid$(newexpr, i, 1)
        If InStr(1, DIGITS & ".", char) Then
            If char = "." And InStr(1, tempnumber, ".") <> 0 Then RaiseError "Float with multiple dots", i
            tempnumber = tempnumber & char
        Else
            If tempnumber <> "" Then
                If tempnumber = "." Then RaiseError "Invalid dot", i
                If Left$(tempnumber, 1) = "." Then
                    tempnumber = "0" & tempnumber
                ElseIf Right$(tempnumber, 1) = "." Then
                    tempnumber = Replace(tempnumber, ".", "")
                End If
                AddItem temparr, Array(tempnumber, "NUMBER")
                tempnumber = ""
            End If
            
            ReDim Preserve temparr(UBound(temparr) + 1)
            If InStr(1, "+-*/^", char) <> 0 Then
                temparr(UBound(temparr)) = Array(char, "OPERATOR")
            ElseIf char = "(" Then
                temparr(UBound(temparr)) = Array(char, "LBRACKET")
            Else
                temparr(UBound(temparr)) = Array(char, "RBRACKET")
            End If
        End If
    Next
    If tempnumber <> "" Then
        If tempnumber = "." Then RaiseError "Invalid dot", i
        If Left$(tempnumber, 1) = "." Then
            tempnumber = "0" & tempnumber
        ElseIf Right$(tempnumber, 1) = "." Then
            tempnumber = Replace(tempnumber, ".", "")
        End If
        AddItem temparr, Array(tempnumber, "NUMBER")
        tempnumber = ""
    End If
    
    Tokenize = temparr
End Function

Private Function CheckTokens(ByVal Expr As String) As Variant
    Dim temparr() As Variant, i As Long, v0 As String, v1 As String
    temparr = Tokenize(Expr)
    
    For i = 0 To UBound(temparr)
        Select Case temparr(i)(1)
            Case "NUMBER"
                If i < UBound(temparr) Then
                    v1 = temparr(i + 1)(1)
                    If v1 <> "RBRACKET" And v1 <> "OPERATOR" Then RaiseError "Invalid number: """ & temparr(i)(0) & """.", i + 1
                End If
            Case "OPERATOR"
                v0 = temparr(i + 1)(0)
                v1 = temparr(i + 1)(1)
                If v1 <> "NUMBER" And v1 <> "LBRACKET" And v0 <> "+" And v0 <> "-" Then RaiseError "Invalid operator: """ & temparr(i)(0) & """.", i + 1
            Case "LBRACKET"
                v0 = temparr(i + 1)(0)
                v1 = temparr(i + 1)(1)
                If v1 <> "NUMBER" And v1 <> "LBRACKET" And v0 <> "+" And v0 <> "-" Then RaiseError "Invalid left bracket.", i + 1
            Case "RBRACKET"
                If i < UBound(temparr) Then
                    v1 = temparr(i + 1)(1)
                    If v1 <> "RBRACKET" And v1 <> "OPERATOR" Then RaiseError "Invalid right bracket.", i + 1
                End If
        End Select
    Next
    
    CheckTokens = temparr
End Function


Private Function Parser(ByVal Expr As String) As Variant
    Dim temparr() As Variant, i As Long, temparr2() As Variant, signal As String
    temparr = CheckTokens(Expr)
    temparr2 = Array()
    
    For i = 0 To UBound(temparr)
        If temparr(i)(0) = "-" Or temparr(i)(0) = "+" Then
            If signal <> "" Then
                If signal = temparr(i)(0) Then
                    signal = "+"
                Else
                    signal = "-"
                End If
                If temparr(i + 1)(0) <> "-" And temparr(i + 1)(0) <> "+" Then
                    AddItem temparr2, Array(signal, "OPERATOR")
                    signal = ""
                End If
            ElseIf signal = "" And temparr(i + 1)(0) <> "-" And temparr(i + 1)(0) <> "+" Then
                AddItem temparr2, temparr(i)
            ElseIf signal = "" And (temparr(i + 1)(0) = "-" Or temparr(i + 1)(0) = "+") Then
                signal = temparr(i)(0)
            End If
        Else
            AddItem temparr2, temparr(i)
        End If
    Next
    
    Parser = NegativeNumbers(temparr2)
End Function

Private Function NegativeNumbers(temparr2 As Variant) As Variant
    Dim temparr As Variant, i As Long, char As String, nextchar As String
    temparr = Array()
    
    For i = 0 To UBound(temparr2)
        If temparr2(i)(0) = "+" Or temparr2(i)(0) = "-" Then
            If i = 0 Then
            
                If temparr2(1)(1) = "NUMBER" Then
                    
                    char = temparr2(0)(0)
                    nextchar = temparr2(1)(0)
                    
                    If InStr(1, "+-", Left$(nextchar, 1)) = 0 Then
                        temparr2(1)(0) = char & nextchar
                    ElseIf Left$(nextchar, 1) <> char And InStr(1, "+-", nextchar, 1) Then
                        temparr2(1)(0) = "-" & Right(nextchar, Len(nextchar) - 1)
                    ElseIf Left$(nextchar, 1) = char Then
                        temparr2(1)(0) = "+" & Right(nextchar, Len(nextchar) - 1)
                    End If
                    
                Else
                    AddItem temparr, temparr2(i)
                End If
            Else
                If (temparr2(i - 1)(1) = "OPERATOR" Or temparr2(i - 1)(1) = "LBRACKET") And temparr2(i + 1)(1) = "NUMBER" Then
                    char = temparr2(i)(0)
                    nextchar = temparr2(i + 1)(0)
                    
                    If InStr(1, "+-", Left$(nextchar, 1)) = 0 Then
                        temparr2(i + 1)(0) = char & nextchar
                    ElseIf Left$(nextchar, 1) <> char And InStr(1, "+-", Left$(nextchar, 1)) Then
                        temparr2(i + 1)(0) = "-" & Right(temparr2(i + 1)(0), Len(temparr2(i + 1)(0)) - 1)
                    ElseIf Left$(nextchar, 1) = char Then
                        temparr2(i + 1)(0) = "+" & Right(nextchar, Len(nextchar) - 1)
                    End If
                    
                Else
                    AddItem temparr, temparr2(i)
                End If
            End If
        Else
            AddItem temparr, temparr2(i)
        End If
    Next
    
    NegativeNumbers = temparr
End Function

Private Function Calc(ByVal X As String, ByVal OPERATOR As String, ByVal y As String) As String
    Dim num1 As Double, num2 As Double
    num1 = CDbl(Replace(X, ".", ","))
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

Private Function CalcExpr(Arr As Variant) As Variant
    Dim temparr As Variant, i As Long, OPERATORS As Long, pos1 As Long, pos2 As Long, op As Long, result As String
    
    For i = 0 To UBound(Arr)
        If Arr(i)(1) = "OPERATOR" Then OPERATORS = OPERATORS + 1
    Next
    
    If OPERATORS = 0 Then
        CalcExpr = Arr
        Exit Function
    End If
    
    For i = 0 To OPERATORS - 1
    
        pos1 = IndexOf(Arr, "^")
        If pos1 <> -1 Then
            op = pos1
        Else
            pos1 = IndexOf(Arr, "*")
            pos2 = IndexOf(Arr, "/")
            If pos1 <> -1 Or pos2 <> -1 Then
                If (pos1 < pos2 And pos1 <> -1) Or pos2 = -1 Then
                    op = pos1
                ElseIf (pos2 < pos1 And pos2 <> -1) Or pos1 = -1 Then
                    op = pos2
                End If
            Else
                pos1 = IndexOf(Arr, "+")
                pos2 = IndexOf(Arr, "-")
                If (pos1 < pos2 And pos1 <> -1) Or pos2 = -1 Then
                    op = pos1
                ElseIf (pos2 < pos1 And pos2 <> -1) Or pos1 = -1 Then
                    op = pos2
                End If
            End If
        End If
        
        result = Calc(Arr(op - 1)(0), Arr(op)(0), Arr(op + 1)(0))
        Pop Arr, op
        Pop Arr, op
        Arr(op - 1)(0) = result
    Next
    
    If IsArray(Arr(0)) Then
        CalcExpr = Arr(0)
    Else
        CalcExpr = Arr
    End If
End Function

Private Function Interpreter(ByVal Expr As String) As Variant
    Dim temparr As Variant, tempexpr As Variant, i As Long, j As Long, BCount As Long, lb As Long, rb As Long, result As Variant
    temparr = Parser(Expr)
    
    For i = 0 To UBound(temparr)
        If temparr(i)(0) = "(" Then BCount = BCount + 1
    Next
    
    For i = 0 To BCount - 1
        lb = FindBrackets(temparr)(0)
        rb = FindBrackets(temparr)(1)
        
        tempexpr = Array()
        For j = lb + 1 To rb - 1
            AddItem tempexpr, temparr(j)
        Next
        result = CalcExpr(tempexpr)
        For j = lb + 1 To rb
            Pop temparr, lb + 1
        Next
        If IsArray(result(0)) Then
            temparr(lb) = result(0)
        Else
            temparr(lb) = result
        End If
        temparr = NegativeNumbers(temparr)
    Next
    
    result = CalcExpr(temparr)
    If IsArray(result(0)) Then
        result(0)(0) = Replace(result(0)(0), "+", "")
        If Right$(result(0)(0), 1) = "0" Then result(0)(0) = Replace(result(0)(0), "-", "")
        Interpreter = result(0)
    Else
        result(0) = Replace(result(0), "+", "")
        If Right$(result(0), 1) = "0" Then result(0) = Replace(result(0), "-", "")
        Interpreter = result
    End If
End Function

Public Function Evaluate(ByVal Expr As String) As String
    Evaluate = Interpreter(Expr)(0)
End Function
    
Private Function RaiseError(ByVal Message As String, Optional ByVal Index As Long = -1)
    If Index <> -1 Then
        Err.Raise vbObjectError, "Eval", "Error at index " & Index & ": " & Message
    Else
        Err.Raise vbObjectError, "Eval", Message
    End If
End Function
