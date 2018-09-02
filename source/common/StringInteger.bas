Attribute VB_Name = "common_StringInteger"
Option Explicit

Public Type StringInteger
    Value As String
    Negative As Boolean
End Type

Public Function StringInteger_ToString(ByRef lhs As StringInteger) As String
    StringInteger_ToString = IIf(lhs.Negative, "-", vbNullString) & lhs.Value
End Function

Public Function StringInteger(ByVal Value As String, ByVal Negative As Boolean) As StringInteger
    Dim result As StringInteger
    result.Value = Value
    result.Negative = Negative
    
    StringInteger = result
End Function

Public Function StringInteger_IsZero(ByRef lhs As StringInteger) As Boolean
    StringInteger_IsZero = lhs.Value = "0"
End Function

Public Function StringInteger_Compare(ByRef lhs As StringInteger, ByRef rhs As StringInteger) As Integer
    If lhs.Negative <> rhs.Negative Then
        StringInteger_Compare = IIf(lhs.Negative, -1, 1)
        Exit Function
    End If
    
    If lhs.Value = rhs.Value Then
        StringInteger_Compare = 0
        Exit Function
    End If
    
    Dim lhsLength As Long: lhsLength = Len(lhs.Value)
    Dim rhsLength As Long: rhsLength = Len(rhs.Value)
    
    If lhsLength = rhsLength Then
        Dim place As Long
        For place = 1 To lhsLength
            Dim lhsDigit As Long: lhsDigit = AscW(Mid$(lhs.Value, place, 1)) - xmsoCharacterConstant0
            Dim rhsDigit As Long: rhsDigit = AscW(Mid$(rhs.Value, place, 1)) - xmsoCharacterConstant0
            
            If lhsDigit <> rhsDigit Then
                StringInteger_Compare = IIf(lhs.Negative, IIf(lhsDigit > rhsDigit, -1, 1), IIf(lhsDigit > rhsDigit, 1, -1))
                Exit Function
            End If
        Next
    Else
        StringInteger_Compare = IIf(lhs.Negative, IIf(lhsLength > rhsLength, -1, 1), IIf(lhsLength < rhsLength, -1, 1))
        Exit Function
    End If
End Function

Public Function StringInteger_IsLessThan(ByRef lhs As StringInteger, ByRef rhs As StringInteger) As Boolean
    StringInteger_IsLessThan = (StringInteger_Compare(lhs, rhs) = -1)
End Function

Public Function StringInteger_IsGreaterThan(ByRef lhs As StringInteger, ByRef rhs As StringInteger) As Boolean
    StringInteger_IsGreaterThan = (StringInteger_Compare(lhs, rhs) = 1)
End Function

Public Function StringInteger_IsEqualTo(ByRef lhs As StringInteger, ByRef rhs As StringInteger) As Boolean
    StringInteger_IsEqualTo = (StringInteger_Compare(lhs, rhs) = 0)
End Function

Public Function StringInteger_IsLessThanOrEqualTo(ByRef lhs As StringInteger, ByRef rhs As StringInteger) As Boolean
    Dim comparison As Integer: comparison = StringInteger_Compare(lhs, rhs)
    StringInteger_IsLessThanOrEqualTo = (comparison = -1) Or (comparison = 0)
End Function

Public Function StringInteger_IsGreaterThanOrEqualTo(ByRef lhs As StringInteger, ByRef rhs As StringInteger) As Boolean
    Dim comparison As Integer: comparison = StringInteger_Compare(lhs, rhs)
    StringInteger_IsGreaterThanOrEqualTo = (comparison = 1) Or (comparison = 0)
End Function

Public Function StringInteger_AbsoluteValue(ByRef lhs As StringInteger) As StringInteger
    Dim result As StringInteger
    result.Value = lhs.Value
    result.Negative = False
    
    StringInteger_AbsoluteValue = result
End Function

Public Function StringInteger_Negate(ByRef lhs As StringInteger) As StringInteger
    StringInteger_Negate = StringInteger(lhs.Value, Not lhs.Negative)
End Function

Public Function StringInteger_Add(ByRef lhs As StringInteger, ByRef rhs As StringInteger) As StringInteger
    If lhs.Value = "0" Then
        StringInteger_Add = rhs
        Exit Function
    End If
    
    If rhs.Value = "0" Then
        StringInteger_Add = lhs
        Exit Function
    End If
    
    If lhs.Negative <> rhs.Negative Then
        StringInteger_Add = StringInteger_Subtract(lhs, StringInteger(rhs.Value, Not rhs.Negative))
        Exit Function
    End If
    
    Dim lhsNumber As String: lhsNumber = lhs.Value
    Dim rhsNumber As String: rhsNumber = rhs.Value
    Dim lhsLength As Long: lhsLength = Len(lhsNumber)
    Dim rhsLength As Long: rhsLength = Len(rhsNumber)
    
    Dim padLength As Long: padLength = Math_Max(lhsLength, rhsLength)
    
    lhsNumber = String(padLength - lhsLength, "0") & lhsNumber
    rhsNumber = String(padLength - rhsLength, "0") & rhsNumber
    
    Dim resultNumber As String
    Dim carry As Long: carry = 0
    Dim place As Long
    For place = padLength To 1 Step -1
        Dim lhsDigit As Long: lhsDigit = AscW(Mid$(lhsNumber, place, 1)) - xmsoCharacterConstant0
        Dim rhsDigit As Long: rhsDigit = AscW(Mid$(rhsNumber, place, 1)) - xmsoCharacterConstant0
        Dim digitSum As Long: digitSum = lhsDigit + rhsDigit + carry
        If digitSum > 9 Then
            carry = digitSum \ 10
            digitSum = digitSum - 10
        Else
            carry = 0
        End If
        
        resultNumber = ChrW$(xmsoCharacterConstant0 + digitSum) & resultNumber
    Next
    
    If carry <> 0 Then resultNumber = ChrW$(xmsoCharacterConstant0 + carry) & resultNumber
    
    Dim result As StringInteger
    result.Value = resultNumber
    result.Negative = IIf(result.Value = "0", False, lhs.Negative)
    
    StringInteger_Add = result
End Function

Public Function StringInteger_Subtract(ByRef lhs As StringInteger, ByRef rhs As StringInteger) As StringInteger
    If lhs.Value = "0" Then
        StringInteger_Subtract = StringInteger_Negate(rhs)
        Exit Function
    End If
    
    If rhs.Value = "0" Then
        StringInteger_Subtract = StringInteger_Negate(lhs)
        Exit Function
    End If
    
    If lhs.Negative <> rhs.Negative Then
        StringInteger_Subtract = StringInteger_Add(lhs, StringInteger(rhs.Value, Not rhs.Negative))
        Exit Function
    End If
    
    Dim lhsNumber As String: lhsNumber = lhs.Value
    Dim rhsNumber As String: rhsNumber = rhs.Value
    
    Dim comparison As Integer: comparison = StringInteger_Compare(StringInteger(lhs.Value, False), StringInteger(rhs.Value, False))
    If comparison = -1 Then Utility_Swap lhsNumber, rhsNumber
    Dim resultNegative As Boolean: resultNegative = IIf(comparison = -1, rhs.Negative, lhs.Negative)
    
    Dim lhsLength As Long: lhsLength = Len(lhsNumber)
    Dim rhsLength As Long: rhsLength = Len(rhsNumber)
    
    Dim padLength As Long: padLength = Math_Max(lhsLength, rhsLength)
    
    lhsNumber = String(padLength - lhsLength, "0") & lhsNumber
    rhsNumber = String(padLength - rhsLength, "0") & rhsNumber
    
    Dim resultNumber As String, borrow As Long
    Dim place As Long
    For place = padLength To 1 Step -1
        Dim lhsDigit As Long: lhsDigit = AscW(Mid$(lhsNumber, place, 1)) - xmsoCharacterConstant0
        Dim rhsDigit As Long: rhsDigit = AscW(Mid$(rhsNumber, place, 1)) - xmsoCharacterConstant0
        Dim digitDifference As Long: digitDifference = lhsDigit - rhsDigit
        If digitDifference < 0 Then
            Dim borrowPlace As Long: borrowPlace = place - 1
            Do
                If AscW(Mid$(lhsNumber, borrowPlace, 1)) - xmsoCharacterConstant0 <> 0 Then Exit Do
                borrowPlace = borrowPlace - 1
            Loop
            
            Mid$(lhsNumber, borrowPlace, 1) = CStr((AscW(Mid$(lhsNumber, borrowPlace, 1)) - xmsoCharacterConstant0) - 1)
            
            Dim borrowIndex As Long
            For borrowIndex = borrowPlace + 1 To place - 1
                Mid$(lhsNumber, borrowIndex, 1) = CStr(9)
            Next
            
            digitDifference = digitDifference + 10
        End If
        
        resultNumber = ChrW$(xmsoCharacterConstant0 + Abs(digitDifference)) & resultNumber
    Next
    
    Dim result As StringInteger
    result.Value = String_TrimStart(resultNumber, "0")
    If result.Value = vbNullString Then result.Value = "0"
    result.Negative = IIf(result.Value = "0", False, resultNegative)
    
    StringInteger_Subtract = result
End Function

Public Function StringInteger_Multiply(ByRef lhs As StringInteger, ByRef rhs As StringInteger) As StringInteger
    If (lhs.Value = "0") Or (rhs.Value = "0") Then
        StringInteger_Multiply = StringInteger("0", False)
        Exit Function
    End If
    
    If lhs.Value = "1" Then
        StringInteger_Multiply = StringInteger(rhs.Value, IIf(lhs.Negative, Not rhs.Negative, rhs.Negative))
        Exit Function
    End If
    
    If rhs.Value = "1" Then
        StringInteger_Multiply = StringInteger(lhs.Value, IIf(rhs.Negative, Not lhs.Negative, lhs.Negative))
        Exit Function
    End If
    
    Dim result As StringInteger: result = StringInteger("0", False)
    
    Dim lhsNumber As String: lhsNumber = lhs.Value
    Dim rhsNumber As String: rhsNumber = rhs.Value
    Dim lhsLength As Long: lhsLength = Len(lhsNumber)
    Dim rhsLength As Long: rhsLength = Len(rhsNumber)
    
    If rhsLength > lhsLength Then
        StringInteger_Multiply = StringInteger_Multiply(rhs, lhs)
        Exit Function
    End If
    
    Dim rhsPlace As Long
    For rhsPlace = rhsLength To 1 Step -1
        Dim rhsDigit As Long: rhsDigit = AscW(Mid$(rhsNumber, rhsPlace, 1)) - xmsoCharacterConstant0
        Dim productString As String: productString = vbNullString
        
        Dim carry As Long: carry = 0
        Dim digitProduct As Long: digitProduct = 0
        Dim lhsPlace As Long
        For lhsPlace = lhsLength To 1 Step -1
            Dim lhsDigit As Long: lhsDigit = AscW(Mid$(lhsNumber, lhsPlace, 1)) - xmsoCharacterConstant0
            
            digitProduct = (rhsDigit * lhsDigit) + carry
            productString = (digitProduct Mod 10) & productString
            carry = digitProduct \ 10
        Next
        If carry > 0 Then productString = carry & productString
        productString = productString & String(rhsLength - rhsPlace, "0")
        
        result = StringInteger_Add(result, StringInteger(productString, False))
    Next
    
    result.Value = String_TrimStart(result.Value, "0")
    If result.Value = vbNullString Then result.Value = "0"
    result.Negative = lhs.Negative Xor rhs.Negative
    
    StringInteger_Multiply = result
End Function

Public Function StringInteger_Modulo(ByRef lhs As StringInteger, ByRef rhs As StringInteger) As StringInteger
    Dim result As StringInteger
    
    If StringInteger_IsLessThan(StringInteger(lhs.Value, False), StringInteger(rhs.Value, False)) Then
        result.Value = lhs.Value
        result.Negative = lhs.Negative
    ElseIf StringInteger_IsEqualTo(lhs, rhs) Then
        result.Value = "0"
        result.Negative = False
    Else
        Dim counter As StringInteger: counter = StringInteger("1", False)
        
        Dim last As StringInteger: last = rhs
        While StringInteger_IsLessThan(last, StringInteger_AbsoluteValue(lhs))
            counter = StringInteger_Increment(counter)
            last = StringInteger_Multiply(rhs, counter)
        Wend
        
        result.Value = StringInteger_Subtract(StringInteger_AbsoluteValue(lhs), StringInteger_Multiply(rhs, StringInteger_Decrement(counter))).Value
        result.Negative = lhs.Negative
    End If
    
    StringInteger_Modulo = result
End Function

Public Function StringInteger_Increment(ByRef lhs As StringInteger) As StringInteger
    StringInteger_Increment = StringInteger_Add(lhs, StringInteger("1", False))
End Function

Public Function StringInteger_Decrement(ByRef lhs As StringInteger) As StringInteger
    StringInteger_Decrement = StringInteger_Add(lhs, StringInteger("1", True))
End Function
