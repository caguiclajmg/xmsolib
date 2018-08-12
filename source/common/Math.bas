Attribute VB_Name = "common_Math"
Option Explicit

Public Enum xmsoNumericComparisonOperator
    xmsoNumericComparisonOperatorLessThan
    xmsoNumericComparisonOperatorLessThanOrEqual
    xmsoNumericComparisonOperatorEqual
    xmsoNumericComparisonOperatorGreaterThanOrEqual
    xmsoNumericComparisonOperatorGreaterThan
    xmsoNumericComparisonOperatorBetweenExclusive
    xmsoNumericComparisonOperatorBetweenInclusive
End Enum

Public Function Math_EvaluateNumericComparison(ByVal operator As xmsoNumericComparisonOperator, ParamArray parameters() As Variant) As Boolean
    Dim indexStart As Long: indexStart = LBound(parameters)
    
    Select Case operator
        Case xmsoNumericComparisonOperatorLessThan
            Math_EvaluateNumericComparison = parameters(indexStart) < parameters(indexStart + 1)
            
        Case xmsoNumericComparisonOperatorLessThanOrEqual
            Math_EvaluateNumericComparison = parameters(indexStart) <= parameters(indexStart + 1)
            
        Case xmsoNumericComparisonOperatorEqual
            Math_EvaluateNumericComparison = parameters(indexStart) = parameters(indexStart + 1)
        
        Case xmsoNumericComparisonOperatorGreaterThanOrEqual
            Math_EvaluateNumericComparison = parameters(indexStart) >= parameters(indexStart + 1)
        
        Case xmsoNumericComparisonOperatorGreaterThan
            Math_EvaluateNumericComparison = parameters(indexStart) > parameters(indexStart + 1)
        
        Case xmsoNumericComparisonOperatorBetweenExclusive
            Math_EvaluateNumericComparison = (parameters(indexStart) > parameters(indexStart + 1)) And (parameters(indexStart) < parameters(indexStart + 2))
            
        Case xmsoNumericComparisonOperatorBetweenInclusive
            Math_EvaluateNumericComparison = (parameters(indexStart) >= parameters(indexStart + 1)) And (parameters(indexStart) <= parameters(indexStart + 2))
            
        Case Else
            Err.Raise 5, Description:="Unknown comparison operator."
    End Select
End Function

Public Function Math_Fibonacci(ByVal n As Long) As Double
    If n = 0 Then
        Math_Fibonacci = 0#
        Exit Function
    End If
    
    If n = 1 Then
        Math_Fibonacci = 1#
        Exit Function
    End If
    
    Dim previous As Double: previous = 0#
    Dim current As Double: current = 1#
    
    Dim i As Long
    For i = 2 To n
        Dim newCurrent As Double: newCurrent = previous + current
        previous = current
        current = newCurrent
    Next
    
    Math_Fibonacci = current
End Function
