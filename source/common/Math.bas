Attribute VB_Name = "common_Math"
Option Explicit

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
