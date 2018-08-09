Attribute VB_Name = "common_Array"
Option Explicit

Public Function Array_Count(ByVal arr As Variant) As Long
    Array_Count = UBound(arr) - LBound(arr) + 1
End Function

Public Function Array_Equals(ByVal arr As Variant, ByVal other As Variant) As Boolean
    If Array_Count(arr) <> Array_Count(other) Then
        Array_Equals = False
        Exit Function
    End If
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> other(i) Then
            Array_Equals = False
            Exit Function
        End If
    Next
    
    Array_Equals = True
End Function

Public Function Array_Contains(ByVal arr As Variant, ByVal match As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = match Then
            Array_Contains = True
            Exit Function
        End If
    Next
    
    Array_Contains = False
End Function
