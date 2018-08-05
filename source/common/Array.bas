Attribute VB_Name = "common_Array"
Option Explicit

Public Function Array_Count(ByRef arr As Variant) As Long
    Array_Count = UBound(arr) - LBound(arr) + 1
End Function

Public Function Array_Equals(ByRef arr As Variant, ByRef other As Variant) As Boolean
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
