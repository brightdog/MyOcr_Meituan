Attribute VB_Name = "modLD"
Option Explicit

Public Function LD(s, t) As Long
   
    Dim n, m As Long
    n = Len(s)
    m = Len(t)
 
    If n = 0 Or m = 0 Then
        Exit Function
    End If
 
    Dim d() As Long
    ReDim d(n, m) As Long
       
    Dim i, j
       
    For i = 0 To n
        d(i, 0) = i
    Next

    For j = 0 To m
        d(0, j) = j
    Next
 
    Dim s_i, t_j
    Dim Cost
   
    For i = 1 To n
        s_i = Mid(s, i, 1)

        For j = 1 To m
            t_j = Mid(t, j, 1)

            If s_i = t_j Then
                Cost = 0
            Else
                Cost = 1
            End If

            d(i, j) = Minimum(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + Cost)
        Next
    Next
 
    LD = d(n, m)
End Function
 
Private Function Minimum(a, b, c) As Long
    Dim mi As Long
    mi = a

    If b < mi Then
        mi = b
    End If

    If c < mi Then
        mi = c
    End If
 
    Minimum = mi
End Function
