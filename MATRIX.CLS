VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matrix"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Public Function Matrix(lb As Variant, Cols As Variant, Rows As Variant, ParamArray dat() As Variant)
On Error GoTo NotAnArray
ReDim c(lb To Rows, lb To Cols) As Variant
k = LBound(dat)
For i = lb To Rows
    For j = lb To Cols
            If k > UBound(dat) Then
            c(i, j) = 0
            Else
            If IsMissing(dat(k)) Then dat(k) = 0
            c(i, j) = dat(k)
            k = k + 1
            End If
      Next j
Next i
Matrix = c
Exit Function
NotAnArray:
Matrix = Null
Exit Function
End Function

Public Function ZeroMatrix(n As Variant, Optional lb As Variant) As Variant
On Error GoTo EHA
If IsMissing(lb) Then lb = 0
ReDim a(lb To n, lb To n)
For i = lb To n
        For j = lb To n
                a(i, j) = 0
       Next j
  Next i
ZeroMatrix = a
Exit Function
EHA:
ZeroMatrix = Null
Exit Function
End Function

Public Function IdMatrix(n As Variant, Optional lb As Variant) As Variant
On Error GoTo EH
If IsMissing(lb) Then lb = 0
ReDim a(lb To n, lb To n)
For i = lb To n
        For j = lb To n
        If (i = j) Then
                a(i, j) = 1
        Else
                a(i, j) = 0
        End If
    Next j
  Next i
IdMatrix = a
Exit Function
EH:
IdMatrix = Null
Exit Function
End Function


Public Function MatrixAdd(a As Variant, b As Variant) As Variant
On Error GoTo notArr
If Not (IsArray(a)) Then GoTo notArr
If Not (IsArray(b)) Then GoTo notArr
lb = LBound(a)
ub = UBound(a)
ReDim c(lb To ub, lb To ub)
For i = lb To ub
    For j = lb To ub
        c(i, j) = a(i, j) + b(i, j)
    Next j
Next i
MatrixAdd = c
Exit Function
notArr:
MatrixAdd = Null
Exit Function
End Function

Public Function MatrixDet(ByVal a As Variant) As Variant
On Error GoTo errH
lb = LBound(a)
ub = UBound(a)
MatrixDet = 1
For i = lb To ub - 1
If a(i, i) = 0 Then
FGOS = 0
    For L = i + 1 To ub
    If a(L, i) <> 0 Then
            For M = lb To ub
            Q = a(L, M)
            a(L, M) = a(i, M)
            a(i, M) = Q
            Next M
            MatrixDet = -MatrixDet
            FGOS = -1
            Exit For   'L'
    End If
    Next L
    If FGOS <> -1 Then MatrixDet = 0
End If
If MatrixDet = 0 Then Exit Function
For j = i + 1 To ub
t = a(i, j) / a(i, i)
For k = lb To ub
a(k, j) = a(k, j) - a(k, i) * t
Next k
Next j
Next i
For i = lb To ub
MatrixDet = MatrixDet * a(i, i)
Next i
Exit Function
errH:
MatrixDet = Null
Exit Function
End Function

Public Function MatrixInv(a As Variant) As Variant
On Error GoTo errHandler
s = MatrixDet(a)
If IsNull(s) Then GoTo errHandler
If s = 0 Then
    MatrixInv = Null
    Exit Function
End If
lb = LBound(a)
ub = UBound(a)
ReDim IA(lb To ub, lb To ub) As Variant
ReDim sm(lb To ub - 1, lb To ub - 1) As Variant

For v = lb To ub
    For w = lb To ub
    le = 0
            For x = lb To ub - 1
                ce = 0
                For y = lb To ub - 1
                    If v = x Then le = 1
                    If y = w Then ce = 1
                    sm(x, y) = a(x + le, y + ce)
                Next y
            Next x
        IA(w, v) = (-1) ^ (w + v) * MatrixDet(sm) / s
    Next w
Next v
MatrixInv = IA()
Exit Function
errHandler:
MatrixInv = Null
Exit Function
 End Function


Public Function MatrixProd(a As Variant, b As Variant) As Variant
On Error GoTo errHa
If Not (IsArray(a)) Then GoTo errHa
If Not (IsArray(b)) Then GoTo errHa
c1 = UBound(a, 2)
r2 = UBound(b, 1)
If c1 <> r2 Then GoTo errHa
r1 = UBound(a, 1)
c2 = UBound(b, 2)
lb1 = LBound(a, 1)
lb2 = LBound(b, 2)
lb = LBound(a, 2)
ReDim c(lb1 To r1, lb2 To c2) As Variant
For i = lb1 To r1
    For j = lb2 To c2
            c(i, j) = 0
            For k = lb To c1   'or    "R2"
            c(i, j) = c(i, j) + a(i, k) * b(k, j)
            Next k
    Next j
Next i
MatrixProd = c
Exit Function
errHa:
MatrixProd = Null
Exit Function
End Function


Public Function MatrixSubtratct(a As Variant, b As Variant) As Variant
On Error GoTo notArray
If Not (IsArray(a)) Then GoTo notArray
If Not (IsArray(b)) Then GoTo notArray
lb = LBound(a)
ub = UBound(a)
ReDim c(lb To ub, lb To ub)
For i = lb To ub
    For j = lb To ub
        c(i, j) = a(i, j) - b(i, j)
    Next j
Next i
MatrixSubtratct = c
Exit Function
notArray:
MatrixSubtratct = Null
Exit Function

End Function



Public Function MatrixMul(ByVal M As Variant, b As Variant) As Variant
On Error GoTo han
If Not (IsArray(M)) Then GoTo han
If IsArray(b) Then GoTo han
For i = LBound(M) To UBound(M)
    For j = LBound(M) To UBound(M)
            M(i, j) = M(i, j) * b
    Next j
Next i
MatrixMul = M
Exit Function
han:
MatrixMul = Null
Exit Function
End Function
