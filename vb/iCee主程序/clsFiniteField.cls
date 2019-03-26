VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFiniteField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'¶þÎ¬Âë
#Const UseInverseLUT = False

Private Const m_nMaxOrderShift As Long = 14
Private Const m_nMaxOrder As Long = 2 ^ m_nMaxOrderShift

Private m_nOrder As Long
Private m_nPolynomial As Long
Private m_nPrimitiveRoot As Long

Private m_nDiscreteLog(m_nMaxOrder - 1) As Integer
Private m_nDiscreteExp(m_nMaxOrder - 1) As Integer
#If UseInverseLUT Then
Private m_nInverse(m_nMaxOrder - 1) As Integer
#End If

Friend Property Get Order() As Long
Order = m_nOrder
End Property

Friend Property Get MinimalPolynomial() As Long
MinimalPolynomial = m_nPolynomial
End Property

Friend Property Get PrimitiveRoot() As Long
PrimitiveRoot = m_nPrimitiveRoot
End Property

'internal function
Friend Function pFindPrimitiveRoot() As Long
Dim nPrimeFactor(7) As Long
Dim t(m_nMaxOrderShift - 1) As Long
Dim nPrimeFactorCount As Long
Dim nPrimitiveRoot As Long
Dim i As Long, j As Long, K As Long
Dim m As Long
'///first factor order-1 using naive method
i = m_nOrder - 1
If (i And 1&) = 0 Then
 nPrimeFactor(0) = i \ 2&
 nPrimeFactorCount = 1
 Do Until i And 1&
  i = i \ 2&
 Loop
End If
m = Sqr(i) + 2
j = 3
Do Until j > m
 K = i \ j
 If K * j = i Then
  Do
   i = K
   K = i \ j
  Loop While K * j = i
  nPrimeFactor(nPrimeFactorCount) = (m_nOrder - 1) \ j
  nPrimeFactorCount = nPrimeFactorCount + 1
  m = Sqr(i) + 2
 End If
 j = j + 2
Loop
If i > 1 Then
 nPrimeFactor(nPrimeFactorCount) = (m_nOrder - 1) \ i
 nPrimeFactorCount = nPrimeFactorCount + 1
End If
'///try primitive root using naive method
For nPrimitiveRoot = 2 To m_nOrder - 1
 '///build table
 t(0) = nPrimitiveRoot
 j = nPrimitiveRoot
 For i = 1 To m_nMaxOrderShift - 1
  j = Multiply(j, j)
  t(i) = j
 Next i
 '///try it
 For i = 0 To nPrimeFactorCount - 1
  m = nPrimeFactor(i)
  j = 1
  For K = 0 To m_nMaxOrderShift - 1
   If m And 1& Then j = Multiply(j, t(K))
   m = m \ 2&
   If m = 0 Then Exit For
  Next K
  If j = 1 Then Exit For
 Next i
 '///check if we find it
 If i >= nPrimeFactorCount Then
  pFindPrimitiveRoot = nPrimitiveRoot
  Exit Function
 End If
Next nPrimitiveRoot
End Function

'nOrder: should be a prime or power of two, max=16384
'nPolynomial: only used when order is power of two
'nPrimitiveRoot: 0 for automatic find a primitive root, or the input value must be a primitive root
Friend Function Init(ByVal nOrder As Long, Optional ByVal nPolynomial As Long, Optional ByVal nPrimitiveRoot As Long) As Boolean
Dim i As Long, j As Long, K As Long
'///
If nOrder <= 1 Or nOrder > m_nMaxOrder Then Exit Function
If nOrder = 2 Then
 m_nOrder = 2
 m_nPolynomial = 2
 m_nPrimitiveRoot = 1
 m_nDiscreteLog(0) = 0
 m_nDiscreteLog(1) = 0
 m_nDiscreteExp(0) = 1
 m_nDiscreteExp(1) = 1
 #If UseInverseLUT Then
 m_nInverse(0) = 0
 m_nInverse(1) = 1
 #End If
ElseIf (nOrder And (nOrder - 1)) = 0 Then
 '///power of two
 If nPolynomial = 0 Then Exit Function
 m_nOrder = nOrder
 m_nPolynomial = nPolynomial
 If nPrimitiveRoot = 0 Then nPrimitiveRoot = pFindPrimitiveRoot
 m_nPrimitiveRoot = nPrimitiveRoot
 '///calc discrete log LUT
 m_nDiscreteLog(0) = 0
 m_nDiscreteLog(1) = 0
 m_nDiscreteLog(nPrimitiveRoot) = 1
 m_nDiscreteExp(0) = 1
 m_nDiscreteExp(nOrder - 1) = 1
 m_nDiscreteExp(1) = nPrimitiveRoot
 j = nPrimitiveRoot
 For i = 2 To nOrder - 2
  j = Multiply(j, nPrimitiveRoot)
  m_nDiscreteLog(j) = i
  m_nDiscreteExp(i) = j
 Next i
 '///calc inverse LUT
 #If UseInverseLUT Then
 m_nInverse(0) = 0
 m_nInverse(1) = 1
 For i = 1 To nOrder - 2
  m_nInverse(m_nDiscreteExp(i)) = m_nDiscreteExp(nOrder - 1 - i)
 Next i
 #End If
Else
 '///assume it's a prime
 m_nOrder = nOrder
 m_nPolynomial = nOrder
 If nPrimitiveRoot = 0 Then nPrimitiveRoot = pFindPrimitiveRoot
 m_nPrimitiveRoot = nPrimitiveRoot
 '///calc discrete log LUT
 m_nDiscreteLog(0) = 0
 m_nDiscreteLog(1) = 0
 m_nDiscreteLog(nPrimitiveRoot) = 1
 m_nDiscreteExp(0) = 1
 m_nDiscreteExp(nOrder - 1) = 1
 m_nDiscreteExp(1) = nPrimitiveRoot
 j = nPrimitiveRoot
 For i = 2 To nOrder - 2
  j = (j * nPrimitiveRoot) Mod nOrder
  m_nDiscreteLog(j) = i
  m_nDiscreteExp(i) = j
 Next i
 '///calc inverse LUT
 #If UseInverseLUT Then
 m_nInverse(0) = 0
 m_nInverse(1) = 1
 For i = 1 To nOrder - 2
  m_nInverse(m_nDiscreteExp(i)) = m_nDiscreteExp(nOrder - 1 - i)
 Next i
 #End If
End If
'///over
Init = True
End Function

'input should be >=0, <order
Friend Function Add(ByVal n1 As Long, ByVal n2 As Long) As Long
If (m_nOrder And (m_nOrder - 1)) = 0 Then
 Add = n1 Xor n2
Else
 n1 = n1 + n2
 If n1 >= m_nOrder Then n1 = n1 - m_nOrder
 Add = n1
End If
End Function

'input should be >=0, <order
Friend Function Negative(ByVal n1 As Long) As Long
If (m_nOrder And (m_nOrder - 1)) = 0 Or n1 = 0 Then
 Negative = n1
Else
 Negative = m_nOrder - n1
End If
End Function

'input should be >=0, <order
Friend Function Subtract(ByVal n1 As Long, ByVal n2 As Long) As Long
If (m_nOrder And (m_nOrder - 1)) = 0 Then
 Subtract = n1 Xor n2
Else
 n1 = n1 - n2
 If n1 < 0 Then n1 = n1 + m_nOrder
 Subtract = n1
End If
End Function

'input should be >=0, <order
Friend Function MultiplyLUT(ByVal n1 As Long, ByVal n2 As Long) As Long
If n1 = 0 Or n2 = 0 Then Exit Function
n1 = m_nDiscreteLog(n1) + m_nDiscreteLog(n2)
If n1 >= m_nOrder - 1 Then n1 = n1 - m_nOrder + 1
MultiplyLUT = m_nDiscreteExp(n1)
End Function

'input should be >=0, <order
Friend Function Multiply(ByVal n1 As Long, ByVal n2 As Long) As Long
Dim i As Long, j As Long
'///
If n1 = 0 Or n2 = 0 Then Exit Function
If m_nOrder = 2 Then
 Multiply = 1
ElseIf (m_nOrder And (m_nOrder - 1)) = 0 Then
 i = 1
 Do
  If n2 And i Then
   n2 = n2 Xor i
   j = j Xor n1
  End If
  If n2 = 0 Then Exit Do
  i = i + i
  n1 = n1 + n1
  If n1 And m_nOrder Then n1 = (n1 Xor m_nPolynomial) And (m_nOrder - 1)
 Loop
 Multiply = j
Else
 Multiply = (n1 * n2) Mod m_nOrder
End If
End Function

'input should be >=0, <order
Friend Function InverseLUT(ByVal n As Long) As Long
If n = 0 Then Exit Function
#If UseInverseLUT Then
InverseLUT = m_nInverse(n)
#Else
InverseLUT = m_nDiscreteExp(m_nOrder - 1 - m_nDiscreteLog(n))
#End If
End Function

'input should be >=0, <order
Friend Function Inverse(ByVal b As Long) As Long
Dim a As Long, Q As Long, t As Long
Dim Y As Long, y1 As Long
'///
If b = 0 Then Exit Function
If m_nOrder = 2 Then
 Inverse = 1
ElseIf (m_nOrder And (m_nOrder - 1)) = 0 Then
Else
 '///extended GCD
 a = m_nOrder
 'x := 0    lastx := 1 'does not need
 'y := 1    lasty := 0
 Y = 1
 'while b ¡Ù 0
 Do While b <> 0
  'quotient := a div b
  Q = a \ b
  '(a, b) := (b, a mod b)
  t = a - Q * b
  a = b
  b = t
  '(x, lastx) := (lastx - quotient*x, x) 'does not need
  '(y, lasty) := (lasty - quotient*y, y)
  t = y1 - Q * Y
  y1 = Y
  Y = t
 Loop
 'return (lastx, lasty)
 y1 = y1 Mod m_nOrder
 If y1 < 0 Then y1 = y1 + m_nOrder
 Inverse = y1
End If
End Function

'input should be >=0, <order
Friend Function DivideLUT(ByVal n1 As Long, ByVal n2 As Long) As Long
If n1 = 0 Or n2 = 0 Then Exit Function
n1 = m_nDiscreteLog(n1) - m_nDiscreteLog(n2)
If n1 < 0 Then n1 = n1 + m_nOrder - 1
DivideLUT = m_nDiscreteExp(n1)
End Function

'input should be >=0, <order
Friend Function Divide(ByVal n1 As Long, ByVal n2 As Long) As Long
Dim a As Long, Q As Long, t As Long
Dim Y As Long, y1 As Long
'///
If n1 = 0 Or n2 = 0 Then Exit Function
If m_nOrder = 2 Then
 Divide = 1
ElseIf (m_nOrder And (m_nOrder - 1)) = 0 Then
 'TODO:
 Debug.Assert False
Else
 '///extended GCD
 a = m_nOrder
 'x := 0    lastx := 1 'does not need
 'y := 1    lasty := 0
 Y = 1
 'while b ¡Ù 0
 Do While n2 <> 0
  'quotient := a div b
  Q = a \ n2
  '(a, b) := (b, a mod b)
  t = a - Q * n2
  a = n2
  n2 = t
  '(x, lastx) := (lastx - quotient*x, x) 'does not need
  '(y, lasty) := (lasty - quotient*y, y)
  t = y1 - Q * Y
  y1 = Y
  Y = t
 Loop
 'return (lastx, lasty)
 y1 = (n1 * y1) Mod m_nOrder
 If y1 < 0 Then y1 = y1 + m_nOrder
 Divide = y1
End If
End Function

'input should be >=0, <order
Friend Function DiscreteLog(ByVal n As Long) As Long
DiscreteLog = m_nDiscreteLog(n)
End Function

Friend Function DiscreteExp(ByVal n As Long) As Long
n = n Mod (m_nOrder - 1)
If n < 0 Then n = n + m_nOrder - 1
DiscreteExp = m_nDiscreteExp(n)
End Function


