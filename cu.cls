VERSION 1.0 CLASS

123456 

BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function adding(mat1, mat2)

    Dim Matrix1(), Matrix2() As Variant
    Matrix1 = mat1.Value
    Matrix2 = mat2.Value
    
    Dim R1, C1, R2, C2 As Double
    R1 = UBound(Matrix1, 1) - LBound(Matrix1, 1) + 1
    C1 = UBound(Matrix1, 2) - LBound(Matrix1, 2) + 1
    R2 = UBound(Matrix2, 1) - LBound(Matrix2, 1) + 1
    C2 = UBound(Matrix2, 2) - LBound(Matrix2, 2) + 1
    
    If (R1 = R2) And (C1 = C2) Then
    
        Dim i1, j1 As Integer
        ReDim Add(1 To R1, 1 To C1) As Variant
        
        For i1 = 1 To R1
            For j1 = 1 To C1
                Add(i1, j1) = Matrix1(i1, j1) + Matrix2(i1, j1)
            Next j1
        Next i1
        
        adding = Add
    Else
    
        adding = "cannot add"
        
        
    End If
          
End Function

Function MatrixMultiplication(mat1, mat2)

    Dim Matrix1(), Matrix2() As Variant
    Matrix1 = mat1.Value
    Matrix2 = mat2.Value
    
    Dim R1, C1, R2, C2 As Double
    R1 = UBound(Matrix1, 1) - LBound(Matrix1, 1) + 1
    C1 = UBound(Matrix1, 2) - LBound(Matrix1, 2) + 1
    R2 = UBound(Matrix2, 1) - LBound(Matrix2, 1) + 1
    C2 = UBound(Matrix2, 2) - LBound(Matrix2, 2) + 1
    
    If C1 = R2 Then
        
        Dim i1, j1, k1 As Integer
        ReDim Multiply(1 To R1, 1 To C2) As Variant
        
        For i1 = 1 To R1
            For j1 = 1 To C2
                For k1 = 1 To C1
                Multiply(i1, j1) = Multiply(i1, j1) + Matrix1(i1, k1) * Matrix2(k1, j1)
                Next k1
            Next j1
        Next i1
        
        MatrixMultiplication = Multiply
    Else
        MatrixMultiplication = "cannot multiply"
    
    End If

End Function
        
        

