Attribute VB_Name = "func"
Option Explicit

Sub TestSplineInterpolator()
    Dim sp As SplineInterpolator
    Set sp = New SplineInterpolator
    
    ' ��FSheet1 �� A1:A10 �� x �l�AB1:B10 �� y �l������Ƃ���
    Dim xRange As Range, yRange As Range
    Set xRange = ThisWorkbook.Sheets("Sheet1").Range("A12:A21")
    Set yRange = ThisWorkbook.Sheets("Sheet1").Range("B12:B21")
    
    ' �������F���̓f�[�^�̌��؂ƌW���v�Z����x�������s
    Call sp.Init(xRange, yRange)
    ' �C�ӂ� x �l�ŕ�Ԍ��ʂ��擾�i��Fx = 5.5�j
    Dim interpVal As Double
    interpVal = sp.Evaluate(5.5)
    
    Debug.Print "x = 5.5 �̕�Ԍ���: "; interpVal
End Sub

Public Function ExcelInterpolator(xRange As Range, yRange As Range, v As Double) As Double
    Dim sp As SplineInterpolator
    Set sp = New SplineInterpolator

    Call sp.Init(xRange, yRange)
    Dim interpVal As Double
    interpVal = sp.Evaluate(v)
    
    ExcelInterpolator = interpVal
End Function
