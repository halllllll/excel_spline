Attribute VB_Name = "func"
Option Explicit

Sub TestSplineInterpolator()
    Dim sp As SplineInterpolator
    Set sp = New SplineInterpolator
    
    ' 例：Sheet1 の A1:A10 に x 値、B1:B10 に y 値があるとする
    Dim xRange As Range, yRange As Range
    Set xRange = ThisWorkbook.Sheets("Sheet1").Range("A12:A21")
    Set yRange = ThisWorkbook.Sheets("Sheet1").Range("B12:B21")
    
    ' 初期化：入力データの検証と係数計算を一度だけ実行
    Call sp.Init(xRange, yRange)
    ' 任意の x 値で補間結果を取得（例：x = 5.5）
    Dim interpVal As Double
    interpVal = sp.Evaluate(5.5)
    
    Debug.Print "x = 5.5 の補間結果: "; interpVal
End Sub

Public Function ExcelInterpolator(xRange As Range, yRange As Range, v As Double) As Double
    Dim sp As SplineInterpolator
    Set sp = New SplineInterpolator

    Call sp.Init(xRange, yRange)
    Dim interpVal As Double
    interpVal = sp.Evaluate(v)
    
    ExcelInterpolator = interpVal
End Function
