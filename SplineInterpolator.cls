VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SplineInterpolator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 内部データ
Private m_x() As Double     ' x 値の配列
Private m_y() As Double     ' y 値の配列
Private m_h() As Double     ' 隣接点間の区間幅
Private m_b() As Double     ' 補間係数 b
Private m_c() As Double     ' 補間係数 c
Private m_d() As Double     ' 補間係数 d
Private m_n As Long         ' データ点数 - 1（セグメント数）

' 初期化メソッド：xRange, yRange のセル範囲からデータを取り込み、係数計算を実行
Public Sub Init(xRange As Range, yRange As Range)
    Dim i As Long
    ' --- 入力チェック ---
    If xRange.Count <> yRange.Count Then
        Err.Raise vbObjectError + 1000, "SplineInterpolator", "xRange と yRange の要素数は一致していなければなりません。"
    End If
    
    m_n = xRange.Count - 1
    If m_n < 1 Then
        Err.Raise vbObjectError + 1001, "SplineInterpolator", "データ点は2点以上必要です。"
    End If
    
    ' --- データ取り込み ---
    ReDim m_x(0 To m_n)
    ReDim m_y(0 To m_n)
    For i = 0 To m_n
        m_x(i) = xRange.Cells(i + 1, 1).Value
        m_y(i) = yRange.Cells(i + 1, 1).Value
    Next i
    
    ' --- x 値が昇順かチェック ---
    For i = 0 To m_n - 1
        If m_x(i) >= m_x(i + 1) Then
            Err.Raise vbObjectError + 1002, "SplineInterpolator", "xRange は昇順でなければなりません。"
        End If
    Next i
    
    ' --- 区間幅 h の計算 ---
    ReDim m_h(0 To m_n - 1)
    For i = 0 To m_n - 1
        m_h(i) = m_x(i + 1) - m_x(i)
    Next i
    
    ' --- α の計算 (端点は 0) ---
    Dim alpha() As Double
    ReDim alpha(0 To m_n)
    alpha(0) = 0
    alpha(m_n) = 0
    For i = 1 To m_n - 1
        alpha(i) = (3# / m_h(i)) * (m_y(i + 1) - m_y(i)) - (3# / m_h(i - 1)) * (m_y(i) - m_y(i - 1))
    Next i
    
    ' --- 三重対角行列の解法（Thomas アルゴリズム） ---
    Dim l() As Double, mu() As Double, z() As Double
    ReDim l(0 To m_n)
    ReDim mu(0 To m_n)
    ReDim z(0 To m_n)
    ReDim m_c(0 To m_n)
    
    l(0) = 1#
    mu(0) = 0
    z(0) = 0
    
    For i = 1 To m_n - 1
        l(i) = 2# * (m_x(i + 1) - m_x(i - 1)) - m_h(i - 1) * mu(i - 1)
        mu(i) = m_h(i) / l(i)
        z(i) = (alpha(i) - m_h(i - 1) * z(i - 1)) / l(i)
    Next i
    
    l(m_n) = 1#
    z(m_n) = 0
    m_c(m_n) = 0
    
    ' --- 補間係数 b, d の計算 ---
    Dim iIndex As Long
    ReDim m_b(0 To m_n - 1)
    ReDim m_d(0 To m_n - 1)
    For iIndex = m_n - 1 To 0 Step -1
        m_c(iIndex) = z(iIndex) - mu(iIndex) * m_c(iIndex + 1)
        m_b(iIndex) = (m_y(iIndex + 1) - m_y(iIndex)) / m_h(iIndex) - m_h(iIndex) * (m_c(iIndex + 1) + 2# * m_c(iIndex)) / 3#
        m_d(iIndex) = (m_c(iIndex + 1) - m_c(iIndex)) / (3# * m_h(iIndex))
    Next iIndex
End Sub

' 任意の xValue に対して補間値を返すメソッド
Public Function Evaluate(xValue As Double) As Double
    Dim i As Long, idx As Long
    ' --- 範囲外の処理 ---
    If xValue <= m_x(0) Then
        idx = 0
    ElseIf xValue >= m_x(m_n) Then
        idx = m_n - 1
    Else
        For i = 0 To m_n - 1
            If xValue >= m_x(i) And xValue <= m_x(i + 1) Then
                idx = i
                Exit For
            End If
        Next i
    End If
    
    Dim dx As Double
    dx = xValue - m_x(idx)
    Evaluate = m_y(idx) + m_b(idx) * dx + m_c(idx) * dx ^ 2 + m_d(idx) * dx ^ 3
End Function

