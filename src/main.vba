
' 度分秒の計算を行う関数（Sok）
' Develeping & Refactoring by Sakamochanq.
' Special thanks Github Copilot.

'---------------------------------------------------------------------------------'

' ランダムな度分秒を返す関数（デバッグ用）
Function sok_random() As String
    Dim deg As Integer
    Dim min As Integer
    Dim sec As Integer

    Randomize ' 毎回異なる乱数にする

    deg = Int(Rnd() * 360)      ' 0 ～ 359
    min = Int(Rnd() * 60)       ' 0 ～ 59
    sec = Int(Rnd() * 60)       ' 0 ～ 59

    sok_random = deg & "°" & min & "′" & sec & "″"
End Function



' sok_add関数：DMS文字列を加算する
Function sok_add(ParamArray angles() As Variant) As String
    Dim total As Double
    Dim i As Long
    
    total = 0
    For i = LBound(angles) To UBound(angles)
        total = total + DMSStringToDecimal(CStr(angles(i)))
    Next i
    
    sok_add = DecimalToDMSString(total)
End Function

' sok_sub関数：DMS文字列を減算する（左から順に）
Function sok_sub(ParamArray angles() As Variant) As String
    Dim total As Double
    Dim i As Long
    
    If UBound(angles) < 0 Then
        Exit Function
    End If
    
    total = DMSStringToDecimal(CStr(angles(0)))
    
    For i = 1 To UBound(angles)
        total = total - DMSStringToDecimal(CStr(angles(i)))
    Next i
    
    sok_sub = DecimalToDMSString(total)
End Function



' Function sok_def(ParamArray args() As Variant) As String
'     Dim totalSeconds As Double
'     totalSeconds = 0
    
'     Dim i As Long
'     For i = LBound(args) To UBound(args)
'         Dim s As String
'         s = Trim(CStr(args(i)))
        
'         If s = "" Then GoTo NextArg
        
'         ' 符号を判定
'         Dim sign As Double
'         sign = 1
'         If Left(s, 1) = "-" Then
'             sign = -1
'             s = Mid(s, 2)
'         End If
        
'         ' 度分秒を数値に分解
'         Dim deg As Double, min As Double, sec As Double
'         deg = 0: min = 0: sec = 0
        
'         If InStr(s, "°") > 0 Then
'             deg = Val(Left(s, InStr(s, "°") - 1))
'             s = Mid(s, InStr(s, "°") + 1)
'         End If
        
'         If InStr(s, "′") > 0 Then
'             min = Val(Left(s, InStr(s, "′") - 1))
'             s = Mid(s, InStr(s, "′") + 1)
'         End If
        
'         If InStr(s, "″") > 0 Then
'             sec = Val(Left(s, InStr(s, "″") - 1))
'         End If
        
'         ' 秒に変換して加算
'         totalSeconds = totalSeconds + sign * (deg * 3600 + min * 60 + sec)
' NextArg:
'     Next i
    
'     ' 結果を度分秒に変換
'     Dim finalDeg As Long, finalMin As Long, finalSec As Long
'     Dim absSec As Long
'     Dim neg As Boolean: neg = (totalSeconds < 0)
    
'     absSec = Abs(totalSeconds)
'     finalDeg = Int(absSec / 3600)
'     finalMin = Int((absSec Mod 3600) / 60)
'     finalSec = absSec Mod 60
    
'     If neg Then finalDeg = -finalDeg
    
'     sok_def = finalDeg & "°" & finalMin & "′" & finalSec & "″"
' End Function



' sok_sum: 指定範囲内の度分秒文字列を合計し、結果を "度°分′秒″" 形式で返す関数
Public Function sok_sumAll(rng As Range) As String
    Dim cell As Range
    Dim totalDegrees As Long
    Dim totalMinutes As Long
    Dim totalSeconds As Long
    Dim deg As Long, min As Long, sec As Long
    Dim value As String
    
    ' セル範囲を1つずつ処理
    For Each cell In rng
        If Not IsEmpty(cell.value) Then
            ' セル内の文字列を取得
            value = cell.value
            
            ' 記号を数字に分解して取得
            deg = CLng(Split(value, "°")(0))
            min = CLng(Split(Split(value, "°")(1), "′")(0))
            sec = CLng(Split(Split(Split(value, "°")(1), "′")(1), "″")(0))
            
            ' 合計値に加算
            totalDegrees = totalDegrees + deg
            totalMinutes = totalMinutes + min
            totalSeconds = totalSeconds + sec
        End If
    Next cell
    
    ' 秒→分へ繰り上げ
    totalMinutes = totalMinutes + (totalSeconds \ 60)
    totalSeconds = totalSeconds Mod 60

    ' 分→度へ繰り上げ
    totalDegrees = totalDegrees + (totalMinutes \ 60)
    totalMinutes = totalMinutes Mod 60
    
    ' 結果を文字列として返す
    sok_sumAll = totalDegrees & "°" & totalMinutes & "′" & totalSeconds & "″"
End Function



'----------------------------------------------------------------------------------'

' 度分秒文字列を度（Decimal Degree）に変換
Private Function DMSStringToDecimal(s As String) As Double
    Dim deg As Double, min As Double, sec As Double
    s = Trim(s)
    
    ' 負の符号を保持
    Dim sign As Double
    sign = 1
    If Left(s, 1) = "-" Then
        sign = -1
        s = Mid(s, 2)
    End If
    
    ' 度分秒を抽出
    If InStr(s, "°") > 0 Then
        deg = Val(Left(s, InStr(s, "°") - 1))
        s = Mid(s, InStr(s, "°") + 1)
    End If
    If InStr(s, "′") > 0 Then
        min = Val(Left(s, InStr(s, "′") - 1))
        s = Mid(s, InStr(s, "′") + 1)
    End If
    If InStr(s, "″") > 0 Then
        sec = Val(Left(s, InStr(s, "″") - 1))
    End If
    
    DMSStringToDecimal = sign * (deg + min / 60 + sec / 3600)
End Function

' 10進数度を度分秒文字列に変換
Private Function DecimalToDMSString(ByVal decimalDegrees As Double) As String
    Dim deg As Long
    Dim min As Long
    Dim sec As Long
    Dim remainder As Double
    
    If decimalDegrees < 0 Then
        DecimalToDMSString = "-" & DecimalToDMSString(-decimalDegrees)
        Exit Function
    End If
    
    deg = Int(decimalDegrees)
    remainder = (decimalDegrees - deg) * 60
    min = Int(remainder)
    sec = Round((remainder - min) * 60)
    
    ' 秒が60になる場合の繰り上がり処理
    If sec = 60 Then
        sec = 0
        min = min + 1
    End If
    If min = 60 Then
        min = 0
        deg = deg + 1
    End If

    DecimalToDMSString = deg & "°" & min & "′" & sec & "″"
End Function

'----------------------------------------------------------------------------------'



' sin度分秒の計算
Function sok_sin(dmsString As String) As Double
    Dim deg As Double
    deg = DMSStringToDecimal(dmsString)
    sok_sin = Sin(Application.WorksheetFunction.Radians(deg))
End Function

' cos度分秒の計算
Function sok_cos(dmsString As String) As Double
    Dim deg As Double
    deg = DMSStringToDecimal(dmsString)
    sok_cos = Cos(Application.WorksheetFunction.Radians(deg))
End Function



' 方位角から方位を返す関数
Function sok_compass(angle As String) As String
    Dim degrees As Double
    Dim minutes As Double
    Dim seconds As Double
    Dim totalDegrees As Double
    Dim parts() As String
    
    ' 入力文字列を分解して度・分・秒に分ける
    angle = Replace(angle, "°", " ")
    angle = Replace(angle, "′", " ")
    angle = Replace(angle, "″", " ")
    parts = Split(angle)
    
    If UBound(parts) >= 0 Then degrees = CDbl(parts(0))
    If UBound(parts) >= 1 Then minutes = CDbl(parts(1))
    If UBound(parts) >= 2 Then seconds = CDbl(parts(2))
    
    ' 総角度を計算
    totalDegrees = degrees + (minutes / 60) + (seconds / 3600)
    
    ' 方位を判定して返す
    Select Case totalDegrees
        Case 0 To 90
            sok_compass = "NE" ' 北東
        Case 90 To 180
            sok_compass = "SE" ' 南東
        Case 180 To 270
            sok_compass = "SW" ' 南西
        Case 270 To 360
            sok_compass = "NW" ' 北西
        Case Else
            sok_compass = "#BONK!" ' ボンク！
    End Select
End Function



' 方位角を計算する関数
Function sok_azimuth(dmsString As String) As String
    Dim deg As Double
    deg = DMSStringToDecimal(dmsString) ' 文字列を10進数度に変換
    
    Dim result As Double
    
    If deg < 90 Then
        result = deg
    ElseIf deg < 180 Then
        result = 180 - deg
    ElseIf deg < 270 Then
        result = deg - 180
    ElseIf deg < 360 Then
        result = 360 - deg
    Else
        result = deg - 360
    End If
    
    sok_azimuth = DecimalToDMSString(result)
End Function

