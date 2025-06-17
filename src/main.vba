
' 度分秒の計算を行う関数（Sok）
' Develeping & Refactoring by Sakamochanq.
' Special thanks Github Copilot.

'---------------------------------------------------------------------------------'

Function sok_def(ParamArray args() As Variant) As String
    Dim totalSeconds As Double
    totalSeconds = 0
    
    Dim i As Long
    For i = LBound(args) To UBound(args)
        Dim s As String
        s = Trim(CStr(args(i)))
        
        If s = "" Then GoTo NextArg
        
        ' 符号を判定
        Dim sign As Double
        sign = 1
        If Left(s, 1) = "-" Then
            sign = -1
            s = Mid(s, 2)
        End If
        
        ' 度分秒を数値に分解
        Dim deg As Double, min As Double, sec As Double
        deg = 0: min = 0: sec = 0
        
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
        
        ' 秒に変換して加算
        totalSeconds = totalSeconds + sign * (deg * 3600 + min * 60 + sec)
NextArg:
    Next i
    
    ' 結果を度分秒に変換
    Dim finalDeg As Long, finalMin As Long, finalSec As Long
    Dim absSec As Long
    Dim neg As Boolean: neg = (totalSeconds < 0)
    
    absSec = Abs(totalSeconds)
    finalDeg = Int(absSec / 3600)
    finalMin = Int((absSec Mod 3600) / 60)
    finalSec = absSec Mod 60
    
    If neg Then finalDeg = -finalDeg
    
    sok_def = finalDeg & "°" & finalMin & "′" & finalSec & "″"
End Function




' sok_sum: 指定範囲内の度分秒文字列を合計し、結果を "度°分′秒″" 形式で返す関数
Public Function sok_sum(rng As Range) As String
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
    sok_sum = totalDegrees & "°" & totalMinutes & "′" & totalSeconds & "″"
End Function

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