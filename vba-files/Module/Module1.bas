Attribute VB_Name = "Module1"
Option Explicit

' 定数定義

' 入出金情報関連定数
Private Const 入出金開始行 As Long = 46 ' 入出金データ開始行
Private Const 入出金日列 As Long = 2     ' B列：入出金日
Private Const 摘要列 As Long = 3         ' C列：摘要
Private Const 入出金金額列 As Long = 4   ' D列：入出金金額
Private Const 残高列 As Long = 5         ' E列：残高
Private Const 約定返済元金列 As Long = 6 ' F列：延滞中の約定返済元金合計

' 期失日関連定数
Private Const 期失日列 As Long = 3  ' C列
Private Const 期失日行 As Long = 25  ' 25行目

' 借入利率関連定数
Private Const 借入利率列 As Long = 2  ' B列：借入利率
Private Const 借入利率開始日列 As Long = 3  ' C列：開始日
Private Const 借入利率開始行 As Long = 29  ' 29行目

' 出力項目列定数（A列～M列）
Private Const 出力_通番列 As Long = 1          ' A列：通番
Private Const 出力_ステータス列 As Long = 2      ' B列：ステータス
Private Const 出力_イベント列 As Long = 3        ' C列：イベント
Private Const 出力_約定返済月列 As Long = 4      ' D列：約定返済月
Private Const 出力_対象元金列 As Long = 5        ' E列：対象元金
Private Const 出力_計算期間開始日列 As Long = 6  ' F列：計算期間開始日
Private Const 出力_区切り列 As Long = 7          ' G列："～"
Private Const 出力_計算期間終了日列 As Long = 8  ' H列：計算期間終了日
Private Const 出力_計算日数列 As Long = 9        ' I列：計算日数
Private Const 出力_利率列 As Long = 10           ' J列：利率
Private Const 出力_積数列 As Long = 11           ' K列：積数
Private Const 出力_利息金額列 As Long = 12       ' L列：利息金額
Private Const 出力_遅延損害金列 As Long = 13     ' M列：遅延損害金

' 返済予定情報の定数
Const 返済予定開始行 As Long = 35
Const 返済予定日列 As Long = 3  ' C列
Const 返済元金列 As Long = 4    ' D列

' 返済予定情報取得関数
' 35行目から開始し、C列が空白になるまで返済予定日、返済元金、返済元金累計を取得
Public Function 返済予定情報取得(targetSheet As Worksheet) As Variant
    Dim currentRow As Long
    Dim dataArray() As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim 返済元金累計 As Double
    
    currentRow = 返済予定開始行
    rowCount = 0
    
    ' データ行数をカウント（C列が空白になるまで）
    Do While targetSheet.Cells(currentRow, 返済予定日列).Value <> ""
        rowCount = rowCount + 1
        currentRow = currentRow + 1
    Loop
    
    ' 35行目のC列が空白の場合も1行として処理
    If rowCount = 0 Then
        rowCount = 1
    End If
    
    ' 配列を初期化（行数 x 3列：返済予定日、返済元金、返済元金累計）
    ReDim dataArray(1 To rowCount, 1 To 3)
    
    ' データを取得してバリデーション
    返済元金累計 = 0
    currentRow = 返済予定開始行
    
    For i = 1 To rowCount
        ' C列：返済予定日
        Dim dateValue As Variant
        dateValue = targetSheet.Cells(currentRow, 返済予定日列).Value
        
        If dateValue = "" Or IsEmpty(dateValue) Then
            ' 空白の場合は日付の初期値を設定
            dataArray(i, 1) = DateSerial(1900, 1, 1)
        Else
            ' 日付型チェック
            If Not IsDate(dateValue) Then
                Err.Raise 13, "返済予定情報取得", currentRow & "行目のC列（返済予定日）が日付ではありません。"
            End If
            dataArray(i, 1) = CDate(dateValue)
        End If
        
        ' D列：返済元金
        Dim principalValue As Variant
        principalValue = targetSheet.Cells(currentRow, 返済元金列).Value
        
        If principalValue = "" Or IsEmpty(principalValue) Then
            ' 空白の場合は0を設定
            dataArray(i, 2) = 0
        Else
            ' 数値型チェック
            If Not IsNumeric(principalValue) Then
                Err.Raise 13, "返済予定情報取得", currentRow & "行目のD列（返済元金）が数値ではありません。"
            End If
            dataArray(i, 2) = CDbl(principalValue)
        End If
        
        ' 返済元金累計を計算
        返済元金累計 = 返済元金累計 + dataArray(i, 2)
        dataArray(i, 3) = 返済元金累計
        
        currentRow = currentRow + 1
        
        ' C列が空白になったら終了（35行目以外）
        If i > 1 And (targetSheet.Cells(currentRow, 返済予定日列).Value = "" Or IsEmpty(targetSheet.Cells(currentRow, 返済予定日列).Value)) Then
            Exit For
        End If
    Next i
    
    返済予定情報取得 = dataArray
End Function

' 入出金情報取得関数
Public Function 入出金情報取得(targetSheet As Worksheet) As Variant
    Dim startRow As Long
    Dim currentRow As Long
    Dim dataArray() As Variant
    Dim rowCount As Long
    Dim i As Long
    
    startRow = 入出金開始行 ' 開始行
    currentRow = startRow
    rowCount = 0
    
    ' データ行数をカウント（B列が空白になるまで、摘要が「返済分」で終わる行は除外）
    Do While targetSheet.Cells(currentRow, 入出金日列).Value <> ""
        Dim 摘要値 As String
        摘要値 = CStr(targetSheet.Cells(currentRow, 摘要列).Value)
        ' 摘要が「返済分」で終わらない場合のみカウント
        If Not (Len(摘要値) >= 3 And Right(摘要値, 3) = "返済分") Then
            rowCount = rowCount + 1
        End If
        currentRow = currentRow + 1
    Loop
    
    ' データが存在しない場合は空の配列を返す
    If rowCount = 0 Then
        入出金情報取得 = Array()
        Exit Function
    End If
    
    ' 配列を初期化（行数 x 5列）
    ReDim dataArray(1 To rowCount, 1 To 5)
    
    ' データを取得してバリデーション
    Dim arrayIndex As Long
    arrayIndex = 1
    currentRow = startRow
    
    Do While targetSheet.Cells(currentRow, 入出金日列).Value <> ""
        Dim 摘要値 As String
        摘要値 = CStr(targetSheet.Cells(currentRow, 摘要列).Value)
        
        ' 摘要が「返済分」で終わらない場合のみ処理
        If Not (Len(摘要値) >= 3 And Right(摘要値, 3) = "返済分") Then
            ' B列：入出金日（日付チェック）
            Dim dateValue As Variant
            dateValue = targetSheet.Cells(currentRow, 入出金日列).Value
            If Not IsDate(dateValue) Then
                Err.Raise 13, "入出金情報取得", currentRow & "行目のB列（入出金日）が日付ではありません。"
            End If
            dataArray(arrayIndex, 1) = CDate(dateValue)
            
            ' C列：摘要（文字列、チェック不要）
            dataArray(arrayIndex, 2) = 摘要値
            
            ' D列：入出金金額（数値チェック）
            Dim amountValue As Variant
            amountValue = targetSheet.Cells(currentRow, 入出金金額列).Value
            If Not IsNumeric(amountValue) Then
                Err.Raise 13, "入出金情報取得", currentRow & "行目のD列（入出金金額）が数値ではありません。"
            End If
            dataArray(arrayIndex, 3) = CDbl(amountValue)
            
            ' E列：残高（数値チェック）
            Dim balanceValue As Variant
            balanceValue = targetSheet.Cells(currentRow, 残高列).Value
            If Not IsNumeric(balanceValue) Then
                Err.Raise 13, "入出金情報取得", currentRow & "行目のE列（残高）が数値ではありません。"
            End If
            dataArray(arrayIndex, 4) = CDbl(balanceValue)
            
            ' F列：延滞中の約定返済元金合計（入力があれば数値チェック）
            Dim principalValue As Variant
            principalValue = targetSheet.Cells(currentRow, 約定返済元金列).Value
            If principalValue <> "" Then
                If Not IsNumeric(principalValue) Then
                    Err.Raise 13, "入出金情報取得", currentRow & "行目のF列（延滞中の約定返済元金合計）が数値ではありません。"
                End If
                dataArray(arrayIndex, 5) = CDbl(principalValue)
            Else
                ' 空白の場合は前の行の値を使用、ただしarrayIndex=1の場合は0を設定
                If arrayIndex = 1 Then
                    dataArray(arrayIndex, 5) = 0
                Else
                    dataArray(arrayIndex, 5) = dataArray(arrayIndex - 1, 5)
                End If
            End If
            
            arrayIndex = arrayIndex + 1
        End If
        
        currentRow = currentRow + 1
    Loop
    
    入出金情報取得 = dataArray
End Function

' 期失日を取得する関数
' 指定されたシートのC列25行目のセル値を返す
Public Function 期失日(targetSheet As Worksheet) As Date
    Dim cellValue As Variant
    
    ' セル値を取得
    cellValue = targetSheet.Cells(期失日行, 期失日列).Value
    
    ' 日付型かチェック
    If Not IsDate(cellValue) Then
        Err.Raise 13, "期失日", "セル値が日付型ではありません。"
    End If
    
    ' 日付型に変換して返す
    期失日 = CDate(cellValue)
End Function

' 借入利率取得関数
' 指定されたシートの29行目からB列が空白になるまで、B列（借入利率）とC列（開始日）のデータを取得
Public Function 借入利率取得(targetSheet As Worksheet) As Variant
    Dim 現在行 As Long
    Dim 結果() As Variant
    Dim 行数 As Long
    Dim i As Long
    
    ' データ行数をカウント
    現在行 = 借入利率開始行
    行数 = 0
    
    Do While targetSheet.Cells(現在行, 借入利率列).Value <> ""
        行数 = 行数 + 1
        現在行 = 現在行 + 1
    Loop
    
    ' データが存在しない場合は空の配列を返す
    If 行数 = 0 Then
        借入利率取得 = Array()
        Exit Function
    End If
    
    ' 結果配列を初期化（行数 x 2列）
    ReDim 結果(1 To 行数, 1 To 2)
    
    ' データを取得
    現在行 = 借入利率開始行
    For i = 1 To 行数
        Dim 利率値 As Variant
        Dim 開始日値 As Variant
        
        ' B列（借入利率）を取得
        利率値 = targetSheet.Cells(現在行, 借入利率列).Value
        If Not IsNumeric(利率値) Then
            Err.Raise 13, "借入利率取得", "借入利率が数値型ではありません。行: " & 現在行
        End If
        結果(i, 1) = CDbl(利率値)
        
        ' C列（開始日）を取得
        開始日値 = targetSheet.Cells(現在行, 借入利率開始日列).Value
        If Not IsDate(開始日値) Then
            Err.Raise 13, "借入利率取得", "開始日が日付型ではありません。行: " & 現在行
        End If
        結果(i, 2) = CDate(開始日値)
        
        現在行 = 現在行 + 1
    Next i
    
    借入利率取得 = 結果
End Function

' 計算期間の最初日を計算する関数
Public Function 計算期間最初日取得(targetSheet As Worksheet) As Date
    Dim 返済予定データ As Variant
    Dim 最初日 As Date
    Dim 入出金データ As Variant
    Dim i As Long
    Dim 入出金最小日付 As Date
    Dim 入出金データ存在 As Boolean
    Dim 返済予定最初日 As Date
    
    ' 返済予定情報を取得
    返済予定データ = 返済予定情報取得(targetSheet)
    
    ' 返済予定データが存在する場合、最初の日付を取得
    If IsArray(返済予定データ) And UBound(返済予定データ, 1) > 0 Then
        返済予定最初日 = 返済予定データ(1, 1) ' 最初の返済予定日
    Else
        返済予定最初日 = DateSerial(1900, 1, 1) ' デフォルト値
    End If
    
    ' 返済予定最初日の前月の1日を初期値として最初日にセット
    最初日 = DateSerial(Year(返済予定最初日), Month(返済予定最初日) - 1, 1)
    
    ' 入出金情報を取得
    入出金データ = 入出金情報取得(targetSheet)
    
    ' 入出金データが存在するかチェック
    入出金データ存在 = IsArray(入出金データ) And UBound(入出金データ, 1) > 0
    
    If 入出金データ存在 Then
        ' 入出金情報の最小日付を取得
        入出金最小日付 = 入出金データ(1, 1) ' 最初の日付で初期化
        For i = 2 To UBound(入出金データ, 1)
            If 入出金データ(i, 1) < 入出金最小日付 Then
                入出金最小日付 = 入出金データ(i, 1)
            End If
        Next i
        
        ' 入出金情報にこの最初日より小さい日付があるかどうか確認
        If 入出金最小日付 < 最初日 Then
            ' あれば、この最初日を返す
            計算期間最初日取得 = 最初日
            Exit Function
        End If
        
        ' 「返済予定最初日」がこの最初日より小さいかをチェック
        If 返済予定最初日 < 最初日 Then
            ' 最初日を返す
            計算期間最初日取得 = 最初日
            Exit Function
        End If
        
        ' 上記以外の場合は、入出金情報の最小日付と「返済予定最初日」を比較して、小さいほうを返す
        If 入出金最小日付 < 返済予定最初日 Then
            計算期間最初日取得 = 入出金最小日付
        Else
            計算期間最初日取得 = 返済予定最初日
        End If
    Else
        ' 入出金データが存在しない場合
        ' 「返済予定最初日」がこの最初日より小さいかをチェック
        If 返済予定最初日 < 最初日 Then
            ' 最初日を返す
            計算期間最初日取得 = 最初日
        Else
            ' 返済予定最初日を返す
            計算期間最初日取得 = 返済予定最初日
        End If
    End If
End Function
