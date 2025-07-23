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

' 遅延損害金利率関連定数
Private Const 遅延損害金利率列 As Long = 2  ' B列：遅延損害金利率
Private Const 遅延損害金利率開始日列 As Long = 3  ' C列：開始日
Private Const 遅延損害金利率開始行 As Long = 15  ' 15行目

' 計算書作成パス関連定数
Private Const 計算書作成パス列 As Long = 3  ' C列：計算書作成パス
Private Const 計算書作成パス行 As Long = 7  ' 7行目

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

' 計算書の作成パス取得関数
' C列7行目から計算書の作成パスを取得し、空白の場合はエラーを発生させる
Public Function 計算書の作成パス取得(targetSheet As Worksheet) As String
    Dim pathValue As Variant
    
    ' C列7行目の値を取得
    pathValue = targetSheet.Cells(計算書作成パス行, 計算書作成パス列).Value
    
    ' 空白チェック
    If pathValue = "" Or IsEmpty(pathValue) Then
        Err.Raise 13, "計算書の作成パス取得", "C列7行目（計算書の作成パス）が空白です。"
    End If
    
    ' 文字列として返す
    計算書の作成パス取得 = CStr(pathValue)
End Function

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

' 正常分出力データ作成関数
' 返済予定情報の2レコード目からループして出力データを作成
Public Function 正常分出力データ作成(targetSheet As Worksheet) As Variant
    Dim 返済予定データ As Variant
    Dim 入出金データ As Variant
    Dim 借入利率データ As Variant
    Dim 遅延損害金利率データ As Variant
    Dim 計算期間最初日 As Date
    Dim 出力結果() As Variant
    Dim 出力行数 As Long
    Dim i As Long, j As Long
    
    ' 1. 返済予定情報取得
    返済予定データ = 返済予定情報取得(targetSheet)
    入出金データ = 入出金情報取得(targetSheet)
    借入利率データ = 借入利率取得(targetSheet)
    遅延損害金利率データ = 遅延損害金利率取得(targetSheet)
    計算期間最初日 = 計算期間最初日取得(targetSheet)
    
    ' データ存在チェック
    If Not IsArray(返済予定データ) Or UBound(返済予定データ, 1) < 2 Then
        Err.Raise 13, "出力データ作成", "返済予定データが不足しています。"
    End If
    
    ' 出力結果配列の初期化（最大想定行数で初期化）
    ReDim 出力結果(1 To 1000, 1 To 13)
    出力行数 = 0
    
    ' 2. 返済予定情報の2レコード目からループ
    For i = 2 To UBound(返済予定データ, 1)
        Dim 返済予定当月データ As Variant
        Dim 返済予定前月データ As Variant
        Dim 期間開始日 As Date
        Dim 期間終了日 As Date
        Dim 分割日リスト() As Date
        Dim 分割日数 As Long
        
        ' 当月と前月のデータを設定
        返済予定当月データ = Array(返済予定データ(i, 1), 返済予定データ(i, 2), 返済予定データ(i, 3))
        返済予定前月データ = Array(返済予定データ(i - 1, 1), 返済予定データ(i - 1, 2), 返済予定データ(i - 1, 3))
        
        ' 3. 計算期間の最初日を取得
        期間開始日 = DateSerial(Year(返済予定当月データ(0)), Month(返済予定当月データ(0)), 1)
        期間開始日 = DateAdd("m", -1, 期間開始日)
        If 期間開始日 < 計算期間最初日 Then
            期間開始日 = 計算期間最初日
        End If
        
        ' 4. 計算期間の最終日を取得（前1か月の月末日）
        期間終了日 = DateSerial(Year(返済予定当月データ(0)), Month(返済予定当月データ(0)), 1)
        期間終了日 = DateAdd("d", -1, 期間終了日)
        
        ' 分割日リストの初期化
        ReDim 分割日リスト(1 To 100)
        分割日数 = 0
        
        ' 6. 未返済年月1の前月が期間内にあるかチェック
        Dim 未返済年月1前月 As Date
        未返済年月1前月 = DateSerial(Year(返済予定前月データ(0)), Month(返済予定前月データ(0)), 1)
        未返済年月1前月 = DateAdd("m", -1, 未返済年月1前月)
        If 未返済年月1前月 >= 期間開始日 And 未返済年月1前月 <= 期間終了日 Then
            分割日数 = 分割日数 + 1
            分割日リスト(分割日数) = 未返済年月1前月
        End If
        
        ' 7. 入出金情報の日付が期間内にあるかチェック
        If IsArray(入出金データ) And UBound(入出金データ, 1) > 0 Then
            For j = 1 To UBound(入出金データ, 1)
                Dim 入出金日 As Date
                入出金日 = 入出金データ(j, 1)
                If 入出金日 >= 期間開始日 And 入出金日 <= 期間終了日 Then
                    ' 既存の分割日と重複しないかチェック
                    Dim 重複フラグ As Boolean
                    重複フラグ = False
                    Dim k As Long
                    For k = 1 To 分割日数
                        If 分割日リスト(k) = 入出金日 Then
                            重複フラグ = True
                            Exit For
                        End If
                    Next k
                    If Not 重複フラグ Then
                        分割日数 = 分割日数 + 1
                        分割日リスト(分割日数) = 入出金日
                    End If
                End If
            Next j
        End If
        
        ' 8. 借入利率データの開始日が期間内にあるかチェック
        If IsArray(借入利率データ) And UBound(借入利率データ, 1) > 0 Then
            For j = 1 To UBound(借入利率データ, 1)
                Dim 借入利率開始日 As Date
                借入利率開始日 = 借入利率データ(j, 2)
                If 借入利率開始日 >= 期間開始日 And 借入利率開始日 <= 期間終了日 Then
                    ' 既存の分割日と重複しないかチェック
                    Dim 重複フラグ2 As Boolean
                    重複フラグ2 = False
                    For k = 1 To 分割日数
                        If 分割日リスト(k) = 借入利率開始日 Then
                            重複フラグ2 = True
                            Exit For
                        End If
                    Next k
                    If Not 重複フラグ2 Then
                        分割日数 = 分割日数 + 1
                        分割日リスト(分割日数) = 借入利率開始日
                    End If
                End If
            Next j
        End If
        
        ' 9. 遅延損害金利率データの開始日が期間内にあるかチェック
        If IsArray(遅延損害金利率データ) And UBound(遅延損害金利率データ, 1) > 0 Then
            For j = 1 To UBound(遅延損害金利率データ, 1)
                Dim 遅延損害金利率開始日 As Date
                遅延損害金利率開始日 = 遅延損害金利率データ(j, 2)
                If 遅延損害金利率開始日 >= 期間開始日 And 遅延損害金利率開始日 <= 期間終了日 Then
                    ' 既存の分割日と重複しないかチェック
                    Dim 重複フラグ3 As Boolean
                    重複フラグ3 = False
                    For k = 1 To 分割日数
                        If 分割日リスト(k) = 遅延損害金利率開始日 Then
                            重複フラグ3 = True
                            Exit For
                        End If
                    Next k
                    If Not 重複フラグ3 Then
                        分割日数 = 分割日数 + 1
                        分割日リスト(分割日数) = 遅延損害金利率開始日
                    End If
                End If
            Next j
        End If
        
        ' 分割日をソート
        If 分割日数 > 1 Then
            Call 分割日ソート(分割日リスト, 分割日数)
        End If
        
        ' 10. 出力レコードの作成
        Dim レコード数 As Long
        レコード数 = IIf(分割日数 = 0, 1, 分割日数 + 1)
        
        For j = 1 To レコード数
            出力行数 = 出力行数 + 1
            
            ' 通番
            出力結果(出力行数, 出力_通番列) = 出力行数
            
            ' ステータス
            出力結果(出力行数, 出力_ステータス列) = "正常"
            
            ' イベント
            出力結果(出力行数, 出力_イベント列) = "約定返済"
            
            ' 約定返済月
            出力結果(出力行数, 出力_約定返済月列) = Format(返済予定当月データ(0), "yyyy/mm")
            
            ' 計算期間開始日
            If j = 1 Then
                出力結果(出力行数, 出力_計算期間開始日列) = 期間開始日
            Else
                出力結果(出力行数, 出力_計算期間開始日列) = 分割日リスト(j - 1)
            End If
            
            ' 計算期間終了日
            If j = レコード数 Then
                出力結果(出力行数, 出力_計算期間終了日列) = 期間終了日
            Else
                出力結果(出力行数, 出力_計算期間終了日列) = DateAdd("d", -1, 分割日リスト(j))
            End If
            
            ' 区切り
            出力結果(出力行数, 出力_区切り列) = "～"
            
            ' 計算日数
            出力結果(出力行数, 出力_計算日数列) = DateDiff("d", 出力結果(出力行数, 出力_計算期間開始日列), 出力結果(出力行数, 出力_計算期間終了日列)) + 1
            
            ' 対象元金の計算
            Dim 対象元金 As Double
            Dim 残高 As Double
            Dim 延滞中約定返済元金 As Double
            
            ' 入出金データから残高と延滞中約定返済元金を取得
            残高 = 0
            延滞中約定返済元金 = 0
            If IsArray(入出金データ) And UBound(入出金データ, 1) > 0 Then
                残高 = 入出金データ(UBound(入出金データ, 1), 4) ' 最後の残高
                延滞中約定返済元金 = 入出金データ(UBound(入出金データ, 1), 5) ' 最後の延滞中約定返済元金
            End If
            
            対象元金 = 残高 - 延滞中約定返済元金
            
            ' 計算期間開始日が返済予定前月データの日付と同じかより大きい場合
            If 出力結果(出力行数, 出力_計算期間開始日列) >= 返済予定前月データ(0) Then
                対象元金 = 対象元金 - 返済予定前月データ(2) ' 返済元金累計を減らす
            End If
            
            出力結果(出力行数, 出力_対象元金列) = 対象元金
            
            ' 利率の取得
            Dim 利率 As Double
            利率 = 0
            If IsArray(借入利率データ) And UBound(借入利率データ, 1) > 0 Then
                For k = 1 To UBound(借入利率データ, 1)
                    If 借入利率データ(k, 2) = 出力結果(出力行数, 出力_計算期間開始日列) Then
                        利率 = 借入利率データ(k, 1)
                        Exit For
                    End If
                Next k
            End If
            出力結果(出力行数, 出力_利率列) = 利率
            
            ' 積数、利息金額、遅延損害金は設定不可（Excel数式あり）
            出力結果(出力行数, 出力_積数列) = ""
            出力結果(出力行数, 出力_利息金額列) = ""
            出力結果(出力行数, 出力_遅延損害金列) = ""
        Next j
    Next i
    
    ' 結果配列のサイズを調整
    If 出力行数 > 0 Then
        ReDim Preserve 出力結果(1 To 出力行数, 1 To 13)
        出力データ作成 = 出力結果
    Else
        出力データ作成 = Array()
    End If
End Function

' 分割日をソートするヘルパー関数
Private Sub 分割日ソート(分割日リスト() As Date, 分割日数 As Long)
    Dim i As Long, j As Long
    Dim temp As Date
    
    For i = 1 To 分割日数 - 1
        For j = i + 1 To 分割日数
            If 分割日リスト(i) > 分割日リスト(j) Then
                temp = 分割日リスト(i)
                分割日リスト(i) = 分割日リスト(j)
                分割日リスト(j) = temp
            End If
        Next j
    Next i
End Sub

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
        If 開始日値 = "" Or IsEmpty(開始日値) Then
            ' 最初のレコードで空白の場合は最小日付をセット
            If i = 1 Then
                結果(i, 2) = DateSerial(1900, 1, 1)
            Else
                Err.Raise 13, "借入利率取得", "開始日が空白です。行: " & 現在行
            End If
        Else
            If Not IsDate(開始日値) Then
                Err.Raise 13, "借入利率取得", "開始日が日付型ではありません。行: " & 現在行
            End If
            結果(i, 2) = CDate(開始日値)
        End If
        
        現在行 = 現在行 + 1
    Next i
    
    借入利率取得 = 結果
End Function

' 遅延損害金利率取得関数
' 指定されたシートの15行目からB列が空白になるまで、B列（遅延損害金利率）とC列（開始日）のデータを取得
Public Function 遅延損害金利率取得(targetSheet As Worksheet) As Variant
    Dim 現在行 As Long
    Dim 結果() As Variant
    Dim 行数 As Long
    Dim i As Long
    
    ' データ行数をカウント
    現在行 = 遅延損害金利率開始行
    行数 = 0
    
    Do While targetSheet.Cells(現在行, 遅延損害金利率列).Value <> ""
        行数 = 行数 + 1
        現在行 = 現在行 + 1
    Loop
    
    ' データが存在しない場合は空の配列を返す
    If 行数 = 0 Then
        遅延損害金利率取得 = Array()
        Exit Function
    End If
    
    ' 結果配列を初期化（行数 x 2列）
    ReDim 結果(1 To 行数, 1 To 2)
    
    ' データを取得
    現在行 = 遅延損害金利率開始行
    For i = 1 To 行数
        Dim 利率値 As Variant
        Dim 開始日値 As Variant
        
        ' B列（遅延損害金利率）を取得
        利率値 = targetSheet.Cells(現在行, 遅延損害金利率列).Value
        If Not IsNumeric(利率値) Then
            Err.Raise 13, "遅延損害金利率取得", "遅延損害金利率が数値型ではありません。行: " & 現在行
        End If
        結果(i, 1) = CDbl(利率値)
        
        ' C列（開始日）を取得
        開始日値 = targetSheet.Cells(現在行, 遅延損害金利率開始日列).Value
        If 開始日値 = "" Or IsEmpty(開始日値) Then
            ' 最初のレコードで空白の場合は最小日付をセット
            If i = 1 Then
                結果(i, 2) = DateSerial(1900, 1, 1)
            Else
                Err.Raise 13, "遅延損害金利率取得", "開始日が空白です。行: " & 現在行
            End If
        Else
            If Not IsDate(開始日値) Then
                Err.Raise 13, "遅延損害金利率取得", "開始日が日付型ではありません。行: " & 現在行
            End If
            結果(i, 2) = CDate(開始日値)
        End If
        
        現在行 = 現在行 + 1
    Next i
    
    遅延損害金利率取得 = 結果
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
    
    ' 返済予定データが存在し、2番目のレコードがある場合、2番目の日付を取得
    If IsArray(返済予定データ) And UBound(返済予定データ, 1) >= 2 Then
        返済予定最初日 = 返済予定データ(2, 1) ' 2番目の返済予定日
    Else
        Err.Raise 13, "計算期間最初日取得", "返済予定情報に2番目のレコードが存在しません。"
    End If
    
    ' 返済予定最初日の前月の1日を初期値として最初日にセット
    最初日 = DateSerial(Year(返済予定最初日), Month(返済予定最初日), 1)
    最初日 = DateAdd("m", -1, 最初日)
    
    ' 入出金情報を取得
    入出金データ = 入出金情報取得(targetSheet)
    
    ' 入出金データが存在するかチェック
    入出金データ存在 = IsArray(入出金データ) And UBound(入出金データ, 1) > 0
    
    ' 未返済前月データとして使用
    Dim 未返済前月 As Date
    未返済前月 = 返済予定データ(1, 1)
    
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
        
        ' 「未返済前月」がこの最初日より小さいかをチェック
        If 未返済前月 < 最初日 Then
            ' 最初日を返す
            計算期間最初日取得 = 最初日
            Exit Function
        End If
        
        ' 上記以外の場合は、入出金情報の最小日付と「未返済前月」を比較して、小さいほうを返す
        If 入出金最小日付 < 未返済前月 Then
            計算期間最初日取得 = 入出金最小日付
        Else
            計算期間最初日取得 = 未返済前月
        End If
    Else
        ' 入出金データが存在しない場合
        ' 「未返済前月」がこの最初日より小さいかをチェック
        If 未返済前月 < 最初日 Then
            ' 最初日を返す
            計算期間最初日取得 = 最初日
        Else
            ' 未返済前月を返す
            計算期間最初日取得 = 未返済前月
        End If
    End If
End Function
