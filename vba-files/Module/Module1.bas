Option Explicit

' 定数定義

' 入出金情報関連定数
Private Const 入出金開始行 As Long = 51 ' 入出金データ開始行
Private Const 入出金日列 As Long = 2     ' B列：入出金日
Private Const 摘要列 As Long = 3         ' C列：摘要
Private Const 入出金金額列 As Long = 4   ' D列：入出金金額
Private Const 残高列 As Long = 5         ' E列：残高
Private Const 約定返済元金列 As Long = 6 ' F列：延滞中の約定返済元金合計

' 期失日関連定数
Private Const 期失日列 As Long = 3  ' C列
Private Const 期失日行 As Long = 30  ' 30行目

' 借入利率関連定数
Private Const 借入利率列 As Long = 2  ' B列：借入利率
Private Const 借入利率開始日列 As Long = 3  ' C列：開始日
Private Const 借入利率開始行 As Long = 34  ' 24行目

' 遅延損害金利率関連定数
Private Const 遅延損害金利率列 As Long = 2  ' B列：遅延損害金利率
Private Const 遅延損害金利率開始日列 As Long = 3  ' C列：開始日
Private Const 遅延損害金利率開始行 As Long = 15  ' 15行目

' 計算書作成パス関連定数
Private Const 計算書作成パス列 As Long = 3  ' C列：計算書作成パス
Private Const 計算書作成パス行 As Long = 7  ' 7行目

' 顧客番号関連定数
Private Const 顧客番号列 As Long = 3  ' C列：顧客番号
Private Const 顧客番号行 As Long = 6  ' 6行目

' 手続理由関連定数
Private Const 手続理由列 As Long = 3  ' C列：手続理由
Private Const 手続理由行 As Long = 10  ' 10行目

' 手続開始日関連定数
Private Const 手続開始日列 As Long = 3  ' C列：手続開始日
Private Const 手続開始日行 As Long = 11  ' 11行目

' ローン口座ステータス関連定数
Private Const ローン口座ステータス列 As Long = 3  ' C列：ローン口座ステータス
Private Const ローン口座ステータス行 As Long = 27  ' 27行目

' 期失理由関連定数
Private Const 期失理由列 As Long = 5  ' E列：期失理由
Private Const 期失理由行 As Long = 30  ' 30行目

' 摘要文字列関連定数
Private Const 借入摘要借入文字列 As String = "借入"     ' 借入を示す摘要文字列
Private Const 借入摘要借換文字列 As String = "借換"     ' 借換を示す摘要文字列
Private Const 返済摘要返済分文字列 As String = "返済分"   ' 返済を示す摘要文字列

' ローン口座ステータス関連定数
Private Const 期失ステータス文字列 As String = "期失"  ' 期失を示すステータス文字列
Private Const 期限切れ理由文字列 As String = "期限切れ"  ' 期限切れを示す理由文字列
Private Const 正常ステータス文字列 As String = "正常"  ' 正常を示すステータス文字列
Private Const 約定返済イベント文字列 As String = "約定返済"  ' 約定返済を示すイベント文字列
Private Const 期失劣後ステータス文字列 As String = "期失（劣後）"  ' 期失（劣後）を示すステータス文字列
Private Const 延滞イベント文字列 As String = "延滞"  ' 延滞を示すイベント文字列
Private Const 内入イベント文字列 As String = "内入"  ' 内入を示すイベント文字列
Private Const 破産イベント文字列 As String = "破産"  ' 破産イベント文字列
Private Const 利息摘要文字列 As String = "利息"  ' 利息を示す摘要文字列
Private Const 遅延損害金摘要文字列 As String = "遅延損害金"  ' 遅延損害金を示す摘要文字列

' 日付関連定数
Private Const 日付初期値 As Date = #1/1/1900#  ' 空白日付の初期値

' ワークシート名関連定数
Private Const ツールシート名 As String = "ツール"  ' ツールシートの名前
Private Const テンプレートシート名 As String = "テンプレート_EXCEL"  ' テンプレートシートの名前

' 出力関連定数
Private Const 出力開始行オフセット As Long = 8  ' A9セルから貼り付けるためのオフセット
Private Const 出力顧客番号行 As Long = 4  ' B4行：顧客番号
Private Const 出力顧客番号列 As Long = 2  ' B列：顧客番号
Private Const 出力手続開始日行 As Long = 2  ' J2行：手続開始日
Private Const 出力手続開始日列 As Long = 10  ' J列：手続開始日
Private Const 出力期失日行 As Long = 3  ' J3行：期失日
Private Const 出力期失日列 As Long = 10  ' J列：期失日
Private Const 出力期失理由行 As Long = 3  ' K3行：期失理由
Private Const 出力期失理由列 As Long = 11  ' K列：期失理由

' 出力項目列定数（A列～S列）
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
Private Const 出力_借入日列 As Long = 14         ' N列：借入日
Private Const 出力_借入額列 As Long = 15         ' O列：借入額
Private Const 出力_返済日列 As Long = 16         ' P列：返済日
Private Const 出力_元金_返済額列 As Long = 17    ' Q列：元金_返済額
Private Const 出力_利息_返済額列 As Long = 18    ' R列：利息_返済額
Private Const 出力_遅損金_返済額列 As Long = 19  ' S列：遅損金_返済額

' 返済予定情報の定数
Const 返済予定開始行 As Long = 40  ' 40行目
Const 返済予定日列 As Long = 3  ' C列
Const 返済元金列 As Long = 4    ' D列

' 返済履歴情報の定数
Const 返済履歴開始行 As Long = 75   ' 75行目
Const 返済履歴日付列 As Long = 2    ' B列：日付
Const 返済履歴摘要列 As Long = 3    ' C列：摘要
Const 返済履歴出金金額列 As Long = 4 ' D列：出金金額

' 削除最後行目の定数
Const 削除最後行目 As Long = 69

' データ貼り付け開始行の定数
Const データ貼り付け開始行 As Long = 出力開始行オフセット + 1

' 期失理由取得関数
' E列25行目から期失理由を取得し、ローン口座ステータスに応じて処理を分岐
Public Function 期失理由取得(targetSheet As Worksheet) As String
    Dim cellValue As Variant
    Dim ローン口座ステータス As String
    
    ' E列25行目の値を取得
    cellValue = targetSheet.Cells(期失理由行, 期失理由列).Value
    
    ' ローン口座ステータスを取得
    ローン口座ステータス = ローン口座ステータス取得(targetSheet)
    
    ' 空白チェック
    If cellValue = "" Or IsEmpty(cellValue) Then
        ' ローン口座ステータスが「期失」の場合は入力必須
        If ローン口座ステータス = 期失ステータス文字列 Then
            Err.Raise 13, "期失理由取得", "期失理由が設定されていません。E25セルに期失理由を入力してください。"
        Else
            ' それ以外の場合は「期限切れ」を設定
            期失理由取得 = 期限切れ理由文字列
            Exit Function
        End If
    End If
    
    ' 文字列として変換
    期失理由取得 = CStr(cellValue)
End Function

' ローン口座ステータス取得関数
' C列22行目からローン口座ステータスを取得し、空白の場合はエラーを発生させる
Public Function ローン口座ステータス取得(targetSheet As Worksheet) As String
    Dim cellValue As Variant
    
    ' C列22行目の値を取得
    cellValue = targetSheet.Cells(ローン口座ステータス行, ローン口座ステータス列).Value
    
    ' 空白チェック
    If cellValue = "" Or IsEmpty(cellValue) Then
        Err.Raise 13, "ローン口座ステータス取得", "ローン口座ステータスが設定されていません。C22セルにローン口座ステータスを入力してください。"
    End If
    
    ' 文字列として変換
    ローン口座ステータス取得 = CStr(cellValue)
End Function

' 手続開始日取得関数
' C列11行目から手続開始日を取得し、空白の場合はエラーを発生させる
Public Function 手続開始日取得(targetSheet As Worksheet) As Date
    Dim cellValue As Variant
    
    ' C列11行目の値を取得
    cellValue = targetSheet.Cells(手続開始日行, 手続開始日列).Value
    
    ' 空白チェック
    If cellValue = "" Or IsEmpty(cellValue) Then
        Err.Raise 13, "手続開始日取得", "手続開始日が設定されていません。C11セルに手続開始日を入力してください。"
    End If
    
    ' 日付型チェック
    If Not IsDate(cellValue) Then
        Err.Raise 13, "手続開始日取得", "手続開始日が日付ではありません。C11セルに正しい日付を入力してください。"
    End If
    
    ' 日付型として変換
    手続開始日取得 = CDate(cellValue)
End Function

' 手続理由取得関数
' C列10行目から手続理由を取得し、空白の場合はエラーを発生させる
Public Function 手続理由取得(targetSheet As Worksheet) As String
    Dim cellValue As Variant
    
    ' C列10行目の値を取得
    cellValue = targetSheet.Cells(手続理由行, 手続理由列).Value
    
    ' 空白チェック
    If cellValue = "" Or IsEmpty(cellValue) Then
        Err.Raise 13, "手続理由取得", "手続理由が設定されていません。C10セルに手続理由を入力してください。"
    End If
    
    ' 文字列として変換
    手続理由取得 = CStr(cellValue)
End Function

' 顧客番号取得関数
' C列6行目から顧客番号を取得し、空白の場合はエラーを発生させる
Public Function 顧客番号取得(targetSheet As Worksheet) As String
    Dim cellValue As Variant
    
    ' C列6行目の値を取得
    cellValue = targetSheet.Cells(顧客番号行, 顧客番号列).Value
    
    ' 空白チェック
    If cellValue = "" Or IsEmpty(cellValue) Then
        Err.Raise 13, "顧客番号取得", "顧客番号が設定されていません。C6セルに顧客番号を入力してください。"
    End If
    
    ' 文字列として変換
    顧客番号取得 = CStr(cellValue)
End Function

' 計算書の作成パス取得関数
' C列7行目から計算書の作成パスを取得し、空白の場合はエラーを発生させる
' パスが存在するフォルダでない場合もエラーを発生させる
Public Function 計算書の作成パス取得(targetSheet As Worksheet) As String
    Dim pathValue As Variant
    Dim pathString As String
    
    ' C列7行目の値を取得
    pathValue = targetSheet.Cells(計算書作成パス行, 計算書作成パス列).Value
    
    ' 空白チェック
    If pathValue = "" Or IsEmpty(pathValue) Then
        Err.Raise 13, "計算書の作成パス取得", "C列7行目（計算書の作成パス）が空白です。"
    End If
    
    ' 文字列として変換
    pathString = CStr(pathValue)
    
    ' パスの妥当性チェック（フォルダが存在するかチェック）
    If Dir(pathString, vbDirectory) = "" Then
        Err.Raise 76, "計算書の作成パス取得", "指定されたパス '" & pathString & "' はフォルダではありません。"
    End If
    
    ' 文字列として返す
    計算書の作成パス取得 = pathString
End Function

' ファイル出力関数
' 計算書の作成パス取得で取得したフォルダに利息計算書ファイルを作成し、正常分出力データを貼り付ける
Public Sub ファイル出力(targetSheet As Worksheet, templateSheet As Worksheet)
    Dim 出力フォルダパス As String
    Dim 出力データ As Variant
    Dim 新しいワークブック As Workbook
    Dim 新しいワークシート As Worksheet
    Dim ファイル名 As String
    Dim 完全ファイルパス As String
    Dim 現在日時 As Date
    Dim 年月日文字列 As String
    Dim 時分秒文字列 As String
    
    On Error GoTo ErrorHandler
    
    ' 1. 計算書の作成パス取得
    出力フォルダパス = 計算書の作成パス取得(targetSheet)
    
    ' 2. 出力データ作成
    出力データ = 出力データ作成(targetSheet)
    
    ' 3. 現在日時を取得してファイル名を作成
    現在日時 = Now
    年月日文字列 = Format(現在日時, "yyyymmdd")
    時分秒文字列 = Format(現在日時, "hhmmss")
    Dim 顧客番号 As String
    顧客番号 = 顧客番号取得(targetSheet)
    ファイル名 = "利息計算書" & 顧客番号 & ".xlsx"
    
    ' 4. 完全ファイルパスを作成
    If Right(出力フォルダパス, 1) <> "\" Then
        完全ファイルパス = 出力フォルダパス & "\" & ファイル名
    Else
        完全ファイルパス = 出力フォルダパス & ファイル名
    End If
    
    ' 5. 新しいワークブックを作成
    Set 新しいワークブック = Workbooks.Add
    
    ' 6. テンプレートシートを新しいワークブックにコピー
    ' 非表示シートの場合、一時的に表示してからコピー
    Dim 元の表示状態 As XlSheetVisibility
    元の表示状態 = templateSheet.Visible
    If templateSheet.Visible <> xlSheetVisible Then
        templateSheet.Visible = xlSheetVisible
    End If
    
    templateSheet.Copy Before:=新しいワークブック.Worksheets(1)
    
    ' 元の表示状態に戻す
    templateSheet.Visible = 元の表示状態
    Set 新しいワークシート = 新しいワークブック.Worksheets(1)
    新しいワークシート.Name = "利息計算書"
    
    ' 元のSheet1を削除
    Application.DisplayAlerts = False
    新しいワークブック.Worksheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
    ' 7. データをA9セルから貼り付け
    If IsArray(出力データ) Then
        Dim 行数 As Long
        Dim 列数 As Long
        行数 = UBound(出力データ, 1)
        列数 = UBound(出力データ, 2)
        
        ' 画面更新を停止してパフォーマンスを向上
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        ' 不要な行を削除（コピーした最後の行から削除最後行目まで削除）
        Dim 最後の行 As Long
        最後の行 = データ貼り付け開始行 + 行数 - 2  ' データ貼り付け開始行から貼り付けた最後の行(「計」行除く)
        Dim 削除開始行 As Long
        削除開始行 = 最後の行 + 1
        
        ' 削除範囲が存在する場合のみ削除
        If 削除開始行 <= 削除最後行目 Then
            新しいワークシート.Rows(削除開始行 & ":" & 削除最後行目).Delete
        End If

        ' メモリ効率を考慮してセル範囲を指定して貼り付け
        Dim 貼り付け範囲 As Range
        Set 貼り付け範囲 = 新しいワークシート.Range("A" & データ貼り付け開始行).Resize(行数, 列数)
        
        ' データを貼り付け
        貼り付け範囲.Value = 出力データ
        
        ' 画面更新を再開
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If
    
    ' 7.5. 顧客番号設定
    新しいワークシート.Cells(出力顧客番号行, 出力顧客番号列).Value = "顧客番号" & 顧客番号
    
    ' 7.6. 手続開始日設定
    Dim 手続開始日 As Date
    手続開始日 = 手続開始日取得(targetSheet)
    新しいワークシート.Cells(出力手続開始日行, 出力手続開始日列).Value = 手続開始日
    
    ' 7.7. 期失日設定
    Dim 期失日 As Date
    期失日 = 期失日取得(targetSheet)
    新しいワークシート.Cells(出力期失日行, 出力期失日列).Value = 期失日
    
    ' 7.8. 期失理由設定
    Dim 期失理由 As String
    期失理由 = 期失理由取得(targetSheet)
    新しいワークシート.Cells(出力期失理由行, 出力期失理由列).Value = 期失理由
    
    ' 8. ファイル保存（既存ファイルがある場合は連番付きで保存）
    Dim 保存ファイルパス As String
    Dim カウンタ As Long
    保存ファイルパス = 完全ファイルパス
    カウンタ = 1
    
    ' 既存ファイルがある場合は連番を付けて新しいファイル名を作成
    Do While Dir(保存ファイルパス) <> ""
        Dim ファイル名部分 As String
        Dim 拡張子部分 As String
        Dim フォルダパス部分 As String
        
        ' ファイルパスを分解
        フォルダパス部分 = Left(完全ファイルパス, InStrRev(完全ファイルパス, "\"))
        ファイル名部分 = Mid(完全ファイルパス, InStrRev(完全ファイルパス, "\") + 1)
        拡張子部分 = Right(ファイル名部分, 5) ' ".xlsx"
        ファイル名部分 = Left(ファイル名部分, Len(ファイル名部分) - 5)
        
        ' 連番付きファイル名を作成
        保存ファイルパス = フォルダパス部分 & ファイル名部分 & "(" & カウンタ & ")" & 拡張子部分
        カウンタ = カウンタ + 1
    Loop
    
    新しいワークブック.SaveAs Filename:=保存ファイルパス, FileFormat:=xlOpenXMLWorkbook
    
    ' 9. ワークブックを閉じる
    新しいワークブック.Close SaveChanges:=False
    
    ' 10. 完了メッセージ
    MsgBox "利息計算書ファイルの出力が完了しました。" & vbCrLf & "保存先: " & 保存ファイルパス, vbInformation, "ファイル出力完了"
    
    Exit Sub
    
ErrorHandler:
    ' エラーが発生した場合はワークブックを閉じる
    If Not 新しいワークブック Is Nothing Then
        新しいワークブック.Close SaveChanges:=False
    End If
    
    ' エラーメッセージを表示
    MsgBox "ファイル出力中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
    Err.Raise Err.Number, "ファイル出力", Err.Description
End Sub

' 計算書作成メイン処理
' ツールシートを対象としてファイル出力を実行する
Public Sub 計算書作成()
    Dim ツールシート As Worksheet
    Dim テンプレートシート As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' ツールシートを取得
    Set ツールシート = ThisWorkbook.Worksheets(ツールシート名)
    
    ' テンプレートシートを取得
    Set テンプレートシート = ThisWorkbook.Worksheets(テンプレートシート名)
    
    ' ファイル出力を実行
    Call ファイル出力(ツールシート, テンプレートシート)
    
    Exit Sub
    
ErrorHandler:
    'MsgBox "計算書作成中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' 返済予定情報取得関数
' 35行目から開始し、C列が空白になるまで返済予定日、返済元金、返済元金累計を取得
Public Function 返済予定情報取得(targetSheet As Worksheet) As Variant
    Dim currentRow As Long
    Dim dataArray() As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim 返済元金累計 As Double
    
    ' 35行目のC列が空白の場合も1行として処理
    currentRow = 返済予定開始行 + 1
    rowCount = 1
    
    ' データ行数をカウント（C列が空白になるまで）
    Do While targetSheet.Cells(currentRow, 返済予定日列).Value <> ""
        rowCount = rowCount + 1
        currentRow = currentRow + 1
    Loop
        
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
            dataArray(i, 1) = 日付初期値
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

' 出力データ作成関数
' 返済予定情報の2レコード目からループして出力データを作成
Public Function 出力データ作成(targetSheet As Worksheet) As Variant
    Dim 返済予定データ As Variant
    Dim 入出金データ As Variant
    Dim 借入利率データ As Variant
    Dim 遅延損害金利率データ As Variant
    Dim 計算期間最初日 As Date
    Dim 出力結果() As Variant
    Dim 出力行数 As Long
    Dim i As Long, j As Long, k As Long
    Dim 摘要 As String
    
    ' 1. 返済予定情報取得
    返済予定データ = 返済予定情報取得(targetSheet)
    入出金データ = 入出金情報全体取得(targetSheet)
    借入利率データ = 借入利率取得(targetSheet)
    遅延損害金利率データ = 遅延損害金利率取得(targetSheet)
    計算期間最初日 = 計算期間最初日取得(targetSheet)
    
    ' データ存在チェック
    If Not IsArray(返済予定データ) Or UBound(返済予定データ, 1) < 2 Then
        Err.Raise 13, "出力データ作成", "返済予定データが不足しています。"
    End If
    
    ' 出力結果配列の初期化（最大想定行数で初期化）
    ReDim 出力結果(1 To 1000, 1 To 19)
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
        
        ' 期失日以降であれば、期失日の前日に設定
        ' ただし、ローン口座ステータスが「期限切れ」の場合は期失日そのものを設定
        Dim 期失日 As Date
        Dim ローン口座ステータス As String
        期失日 = 期失日取得(targetSheet)
        ローン口座ステータス = ローン口座ステータス取得(targetSheet)
        
        If 期間終了日 >= 期失日 Then
            If ローン口座ステータス = 期限切れ理由文字列 Then
                期間終了日 = 期失日
            Else
                期間終了日 = DateAdd("d", -1, 期失日)
            End If
        End If
        
        ' 分割日リストの初期化
        ReDim 分割日リスト(1 To 100)
        分割日数 = 0
        
        ' 6. 返済予定前月データの日付が期間内にあるかチェック
        Dim 返済予定前月日付 As Date
        返済予定前月日付 = 返済予定前月データ(0)
        If 返済予定前月日付 > 期間開始日 And 返済予定前月日付 <= 期間終了日 Then
            分割日数 = 分割日数 + 1
            分割日リスト(分割日数) = 返済予定前月日付
        End If
        
        ' 7. 入出金情報の日付が期間内にあるかチェック
        If IsArray(入出金データ) And UBound(入出金データ, 1) > 0 Then
            For j = 1 To UBound(入出金データ, 1)
                ' 摘要が「返済分」で終わる場合は判断対象外とする
                摘要 = CStr(入出金データ(j, 2))
                ' 返済分の場合はスキップ
                If Right(摘要, Len(返済摘要返済分文字列)) <> 返済摘要返済分文字列 Then
                    Dim 入出金日 As Date
                    入出金日 = 入出金データ(j, 1)
                    If 入出金日 > 期間開始日 And 入出金日 <= 期間終了日 Then
                        
                        ' 既存の分割日と重複しないかチェック
                        Dim 重複フラグ As Boolean
                        重複フラグ = False
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
                    
                End If
            Next j
        End If
        
        ' 8. 借入利率データの開始日が期間内にあるかチェック
        If IsArray(借入利率データ) And UBound(借入利率データ, 1) > 0 Then
            For j = 1 To UBound(借入利率データ, 1)
                Dim 借入利率開始日 As Date
                借入利率開始日 = 借入利率データ(j, 2)
                If 借入利率開始日 > 期間開始日 And 借入利率開始日 <= 期間終了日 Then
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
                If 遅延損害金利率開始日 > 期間開始日 And 遅延損害金利率開始日 <= 期間終了日 Then
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
            出力結果(出力行数, 出力_ステータス列) = 正常ステータス文字列
            
            ' イベント
            出力結果(出力行数, 出力_イベント列) = 約定返済イベント文字列
            
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
            出力結果(出力行数, 出力_計算日数列) = "=H" & (出力行数 + 出力開始行オフセット) & "-F" & (出力行数 + 出力開始行オフセット) & "+1"
            
            ' 対象元金の計算
            Dim 対象元金 As Double
            Dim 残高 As Double
            Dim 延滞中約定返済元金 As Double
            
            ' 入出金データから計算期間開始日と同じかより小さい日付の中で最大日付のデータを取得
            残高 = 0
            延滞中約定返済元金 = 0
            If IsArray(入出金データ) And UBound(入出金データ, 1) > 0 Then
                Dim 計算期間開始日_対象元金 As Date
                計算期間開始日_対象元金 = 出力結果(出力行数, 出力_計算期間開始日列)
                
                Dim 最大日付_入出金 As Date
                Dim 最大日付見つかった As Boolean
                最大日付_入出金 = 日付初期値
                最大日付見つかった = False
                
                ' 計算期間開始日と同じかより小さい日付の中で最大日付を探す
                For k = 1 To UBound(入出金データ, 1)
                    If 入出金データ(k, 1) <= 計算期間開始日_対象元金 And 入出金データ(k, 1) > 最大日付_入出金 Then
                        最大日付_入出金 = 入出金データ(k, 1)
                        残高 = 入出金データ(k, 4)
                        延滞中約定返済元金 = 入出金データ(k, 5)
                        最大日付見つかった = True
                    End If
                Next k
                
                ' 該当するデータが見つからない場合は最初のデータを使用
                If Not 最大日付見つかった And UBound(入出金データ, 1) > 0 Then
                    残高 = 入出金データ(1, 4)
                    延滞中約定返済元金 = 入出金データ(1, 5)
                End If
            End If
            
            対象元金 = 残高 - 延滞中約定返済元金
            
            ' 返済予定情報から計算期間開始日と同じかより小さい日付の中で最大日付のデータの返済元金累計を減らす
            If IsArray(返済予定データ) And UBound(返済予定データ, 1) > 0 Then
                Dim 最大日付_返済予定 As Date
                Dim 返済元金累計_減算 As Double
                最大日付_返済予定 = 日付初期値
                返済元金累計_減算 = 0
                
                For k = 1 To UBound(返済予定データ, 1)
                    If 返済予定データ(k, 1) <= 計算期間開始日_対象元金 And 返済予定データ(k, 1) > 最大日付_返済予定 Then
                        最大日付_返済予定 = 返済予定データ(k, 1)
                        返済元金累計_減算 = 返済予定データ(k, 3) ' 返済元金累計
                    End If
                Next k
                
                対象元金 = 対象元金 - 返済元金累計_減算
            End If
            
            出力結果(出力行数, 出力_対象元金列) = 対象元金
            
            ' 利率の取得
            Dim 利率 As Double
            Dim 利率見つかった As Boolean
            利率 = 0
            利率見つかった = False
            
            If IsArray(借入利率データ) And UBound(借入利率データ, 1) > 0 Then
                Dim 計算期間開始日 As Date
                計算期間開始日 = 出力結果(出力行数, 出力_計算期間開始日列)
                
                ' まず計算期間開始日と同じ日付のデータを探す
                For k = 1 To UBound(借入利率データ, 1)
                    If 借入利率データ(k, 2) = 計算期間開始日 Then
                        利率 = 借入利率データ(k, 1)
                        利率見つかった = True
                        Exit For
                    End If
                Next k
                
                ' 同じ日付のデータがない場合、計算期間開始日より小さい日付の中で最も大きい日付を探す（初回は任意の日付を受け入れ）
                If Not 利率見つかった Then
                    Dim 最大日付 As Date
                    最大日付 = 日付初期値 ' 初期値として最小日付を設定
                    
                    For k = 1 To UBound(借入利率データ, 1)
                        If 借入利率データ(k, 2) < 計算期間開始日 And (借入利率データ(k, 2) > 最大日付 Or (借入利率データ(k, 2) = 最大日付 And 最大日付 = 日付初期値)) Then
                            最大日付 = 借入利率データ(k, 2)
                            利率 = 借入利率データ(k, 1)
                            利率見つかった = True
                        End If
                    Next k
                End If
            End If
            出力結果(出力行数, 出力_利率列) = 利率
            
            ' 積数の数式設定（対象元金×利率×計算日数）
            出力結果(出力行数, 出力_積数列) = "=E" & (出力行数 + 出力開始行オフセット) & "*J" & (出力行数 + 出力開始行オフセット) & "*I" & (出力行数 + 出力開始行オフセット)
            
            ' 利息金額の数式設定
            Dim 現在行番号 As Long
            現在行番号 = 出力行数 + 出力開始行オフセット
            
            If j = 1 Then
                ' J=1の場合：=ROUNDDOWN(K行番号/365,0)
                出力結果(出力行数, 出力_利息金額列) = "=ROUNDDOWN(K" & 現在行番号 & "/365,0)"
            Else
                ' J=1以外の場合：=ROUNDDOWN(SUM(K(J=1時の行番号):K現在の行番号)/365,0)-SUM(L(J=1時の行番号):L現在の行番号-1)
                Dim J1開始行番号 As Long
                J1開始行番号 = (出力行数 - j + 1) + 出力開始行オフセット ' J=1時の行番号を計算
                出力結果(出力行数, 出力_利息金額列) = "=ROUNDDOWN(SUM(K" & J1開始行番号 & ":K" & 現在行番号 & ")/365,0)-SUM(L" & J1開始行番号 & ":L" & (現在行番号 - 1) & ")"
            End If
            
            ' 遅延損害金は空白
            出力結果(出力行数, 出力_遅延損害金列) = ""
            
            ' 借入・返済情報の設定
            Dim 入出金情報全体データ_借入返済 As Variant
            入出金情報全体データ_借入返済 = 入出金情報全体取得(targetSheet)
            
            If IsArray(入出金情報全体データ_借入返済) And UBound(入出金情報全体データ_借入返済, 1) > 0 Then
                Dim 借入設定済み As Boolean
                Dim 返済設定済み As Boolean
                借入設定済み = False
                返済設定済み = False
                
                For k = 1 To UBound(入出金情報全体データ_借入返済, 1)
                    ' 計算期間開始日と一致する場合のみ処理
                    If 入出金情報全体データ_借入返済(k, 1) = 出力結果(出力行数, 出力_計算期間開始日列) Then
                        摘要 = CStr(入出金情報全体データ_借入返済(k, 2))
                        
                        ' 借入情報の設定（借入または借換）
                        If Not 借入設定済み And ((Len(摘要) >= Len(借入摘要借入文字列) And Right(摘要, Len(借入摘要借入文字列)) = 借入摘要借入文字列) Or (Len(摘要) >= Len(借入摘要借換文字列) And Right(摘要, Len(借入摘要借換文字列)) = 借入摘要借換文字列)) Then
                            出力結果(出力行数, 出力_借入日列) = 入出金情報全体データ_借入返済(k, 1)
                            出力結果(出力行数, 出力_借入額列) = 入出金情報全体データ_借入返済(k, 3)
                            借入設定済み = True
                        End If
                        
                        ' 返済情報の設定
                        If Not 返済設定済み And Len(摘要) >= Len(返済摘要返済分文字列) And Right(摘要, Len(返済摘要返済分文字列)) = 返済摘要返済分文字列 Then
                            出力結果(出力行数, 出力_返済日列) = 入出金情報全体データ_借入返済(k, 1)
                            出力結果(出力行数, 出力_元金_返済額列) = 入出金情報全体データ_借入返済(k, 3)
                            返済設定済み = True
                        End If
                        
                        ' 両方設定済みの場合はループを終了
                        If 借入設定済み And 返済設定済み Then
                            Exit For
                        End If
                    End If
                Next k
            End If

            
        Next j
        
        ' 11. 延滞分レコードの値を設定する
        Dim 期失日_延滞 As Date
        期失日_延滞 = 期失日取得(targetSheet)
        
        ' 返済予定データの日付が期失日より小さい場合のみ延滞分レコードを作成
        If 返済予定当月データ(0) < 期失日_延滞 Then
            ' 延滞分レコードの期間設定
            Dim 延滞期間開始日 As Date
            Dim 延滞期間終了日 As Date
            延滞期間開始日 = 返済予定当月データ(0)
            
            ' 計算期間終了日（期失日の前日、ただしローン口座ステータスが「期限切れ」の場合は期失日そのもの）
            If ローン口座ステータス = 期限切れ理由文字列 Then
                延滞期間終了日 = 期失日_延滞
            Else
                延滞期間終了日 = DateAdd("d", -1, 期失日_延滞)
            End If
            
            ' 延滞分レコード用の分割日リストを作成
            Dim 延滞分割日リスト(1 To 100) As Date
            Dim 延滞分割日数 As Long
            延滞分割日数 = 0
            
            ' 遅延損害金利率データの開始日が期間内にあるかチェック
            If IsArray(遅延損害金利率データ) And UBound(遅延損害金利率データ, 1) > 0 Then
                For j = 1 To UBound(遅延損害金利率データ, 1)
                    Dim 遅延損害金利率開始日_延滞 As Date
                    遅延損害金利率開始日_延滞 = 遅延損害金利率データ(j, 2)
                    If 遅延損害金利率開始日_延滞 > 延滞期間開始日 And 遅延損害金利率開始日_延滞 <= 延滞期間終了日 Then
                        ' 既存の分割日と重複しないかチェック
                        Dim 重複フラグ_延滞 As Boolean
                        重複フラグ_延滞 = False
                        For k = 1 To 延滞分割日数
                            If 延滞分割日リスト(k) = 遅延損害金利率開始日_延滞 Then
                                重複フラグ_延滞 = True
                                Exit For
                            End If
                        Next k
                        If Not 重複フラグ_延滞 Then
                            延滞分割日数 = 延滞分割日数 + 1
                            延滞分割日リスト(延滞分割日数) = 遅延損害金利率開始日_延滞
                        End If
                    End If
                Next j
            End If
            
            ' 分割日リストをソート
            If 延滞分割日数 > 1 Then
                For j = 1 To 延滞分割日数 - 1
                    For k = j + 1 To 延滞分割日数
                        If 延滞分割日リスト(j) > 延滞分割日リスト(k) Then
                            Dim temp_延滞 As Date
                            temp_延滞 = 延滞分割日リスト(j)
                            延滞分割日リスト(j) = 延滞分割日リスト(k)
                            延滞分割日リスト(k) = temp_延滞
                        End If
                    Next k
                Next j
            End If
            
            ' 延滞分レコードを分割して作成
            Dim 延滞セグメント開始日 As Date
            Dim 延滞セグメント終了日 As Date
            延滞セグメント開始日 = 延滞期間開始日
            
            For j = 0 To 延滞分割日数
                ' セグメント終了日の設定
                If j = 延滞分割日数 Then
                    延滞セグメント終了日 = 延滞期間終了日
                Else
                    延滞セグメント終了日 = DateAdd("d", -1, 延滞分割日リスト(j + 1))
                End If
                
                ' セグメントが有効な場合のみレコードを作成
                If 延滞セグメント開始日 <= 延滞セグメント終了日 Then
                    出力行数 = 出力行数 + 1
                    
                    ' 通番
                    出力結果(出力行数, 出力_通番列) = 出力行数
                    
                    ' ステータス
                    出力結果(出力行数, 出力_ステータス列) = 延滞イベント文字列
                    
                    ' イベント
                    出力結果(出力行数, 出力_イベント列) = 約定返済イベント文字列
                    
                    ' 約定返済月
                    出力結果(出力行数, 出力_約定返済月列) = Format(返済予定当月データ(0), "yyyy/mm")
                    
                    ' 計算期間開始日
                    出力結果(出力行数, 出力_計算期間開始日列) = 延滞セグメント開始日
                    
                    ' 計算期間終了日
                    出力結果(出力行数, 出力_計算期間終了日列) = 延滞セグメント終了日
                    
                    ' 区切り
                    出力結果(出力行数, 出力_区切り列) = "～"
                    
                    ' 計算日数
                    出力結果(出力行数, 出力_計算日数列) = "=H" & (出力行数 + 出力開始行オフセット) & "-F" & (出力行数 + 出力開始行オフセット) & "+1"
                    
                    ' 対象元金
                    出力結果(出力行数, 出力_対象元金列) = 返済予定当月データ(1) ' 返済元金
                    
                    ' 遅延損害金利率の取得
                    Dim 遅延損害金利率_延滞 As Double
                    Dim 遅延損害金利率見つかった_延滞 As Boolean
                    遅延損害金利率_延滞 = 0
                    遅延損害金利率見つかった_延滞 = False
                    
                    If IsArray(遅延損害金利率データ) And UBound(遅延損害金利率データ, 1) > 0 Then
                        ' まず計算期間開始日と同じ日付のデータを探す
                        For k = 1 To UBound(遅延損害金利率データ, 1)
                            If 遅延損害金利率データ(k, 2) = 延滞セグメント開始日 Then
                                遅延損害金利率_延滞 = 遅延損害金利率データ(k, 1)
                                遅延損害金利率見つかった_延滞 = True
                                Exit For
                            End If
                        Next k
                        
                        ' 同じ日付のデータがない場合、計算期間開始日より小さい日付の中で最も大きい日付を探す
                        If Not 遅延損害金利率見つかった_延滞 Then
                            Dim 最大日付_遅延損害金_延滞 As Date
                            最大日付_遅延損害金_延滞 = 日付初期値
                            
                            For k = 1 To UBound(遅延損害金利率データ, 1)
                                If 遅延損害金利率データ(k, 2) < 延滞セグメント開始日 And (遅延損害金利率データ(k, 2) > 最大日付_遅延損害金_延滞 Or (遅延損害金利率データ(k, 2) = 最大日付_遅延損害金_延滞 And 最大日付_遅延損害金_延滞 = 日付初期値)) Then
                                    最大日付_遅延損害金_延滞 = 遅延損害金利率データ(k, 2)
                                    遅延損害金利率_延滞 = 遅延損害金利率データ(k, 1)
                                    遅延損害金利率見つかった_延滞 = True
                                End If
                            Next k
                        End If
                    End If
                    出力結果(出力行数, 出力_利率列) = 遅延損害金利率_延滞
                    
                    ' 積数の数式設定（対象元金×利率×計算日数）
                    出力結果(出力行数, 出力_積数列) = "=E" & (出力行数 + 出力開始行オフセット) & "*J" & (出力行数 + 出力開始行オフセット) & "*I" & (出力行数 + 出力開始行オフセット)
                    
                    ' 利息金額は空白
                    出力結果(出力行数, 出力_利息金額列) = ""
                    
                    ' 遅延損害金の数式設定（積数/365）
                    出力結果(出力行数, 出力_遅延損害金列) = "=ROUNDDOWN(K" & (出力行数 + 出力開始行オフセット) & "/365,0)"
                    
                    ' 借入・返済情報の設定（最初のセグメントのみ）
                    If j = 0 Then
                        Dim 入出金情報全体データ_借入返済_延滞 As Variant
                        入出金情報全体データ_借入返済_延滞 = 入出金情報全体取得(targetSheet)
                        
                        If IsArray(入出金情報全体データ_借入返済_延滞) And UBound(入出金情報全体データ_借入返済_延滞, 1) > 0 Then
                            Dim k_延滞 As Long
                            Dim 借入設定済み_延滞 As Boolean
                            Dim 返済設定済み_延滞 As Boolean
                            借入設定済み_延滞 = False
                            返済設定済み_延滞 = False
                            
                            For k_延滞 = 1 To UBound(入出金情報全体データ_借入返済_延滞, 1)
                                ' 計算期間開始日と一致する場合のみ処理
                                If 入出金情報全体データ_借入返済_延滞(k_延滞, 1) = 延滞セグメント開始日 Then
                                    Dim 摘要_延滞 As String
                                    摘要_延滞 = CStr(入出金情報全体データ_借入返済_延滞(k_延滞, 2))
                                    
                                    ' 借入情報の設定（借入または借換）
                                    If Not 借入設定済み_延滞 And ((Len(摘要_延滞) >= Len(借入摘要借入文字列) And Right(摘要_延滞, Len(借入摘要借入文字列)) = 借入摘要借入文字列) Or (Len(摘要_延滞) >= Len(借入摘要借換文字列) And Right(摘要_延滞, Len(借入摘要借換文字列)) = 借入摘要借換文字列)) Then
                                        出力結果(出力行数, 出力_借入日列) = 入出金情報全体データ_借入返済_延滞(k_延滞, 1)
                                        出力結果(出力行数, 出力_借入額列) = 入出金情報全体データ_借入返済_延滞(k_延滞, 3)
                                        借入設定済み_延滞 = True
                                    End If
                                    
                                    ' 返済情報の設定
                                    If Not 返済設定済み_延滞 And Len(摘要_延滞) >= Len(返済摘要返済分文字列) And Right(摘要_延滞, Len(返済摘要返済分文字列)) = 返済摘要返済分文字列 Then
                                        出力結果(出力行数, 出力_返済日列) = 入出金情報全体データ_借入返済_延滞(k_延滞, 1)
                                        出力結果(出力行数, 出力_元金_返済額列) = 入出金情報全体データ_借入返済_延滞(k_延滞, 3)
                                        返済設定済み_延滞 = True
                                    End If
                                    
                                    ' 両方設定済みの場合はループを終了
                                    If 借入設定済み_延滞 And 返済設定済み_延滞 Then
                                        Exit For
                                    End If
                                End If
                            Next k_延滞
                        End If
                    End If
                    
                    ' 次のセグメントの開始日を設定
                    If j < 延滞分割日数 Then
                        延滞セグメント開始日 = 延滞分割日リスト(j + 1)
                    End If
                End If
            Next j
        End If
        
    Next i

    ' 12. 期失レコードの処理
    Dim 期失日_期失レコード As Date
    Dim 昨日 As Date
    Dim 手続開始日_期失レコード As Date
    Dim 入出金情報全体データ As Variant
    Dim 返済履歴データ As Variant
    
    期失日_期失レコード = 期失日取得(targetSheet)
    昨日 = Date - 1
    手続開始日_期失レコード = 手続開始日取得(targetSheet)
    入出金情報全体データ = 入出金情報全体取得(targetSheet)
    返済履歴データ = 返済履歴情報取得(targetSheet)
    
    ' 分割日リストの作成
    Dim 期失分割日リスト() As Date
    Dim 期失分割日数 As Long
    ReDim 期失分割日リスト(1 To 100)
    期失分割日数 = 0
    
    ' 手続開始日が期間内にあるかチェック
    If 手続開始日_期失レコード >= 期失日_期失レコード And 手続開始日_期失レコード <= 昨日 Then
        期失分割日数 = 期失分割日数 + 1
        期失分割日リスト(期失分割日数) = 手続開始日_期失レコード
    End If
    
    ' 入出金情報全体取得の日付が期間内にあるかチェック
    If IsArray(入出金情報全体データ) And UBound(入出金情報全体データ, 1) > 0 Then
        For j = 1 To UBound(入出金情報全体データ, 1)
            Dim 入出金日_期失 As Date
            入出金日_期失 = 入出金情報全体データ(j, 1)
            If 入出金日_期失 >= 期失日_期失レコード And 入出金日_期失 <= 昨日 Then
                ' 既存の分割日と重複しないかチェック
                Dim 重複フラグ_期失 As Boolean
                重複フラグ_期失 = False
                For k = 1 To 期失分割日数
                    If 期失分割日リスト(k) = 入出金日_期失 Then
                        重複フラグ_期失 = True
                        Exit For
                    End If
                Next k
                If Not 重複フラグ_期失 Then
                    期失分割日数 = 期失分割日数 + 1
                    期失分割日リスト(期失分割日数) = 入出金日_期失
                End If
            End If
        Next j
    End If
    
    ' 返済履歴の日付が期間内にあるかチェック
    If IsArray(返済履歴データ) And UBound(返済履歴データ, 1) > 0 Then
        For j = 1 To UBound(返済履歴データ, 1)
            Dim 返済履歴日_期失 As Date
            返済履歴日_期失 = 返済履歴データ(j, 1)
            If 返済履歴日_期失 >= 期失日_期失レコード And 返済履歴日_期失 <= 昨日 Then
                ' 既存の分割日と重複しないかチェック
                Dim 重複フラグ2_期失 As Boolean
                重複フラグ2_期失 = False
                For k = 1 To 期失分割日数
                    If 期失分割日リスト(k) = 返済履歴日_期失 Then
                        重複フラグ2_期失 = True
                        Exit For
                    End If
                Next k
                If Not 重複フラグ2_期失 Then
                    期失分割日数 = 期失分割日数 + 1
                    期失分割日リスト(期失分割日数) = 返済履歴日_期失
                End If
            End If
        Next j
    End If
    
    ' 分割日をソート
    If 期失分割日数 > 1 Then
        Call 分割日ソート(期失分割日リスト, 期失分割日数)
    End If
    
    ' 期失レコードの作成
    Dim 期失レコード数 As Long
    期失レコード数 = IIf(期失分割日数 = 0, 1, 期失分割日数 + 1)
    
    ' J=1時の行番号を記録
    Dim J1時の行番号_期失 As Long
    J1時の行番号_期失 = 出力行数 + 1
    
    For j = 1 To 期失レコード数
        出力行数 = 出力行数 + 1
        
        ' 通番
        出力結果(出力行数, 出力_通番列) = 出力行数
        
        ' 計算期間開始日
        If j = 1 Then
            出力結果(出力行数, 出力_計算期間開始日列) = 期失日_期失レコード
        Else
            出力結果(出力行数, 出力_計算期間開始日列) = 期失分割日リスト(j - 1)
        End If
        
        ' 計算期間終了日
        If j = 期失レコード数 Then
            出力結果(出力行数, 出力_計算期間終了日列) = 昨日
        Else
            出力結果(出力行数, 出力_計算期間終了日列) = DateAdd("d", -1, 期失分割日リスト(j))
        End If
        
        ' ステータス
        If 出力結果(出力行数, 出力_計算期間開始日列) = 手続開始日_期失レコード Then
            出力結果(出力行数, 出力_ステータス列) = 期失劣後ステータス文字列
        Else
            出力結果(出力行数, 出力_ステータス列) = 期失ステータス文字列
        End If
        
        ' イベント
        If 出力結果(出力行数, 出力_計算期間開始日列) = 手続開始日_期失レコード Then
            出力結果(出力行数, 出力_イベント列) = 破産イベント文字列
        ElseIf 出力結果(出力行数, 出力_計算期間開始日列) = 期失日_期失レコード Then
            出力結果(出力行数, 出力_イベント列) = 期失理由取得(targetSheet)
        Else
            出力結果(出力行数, 出力_イベント列) = 内入イベント文字列
        End If
        
        ' 約定返済月
        If 出力結果(出力行数, 出力_計算期間終了日列) = 昨日 Then
            出力結果(出力行数, 出力_約定返済月列) = "計算中"
        Else
            出力結果(出力行数, 出力_約定返済月列) = "ー"
        End If
        
        ' 対象元金
        Dim 対象元金_期失 As Double
        対象元金_期失 = 0
        If IsArray(入出金情報全体データ) And UBound(入出金情報全体データ, 1) > 0 Then
            Dim 計算期間開始日_期失 As Date
            計算期間開始日_期失 = 出力結果(出力行数, 出力_計算期間開始日列)
            
            ' 同じ日付のデータを探す
            Dim 同じ日付見つかった_期失 As Boolean
            同じ日付見つかった_期失 = False
            For k = 1 To UBound(入出金情報全体データ, 1)
                If 入出金情報全体データ(k, 1) = 計算期間開始日_期失 Then
                    対象元金_期失 = 入出金情報全体データ(k, 4) ' 残高
                    同じ日付見つかった_期失 = True
                    Exit For
                End If
            Next k
            
            ' 同じ日付がない場合、計算期間開始日より小さい日付の中で最大のものを探す
            If Not 同じ日付見つかった_期失 Then
                Dim 最大日付_期失 As Date
                最大日付_期失 = 日付初期値
                For k = 1 To UBound(入出金情報全体データ, 1)
                    If 入出金情報全体データ(k, 1) < 計算期間開始日_期失 And 入出金情報全体データ(k, 1) > 最大日付_期失 Then
                        最大日付_期失 = 入出金情報全体データ(k, 1)
                        対象元金_期失 = 入出金情報全体データ(k, 4) ' 残高
                    End If
                Next k
            End If
        End If
        出力結果(出力行数, 出力_対象元金列) = 対象元金_期失
        
        ' 計算日数
        出力結果(出力行数, 出力_計算日数列) = "=H" & (出力行数 + 出力開始行オフセット) & "-F" & (出力行数 + 出力開始行オフセット) & "+1"
        
        ' 区切り
        出力結果(出力行数, 出力_区切り列) = "～"
        
        ' 利率（遅延損害金利率）
        Dim 遅延損害金利率_期失 As Double
        遅延損害金利率_期失 = 0
        If IsArray(遅延損害金利率データ) And UBound(遅延損害金利率データ, 1) > 0 Then
            Dim 計算期間開始日_利率_期失 As Date
            計算期間開始日_利率_期失 = 出力結果(出力行数, 出力_計算期間開始日列)
            
            ' 同じ日付のデータを探す
            Dim 利率見つかった_期失 As Boolean
            利率見つかった_期失 = False
            For k = 1 To UBound(遅延損害金利率データ, 1)
                If 遅延損害金利率データ(k, 2) = 計算期間開始日_利率_期失 Then
                    遅延損害金利率_期失 = 遅延損害金利率データ(k, 1)
                    利率見つかった_期失 = True
                    Exit For
                End If
            Next k
            
            ' 同じ日付がない場合、計算期間開始日より小さい日付の中で最も大きい日付を探す（初回は任意の日付を受け入れ）
            If Not 利率見つかった_期失 Then
                Dim 最大日付_利率_期失 As Date
                最大日付_利率_期失 = 日付初期値 ' 初期値として最小日付を設定
                
                For k = 1 To UBound(遅延損害金利率データ, 1)
                    If 遅延損害金利率データ(k, 2) < 計算期間開始日_利率_期失 And (遅延損害金利率データ(k, 2) > 最大日付_利率_期失 Or (遅延損害金利率データ(k, 2) = 最大日付_利率_期失 And 最大日付_利率_期失 = 日付初期値)) Then
                        最大日付_利率_期失 = 遅延損害金利率データ(k, 2)
                        遅延損害金利率_期失 = 遅延損害金利率データ(k, 1)
                    End If
                Next k
            End If
        End If
        出力結果(出力行数, 出力_利率列) = 遅延損害金利率_期失
        
        ' 積数の数式設定
        出力結果(出力行数, 出力_積数列) = "=E" & (出力行数 + 出力開始行オフセット) & "*J" & (出力行数 + 出力開始行オフセット) & "*I" & (出力行数 + 出力開始行オフセット)
        
        ' 利息金額は空白
        出力結果(出力行数, 出力_利息金額列) = ""
        
        ' 遅延損害金の数式設定
        If j = 1 Then
            ' J=1の場合：=ROUNDDOWN(K行番号/365,0)
            出力結果(出力行数, 出力_遅延損害金列) = "=ROUNDDOWN(K" & (出力行数 + 出力開始行オフセット) & "/365,0)"
        Else
            ' J=1以外の場合：=ROUNDDOWN(SUM(K(J=1時の行番号):K現在の行番号)/365,0)-SUM(L(J=1時の行番号):M現在の行番号-1)
            出力結果(出力行数, 出力_遅延損害金列) = "=ROUNDDOWN(SUM(K" & (J1時の行番号_期失 + 出力開始行オフセット) & ":K" & (出力行数 + 出力開始行オフセット) & ")/365,0)-SUM(M" & (J1時の行番号_期失 + 出力開始行オフセット) & ":M" & (出力行数 + 出力開始行オフセット - 1) & ")"
        End If
        
        ' 返済日
        Dim 返済日_期失 As Variant
        返済日_期失 = ""
        Dim 計算期間開始日_返済日 As Date
        計算期間開始日_返済日 = 出力結果(出力行数, 出力_計算期間開始日列)
        
        ' 入出金情報全体取得で同じ日付があるかチェック
        If IsArray(入出金情報全体データ) And UBound(入出金情報全体データ, 1) > 0 Then
            For k = 1 To UBound(入出金情報全体データ, 1)
                If 入出金情報全体データ(k, 1) = 計算期間開始日_返済日 Then
                    返済日_期失 = 計算期間開始日_返済日
                    Exit For
                End If
            Next k
        End If
        
        ' 返済履歴で同じ日付があるかチェック
        If 返済日_期失 = "" And IsArray(返済履歴データ) And UBound(返済履歴データ, 1) > 0 Then
            For k = 1 To UBound(返済履歴データ, 1)
                If 返済履歴データ(k, 1) = 計算期間開始日_返済日 Then
                    返済日_期失 = 計算期間開始日_返済日
                    Exit For
                End If
            Next k
        End If
        出力結果(出力行数, 出力_返済日列) = 返済日_期失
        
        ' 元金_返済額
        Dim 元金返済額_期失 As Variant
        元金返済額_期失 = ""
        If IsArray(入出金情報全体データ) And UBound(入出金情報全体データ, 1) > 0 Then
            For k = 1 To UBound(入出金情報全体データ, 1)
                If 入出金情報全体データ(k, 1) = 計算期間開始日_返済日 Then
                    元金返済額_期失 = 入出金情報全体データ(k, 3) ' 入出金金額
                    Exit For
                End If
            Next k
        End If
        出力結果(出力行数, 出力_元金_返済額列) = 元金返済額_期失
        
        ' 利息_返済額
        Dim 利息返済額_期失 As Variant
        利息返済額_期失 = ""
        If IsArray(返済履歴データ) And UBound(返済履歴データ, 1) > 0 Then
            For k = 1 To UBound(返済履歴データ, 1)
                If 返済履歴データ(k, 1) = 計算期間開始日_返済日 And InStr(返済履歴データ(k, 2), 利息摘要文字列) > 0 Then
                    利息返済額_期失 = 返済履歴データ(k, 3) ' 出金金額
                    Exit For
                End If
            Next k
        End If
        出力結果(出力行数, 出力_利息_返済額列) = 利息返済額_期失
        
        ' 遅損金_返済額
        Dim 遅損金返済額_期失 As Variant
        遅損金返済額_期失 = ""
        If IsArray(返済履歴データ) And UBound(返済履歴データ, 1) > 0 Then
            For k = 1 To UBound(返済履歴データ, 1)
                If 返済履歴データ(k, 1) = 計算期間開始日_返済日 And InStr(返済履歴データ(k, 2), 遅延損害金摘要文字列) > 0 Then
                    遅損金返済額_期失 = 返済履歴データ(k, 3) ' 出金金額
                    Exit For
                End If
            Next k
        End If
        出力結果(出力行数, 出力_遅損金_返済額列) = 遅損金返済額_期失
        
    Next j
    
    ' 期失レコードの「計」行を追加
    出力行数 = 出力行数 + 1
    
    ' 通番（設定しない）
    ' 出力結果(出力行数, 出力_通番列) = 出力行数
    
    ' 約定返済月
    出力結果(出力行数, 出力_約定返済月列) = "計"
    
    ' 対象元金（前レコードの値）
    If 出力行数 > 1 Then
        出力結果(出力行数, 出力_対象元金列) = 出力結果(出力行数 - 1, 出力_対象元金列)
    End If
    
    ' 利息金額の数式設定（レコード全体の利息金額の合計-レコード全体の利息_返済額の合計）
    出力結果(出力行数, 出力_利息金額列) = "=SUM(L" & (1 + 出力開始行オフセット) & ":L" & (出力行数 + 出力開始行オフセット - 1) & ")-SUM(R" & (1 + 出力開始行オフセット) & ":R" & (出力行数 + 出力開始行オフセット - 1) & ")"
    
    ' 遅延損害金の数式設定（レコード全体の遅延損害金の合計-レコード全体の遅損金_返済額の合計）
    出力結果(出力行数, 出力_遅延損害金列) = "=SUM(M" & (1 + 出力開始行オフセット) & ":M" & (出力行数 + 出力開始行オフセット - 1) & ")-SUM(S" & (1 + 出力開始行オフセット) & ":S" & (出力行数 + 出力開始行オフセット - 1) & ")"
    
    ' 結果配列のサイズを調整
    If 出力行数 > 0 Then
        ' 新しい配列を作成して必要な部分をコピー
        Dim 最終結果() As Variant
        ReDim 最終結果(1 To 出力行数, 1 To 19)
        
        Dim copyRow As Long, copyCol As Long
        For copyRow = 1 To 出力行数
            For copyCol = 1 To 19
                最終結果(copyRow, copyCol) = 出力結果(copyRow, copyCol)
            Next copyCol
        Next copyRow
        
        出力データ作成 = 最終結果
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



' 入出金情報全体取得関数（返済分も含む）
' 指定されたシートの入出金開始行から全ての入出金情報を取得（返済分の除外なし）
Public Function 入出金情報全体取得(targetSheet As Worksheet) As Variant
    Dim startRow As Long
    Dim currentRow As Long
    Dim dataArray() As Variant
    Dim rowCount As Long
    Dim i As Long, j As Long
    
    startRow = 入出金開始行 ' 開始行
    currentRow = startRow
    rowCount = 0
    
    ' データ行数をカウント（B列が空白になるまで、全ての行をカウント）
    Do While targetSheet.Cells(currentRow, 入出金日列).Value <> ""
        rowCount = rowCount + 1
        currentRow = currentRow + 1
    Loop
    
    ' データが存在しない場合は空の配列を返す
    If rowCount = 0 Then
        入出金情報全体取得 = Array()
        Exit Function
    End If
    
    ' 配列を初期化（行数 x 5列）
    ReDim dataArray(1 To rowCount, 1 To 5)
    
    ' データを取得してバリデーション
    currentRow = startRow
    
    For i = 1 To rowCount
        ' B列：入出金日（日付チェック）
        Dim dateValue As Variant
        dateValue = targetSheet.Cells(currentRow, 入出金日列).Value
        If Not IsDate(dateValue) Then
            Err.Raise 13, "入出金情報全体取得", currentRow & "行目のB列（入出金日）が日付ではありません。"
        End If
        dataArray(i, 1) = CDate(dateValue)
        
        ' C列：摘要（文字列、チェック不要）
        dataArray(i, 2) = CStr(targetSheet.Cells(currentRow, 摘要列).Value)
        
        ' 同一日付データ処理：現在が返済分で同一日付の非返済分データが存在する場合はエラー
        Dim currentDate As Date
        currentDate = dataArray(i, 1)
        Dim currentRemark As String
        currentRemark = dataArray(i, 2)
        Dim isCurrentRepayment As Boolean
        isCurrentRepayment = (Len(currentRemark) >= Len(返済摘要返済分文字列) And Right(currentRemark, Len(返済摘要返済分文字列)) = 返済摘要返済分文字列)
        
        ' 以前のデータに同一日付があるかチェック
        dim checkIndex as long
        For checkIndex = 1 To i - 1
            If dataArray(checkIndex, 1) = currentDate Then
                Dim previousRemark As String
                previousRemark = CStr(dataArray(checkIndex, 2))
                Dim isPreviousRepayment As Boolean
                isPreviousRepayment = (Len(previousRemark) >= Len(返済摘要返済分文字列) And Right(previousRemark, Len(返済摘要返済分文字列)) = 返済摘要返済分文字列)
                
                ' 両方とも返済分でない場合はエラー
                If Not isCurrentRepayment And Not isPreviousRepayment Then
                    Err.Raise 13, "入出金情報全体取得", "同一日付（" & Format(currentDate, "yyyy/mm/dd") & "）に複数の非返済分データが存在します。データを確認してください。"
                End If
                
                ' 現在が返済分で以前が非返済分の場合、現在のレコードをスキップ
                If isCurrentRepayment And Not isPreviousRepayment Then
                    GoTo NextRecord
                End If
                
                ' 以前が返済分で現在が非返済分の場合、以前のレコードを削除（無効としてマーク）
                If Not isCurrentRepayment And isPreviousRepayment Then
                    ' 以前のレコードを無効としてマーク（日付を最小値に設定）
                    dataArray(checkIndex, 1) = 日付初期値
                End If
            End If
        Next checkIndex
        
        ' D列：入出金金額（数値チェック）
        Dim amountValue As Variant
        amountValue = targetSheet.Cells(currentRow, 入出金金額列).Value
        If Not IsNumeric(amountValue) Then
            Err.Raise 13, "入出金情報全体取得", currentRow & "行目のD列（入出金金額）が数値ではありません。"
        End If
        dataArray(i, 3) = CDbl(amountValue)
        
        ' E列：残高（数値チェック）
        Dim balanceValue As Variant
        balanceValue = targetSheet.Cells(currentRow, 残高列).Value
        If Not IsNumeric(balanceValue) Then
            Err.Raise 13, "入出金情報全体取得", currentRow & "行目のE列（残高）が数値ではありません。"
        End If
        dataArray(i, 4) = CDbl(balanceValue)
        
        ' F列：延滞中の約定返済元金合計（入力があれば数値チェック）
        Dim principalValue As Variant
        principalValue = targetSheet.Cells(currentRow, 約定返済元金列).Value
        If principalValue <> "" Then
            If Not IsNumeric(principalValue) Then
                Err.Raise 13, "入出金情報全体取得", currentRow & "行目のF列（延滞中の約定返済元金合計）が数値ではありません。"
            End If
            dataArray(i, 5) = CDbl(principalValue)
        Else
            ' 空白の場合は前の行の値を使用、ただしi=1の場合は0を設定
            If i = 1 Then
                dataArray(i, 5) = 0
            Else
                dataArray(i, 5) = dataArray(i - 1, 5)
            End If
        End If
        
        ' dataArray(i, 5)算出完了後の追加調整：摘要が「返済分」で終わる場合は入出金金額を減らす
        If Len(dataArray(i, 2)) >= Len(返済摘要返済分文字列) And Right(dataArray(i, 2), Len(返済摘要返済分文字列)) = 返済摘要返済分文字列 Then
            dataArray(i, 5) = dataArray(i, 5) - dataArray(i, 3)
        End If
        
NextRecord:
        currentRow = currentRow + 1
    Next i
    
    ' 無効としてマークされたレコードをフィルタリング（日付が初期値のレコード）
    Dim validCount As Long
    validCount = 0
    For i = 1 To rowCount
        If dataArray(i, 1) <> 日付初期値 Then
            validCount = validCount + 1
        End If
    Next i
    
    ' 有効なレコードがない場合は空配列を返す
    If validCount = 0 Then
        入出金情報全体取得 = Array()
        Exit Function
    End If
    
    ' フィルタリング後の配列を作成
    Dim filteredArray() As Variant
    ReDim filteredArray(1 To validCount, 1 To 5)
    Dim validIndex As Long
    validIndex = 0
    
    For i = 1 To rowCount
        If dataArray(i, 1) <> 日付初期値 Then
            validIndex = validIndex + 1
            For j = 1 To 5
                filteredArray(validIndex, j) = dataArray(i, j)
            Next j
        End If
    Next i
    
    入出金情報全体取得 = filteredArray
End Function

' 期失日を取得する関数
' 指定されたシートのC列25行目のセル値を返す
Public Function 期失日取得(targetSheet As Worksheet) As Date
    Dim cellValue As Variant
    
    ' セル値を取得
    cellValue = targetSheet.Cells(期失日行, 期失日列).Value
    
    ' 日付型かチェック
    If Not IsDate(cellValue) Then
        Err.Raise 13, "期失日", "セル値が日付型ではありません。"
    End If
    
    ' 日付型に変換して返す
    期失日取得 = CDate(cellValue)
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
                結果(i, 2) = 日付初期値
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
                結果(i, 2) = 日付初期値
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

' 返済履歴情報取得関数
' 70行目からB列が空白になるまで、B列（日付）、C列（摘要）、D列（出金金額）のデータを取得
Public Function 返済履歴情報取得(targetSheet As Worksheet) As Variant
    Dim 現在行 As Long
    Dim 結果() As Variant
    Dim 行数 As Long
    Dim i As Long
    
    ' データ行数をカウント
    現在行 = 返済履歴開始行
    行数 = 0
    
    Do While targetSheet.Cells(現在行, 返済履歴日付列).Value <> ""
        行数 = 行数 + 1
        現在行 = 現在行 + 1
    Loop
    
    ' データが存在しない場合は空の配列を返す
    If 行数 = 0 Then
        返済履歴情報取得 = Array()
        Exit Function
    End If
    
    ' 結果配列を初期化（行数 x 3列）
    ReDim 結果(1 To 行数, 1 To 3)
    
    ' データを取得
    現在行 = 返済履歴開始行
    For i = 1 To 行数
        ' B列（日付）を取得
        Dim 日付値 As Variant
        日付値 = targetSheet.Cells(現在行, 返済履歴日付列).Value
        If Not IsDate(日付値) Then
            Err.Raise 13, "返済履歴情報取得", 現在行 & "行目のB列（日付）が日付ではありません。"
        End If
        結果(i, 1) = CDate(日付値)
        
        ' C列（摘要）を取得
        結果(i, 2) = CStr(targetSheet.Cells(現在行, 返済履歴摘要列).Value)
        
        ' D列（出金金額）を取得
        Dim 出金金額値 As Variant
        出金金額値 = targetSheet.Cells(現在行, 返済履歴出金金額列).Value
        If Not IsNumeric(出金金額値) Then
            Err.Raise 13, "返済履歴情報取得", 現在行 & "行目のD列（出金金額）が数値ではありません。"
        End If
        結果(i, 3) = CDbl(出金金額値)
        
        現在行 = 現在行 + 1
    Next i
    
    返済履歴情報取得 = 結果
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
    入出金データ = 入出金情報全体取得(targetSheet)
    
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
        If 未返済前月 < 最初日 and 未返済前月 > 日付初期値 Then
            ' 最初日を返す
            計算期間最初日取得 = 最初日
            Exit Function
        End If
        
        ' 未返済前月が日付初期値の場合は、入出金最小日付を設定
        If 未返済前月 = 日付初期値 Then
            計算期間最初日取得 = 入出金最小日付
        Else
            ' 上記以外の場合は、入出金情報の最小日付と「未返済前月」を比較して、小さいほうを返す
            If 入出金最小日付 < 未返済前月 Then
                計算期間最初日取得 = 入出金最小日付
            Else
                計算期間最初日取得 = 未返済前月
            End If
        End If
    Else
        ' 入出金データが存在しない場合
        ' 「未返済前月」がこの最初日より小さいかをチェック
        If 未返済前月 < 最初日 and 未返済前月 > 日付初期値 Then
            ' 最初日を返す
            計算期間最初日取得 = 最初日
        Else
            ' 未返済前月を返す
            計算期間最初日取得 = 未返済前月
        End If
    End If
End Function







