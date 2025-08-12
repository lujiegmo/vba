Option Explicit

' 定数定義

' 入出金情報関連定数
Public Const 入出金開始行 As Long = 51 ' 入出金データ開始行
Public Const 入出金日列 As Long = 2     ' B列：入出金日
Public Const 摘要列 As Long = 3         ' C列：摘要
Public Const 入出金金額列 As Long = 4   ' D列：入出金金額
Public Const 残高列 As Long = 5         ' E列：残高
Public Const 約定返済元金列 As Long = 6 ' F列：延滞中の約定返済元金合計
Public Const 入出金情報最大行数 As Long = 20  ' 入出金情報の最大設定可能行数

' 期失日関連定数
Public Const 期失日列 As Long = 3  ' C列
Public Const 期失日行 As Long = 30  ' 30行目

' 借入利率関連定数
Public Const 借入利率列 As Long = 2  ' B列：借入利率
Public Const 借入利率開始日列 As Long = 3  ' C列：開始日
Public Const 借入利率開始行 As Long = 34  ' 24行目
Public Const 借入利率最大行数 As Long = 2  ' 借入利率の最大設定可能行数

' 遅延損害金利率関連定数
Public Const 遅延損害金利率列 As Long = 2  ' B列：遅延損害金利率
Public Const 遅延損害金利率開始日列 As Long = 3  ' C列：開始日
Public Const 遅延損害金利率開始行 As Long = 15  ' 15行目
Public Const 遅延損害金利率最大行数 As Long = 3  ' 最大3行まで設定可能

' 計算書作成パス関連定数
Public Const 計算書作成パス列 As Long = 3  ' C列：計算書作成パス
Public Const 計算書作成パス行 As Long = 7  ' 7行目

' 顧客番号関連定数
Public Const 顧客番号列 As Long = 3  ' C列：顧客番号
Public Const 顧客番号行 As Long = 6  ' 6行目

' 手続理由関連定数
Public Const 手続理由列 As Long = 3  ' C列：手続理由
Public Const 手続理由行 As Long = 10  ' 10行目

' 手続開始日関連定数
Public Const 手続開始日列 As Long = 3  ' C列：手続開始日
Public Const 手続開始日行 As Long = 11  ' 11行目

' ローン口座ステータス関連定数
Public Const ローン口座ステータス列 As Long = 3  ' C列：ローン口座ステータス
Public Const ローン口座ステータス行 As Long = 27  ' 27行目

' 期失理由関連定数
Public Const 期失理由列 As Long = 5  ' E列：期失理由
Public Const 期失理由行 As Long = 30  ' 30行目

' 初回借入日関連定数
Public Const 初回借入日列 As Long = 3  ' C列：初回借入日
Public Const 初回借入日行 As Long = 22  ' 22行目

' 契約期限日関連定数
Public Const 契約期限日列 As Long = 3  ' C列：契約期限日
Public Const 契約期限日行 As Long = 23  ' 23行目

' 借入限度額関連定数
Public Const 借入限度額列 As Long = 3  ' C列：借入限度額
Public Const 借入限度額行 As Long = 24  ' 24行目

' 摘要文字列関連定数
Public Const 借入摘要借入文字列 As String = "借入"     ' 借入を示す摘要文字列
Public Const 借入摘要借換文字列 As String = "借換"     ' 借換を示す摘要文字列
Public Const 返済摘要返済分文字列 As String = "返済分"   ' 返済を示す摘要文字列

' ローン口座ステータス関連定数
Public Const 期失ステータス文字列 As String = "期失"  ' 期失を示すステータス文字列
Public Const 期限切れ理由文字列 As String = "期限切れ"  ' 期限切れを示す理由文字列
Public Const 正常ステータス文字列 As String = "正常"  ' 正常を示すステータス文字列
Public Const 約定返済イベント文字列 As String = "約定返済"  ' 約定返済を示すイベント文字列
Public Const 期失劣後ステータス文字列 As String = "期失（劣後）"  ' 期失（劣後）を示すステータス文字列
Public Const 期限切れ劣後ステータス文字列 As String = "期限切れ（劣後）"  ' 期限切れ（劣後）を示すステータス文字列
Public Const 延滞イベント文字列 As String = "延滞"  ' 延滞を示すイベント文字列
Public Const 内入イベント文字列 As String = "内入"  ' 内入を示すイベント文字列
Public Const 破産イベント文字列 As String = "破産"  ' 破産イベント文字列
Public Const 利息摘要文字列 As String = "利息"  ' 利息を示す摘要文字列
Public Const 遅延損害金摘要文字列 As String = "遅延損害金"  ' 遅延損害金を示す摘要文字列

' 日付関連定数
Public Const 日付初期値 As Date = #1/1/1900#  ' 空白日付の初期値

' ワークシート名関連定数
Public Const ツールシート名 As String = "ツール"  ' ツールシートの名前
Public Const テンプレートシート名 As String = "テンプレート_EXCEL"  ' テンプレートシートの名前
Public Const テンプレートワードシート名 As String = "テンプレート_WORD"  ' テンプレートWORDシートの名前

' 出力関連定数
Public Const 出力タイトル行 As Long = 7  ' 出力タイトル行
Public Const 出力開始行オフセット As Long = 8  ' A9セルから貼り付けるためのオフセット
Public Const 出力顧客番号行 As Long = 4  ' B4行：顧客番号
Public Const 出力顧客番号列 As Long = 2  ' B列：顧客番号
Public Const 出力手続開始日行 As Long = 2  ' J2行：手続開始日
Public Const 出力手続開始日列 As Long = 10  ' J列：手続開始日
Public Const 出力期失日行 As Long = 3  ' J3行：期失日
Public Const 出力期失日列 As Long = 10  ' J列：期失日
Public Const 出力期失理由行 As Long = 3  ' K3行：期失理由
Public Const 出力期失理由列 As Long = 11  ' K列：期失理由

' 出力項目列定数（A列～S列）
Public Const 出力_通番列 As Long = 1          ' A列：通番
Public Const 出力_ステータス列 As Long = 2      ' B列：ステータス
Public Const 出力_イベント列 As Long = 3        ' C列：イベント
Public Const 出力_約定返済月列 As Long = 4      ' D列：約定返済月
Public Const 出力_対象元金列 As Long = 5        ' E列：対象元金
Public Const 出力_計算期間開始日列 As Long = 6  ' F列：計算期間開始日
Public Const 出力_区切り列 As Long = 7          ' G列："～"
Public Const 出力_計算期間終了日列 As Long = 8  ' H列：計算期間終了日
Public Const 出力_計算日数列 As Long = 9        ' I列：計算日数
Public Const 出力_利率列 As Long = 10           ' J列：利率
Public Const 出力_積数列 As Long = 11           ' K列：積数
Public Const 出力_利息金額列 As Long = 12       ' L列：利息金額
Public Const 出力_遅延損害金列 As Long = 13     ' M列：遅延損害金
Public Const 出力_借入日列 As Long = 14         ' N列：借入日
Public Const 出力_借入額列 As Long = 15         ' O列：借入額
Public Const 出力_返済日列 As Long = 16         ' P列：返済日
Public Const 出力_元金_返済額列 As Long = 17    ' Q列：元金_返済額
Public Const 出力_利息_返済額列 As Long = 18    ' R列：利息_返済額
Public Const 出力_遅損金_返済額列 As Long = 19  ' S列：遅損金_返済額

' 返済予定情報の定数
Public Const 返済予定開始行 As Long = 40  ' 40行目
Public Const 返済予定日列 As Long = 3  ' C列
Public Const 返済元金列 As Long = 4    ' D列

' 返済履歴情報の定数
Public Const 返済履歴開始行 As Long = 75   ' 75行目
Public Const 返済履歴日付列 As Long = 2    ' B列：日付
Public Const 返済履歴摘要列 As Long = 3    ' C列：摘要
Public Const 返済履歴出金金額列 As Long = 4 ' D列：出金金額
Public Const 返済履歴情報最大行数 As Long = 20  ' 返済履歴情報の最大設定可能行数

' 削除最後行目の定数
Public Const 削除最後行目 As Long = 69

' データ貼り付け開始行の定数
Public Const データ貼り付け開始行 As Long = 出力開始行オフセット + 1

' テンプレート_WORDファイル保存関数
' 引数: 保存先フォルダパス
' 戻り値: 保存されたファイルの完全パス
Public Function テンプレートワードファイル保存(保存先フォルダパス As String, toolSheet As Worksheet) As String
    On Error GoTo ErrorHandler
    
    Dim ワークブック As Workbook
    Dim ワークシート As Worksheet
    Dim 埋め込みオブジェクト As OLEObject
    Dim 基本ファイル名 As String
    Dim 拡張子 As String
    Dim 完全ファイルパス As String
    Dim 保存ファイルパス As String
    Dim カウンタ As Long
    
    ' 現在のワークブックとテンプレート_WORDシートを取得
    Set ワークブック = ThisWorkbook
    Set ワークシート = ワークブック.Worksheets(テンプレートワードシート名)
    
    ' 埋め込みWordファイルを検索
    Set 埋め込みオブジェクト = Nothing
    Dim obj As OLEObject
    For Each obj In ワークシート.OLEObjects
        If obj.progID Like "Word.*" Then
            Set 埋め込みオブジェクト = obj
            Exit For
        End If
    Next obj
    
    ' 埋め込みオブジェクトが見つからない場合はエラー
    If 埋め込みオブジェクト Is Nothing Then
        Err.Raise 1001, "テンプレートワードファイル保存", "テンプレート_WORDシートに埋め込みWordファイルが見つかりません。"
    End If
    
    ' 保存先フォルダパスの末尾にバックスラッシュを追加（必要に応じて）
    If Right(保存先フォルダパス, 1) <> "\" Then
        保存先フォルダパス = 保存先フォルダパス & "\"
    End If
    
    ' 基本ファイル名と拡張子を設定
    Dim 顧客番号 As String
    顧客番号 = 顧客番号取得(toolSheet)
    基本ファイル名 = "利息計算書" & 顧客番号
    拡張子 = ".docx"
    
    ' 完全ファイルパスを作成
    完全ファイルパス = 保存先フォルダパス & 基本ファイル名 & 拡張子
    保存ファイルパス = 完全ファイルパス
    
    ' 連番カウンタを初期化
    カウンタ = 1
    
    ' 既存ファイルがある場合は連番を付けて新しいファイル名を作成
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Do While fso.FileExists(保存ファイルパス)
        Dim ファイル名部分 As String
        Dim 拡張子部分 As String
        Dim フォルダパス部分 As String

        ' ファイルパスを分解
        フォルダパス部分 = Left(完全ファイルパス, InStrRev(完全ファイルパス, "\"))
        ファイル名部分 = Mid(完全ファイルパス, InStrRev(完全ファイルパス, "\") + 1)
        拡張子部分 = Right(ファイル名部分, 5) ' ".docx"
        ファイル名部分 = Left(ファイル名部分, Len(ファイル名部分) - 5)

        ' 連番付きファイル名を作成
        保存ファイルパス = フォルダパス部分 & ファイル名部分 & "(" & カウンタ & ")" & 拡張子部分
        カウンタ = カウンタ + 1
    Loop
    
    ' 埋め込みWordファイルを保存
    埋め込みオブジェクト.Verb xlVerbOpen ' Wordファイルを開く
    
    ' Wordアプリケーションオブジェクトを取得してファイルを保存
    Dim wordApp As Object
    Dim wordDoc As Object
    Set wordApp = GetObject(, "Word.Application")
    Set wordDoc = 埋め込みオブジェクト.Object
    
    ' ファイルを指定パスに保存
    wordDoc.SaveAs2 保存ファイルパス
    wordDoc.Close
    
    ' 保存されたWordファイルを開く
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(保存ファイルパス)
    wordApp.ScreenUpdating = False
    
    ' Wordファイルの内容を一括置換
    ワード全置換実行 wordDoc, toolSheet

    wordApp.ScreenUpdating = True
    
    ' ファイルを指定パスに保存
    wordDoc.SaveAs2 保存ファイルパス
    wordDoc.Close
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
   
    ' 戻り値として保存されたファイルパスを返す
    テンプレートワードファイル保存 = 保存ファイルパス
    
    Exit Function
    
ErrorHandler:
    ' エラーハンドリング
    Dim エラーメッセージ As String
    エラーメッセージ = "テンプレートWordファイルの保存中にエラーが発生しました。" & vbCrLf & _
                    "エラー番号: " & Err.Number & vbCrLf & _
                    "エラー内容: " & Err.Description
    
    ' Wordアプリケーションが開いている場合はクリーンアップ
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    On Error GoTo 0
    
    Err.Raise Err.Number, "テンプレートワードファイル保存", エラーメッセージ
End Function

' Word全置換実行関数
' 引数: Wordドキュメント
' 戻り値: なし（利息計算書に必要な全ての置換を実行）
Public Sub ワード全置換実行(wordDoc As Object, toolSheet As Worksheet)
    On Error GoTo ErrorHandler

    ' テンプレートワードシートから直接データを取得
    Dim ワークブック As Workbook
    Dim ワークシート As Worksheet
    Set ワークブック = toolSheet.Parent
    Set ワークシート = ワークブック.Worksheets(テンプレートワードシート名)
    
    ' 出力開始行+1行から最終行までのデータを取得
    Dim 最終行 As Long
    最終行 = ワークシート.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    
    Dim 出力データ As Variant
    If 最終行 >= 出力開始行オフセット + 1 Then
        出力データ = ワークシート.Range(ワークシート.Cells(出力開始行オフセット + 1, 1), ワークシート.Cells(最終行, ワークシート.UsedRange.Column + ワークシート.UsedRange.Columns.Count - 1)).Value
    Else
        ' データが存在しない場合
        Exit Sub
    End If
    
    ' 出力データの最終行から各計算値を取得
    Dim 貸付金_計 As Variant
    Dim 約定利息金_計 As Variant
    Dim 遅延損害金_計 As Variant
    
    If IsArray(出力データ) And UBound(出力データ, 1) >= 1 Then
        貸付金_計 = 出力データ(UBound(出力データ, 1), 出力_対象元金列) ' 最終行の対象元金列（貸付金）
        
        ' 数式の場合は算出値を取得
        Dim 利息金額セル As Range
        Set 利息金額セル = toolSheet.Cells(UBound(出力データ, 1) + 出力開始行オフセット, 出力_利息金額列) ' 出力開始行オフセット分
        If 利息金額セル.HasFormula Then
            約定利息金_計 = 利息金額セル.Value ' 数式の算出値を取得
        Else
            約定利息金_計 = 出力データ(UBound(出力データ, 1), 出力_利息金額列) ' 最終行の利息金額列（約定利息金）
        End If
        
        遅延損害金_計 = 出力データ(UBound(出力データ, 1), 出力_遅延損害金列) ' 最終行の遅延損害金列（遅延損害金）
    End If
    
    ' 利息計算書用の置換データを定義
    ' 必要に応じて置換ペアを追加・修正してください
    Dim 置換ペア配列 As Variant
    置換ペア配列 = Array( _
        "{{顧客番号}}", 顧客番号取得(toolSheet), _
        "{{手続理由}}", 手続理由取得(toolSheet), _
        "{{手続開始日}}", Format(手続開始日取得(toolSheet), "yyyy年mm月dd日"), _
        "{{初回借入日}}", Format(初回借入日取得(toolSheet), "yyyy年mm月dd日"), _
        "{{契約期限日}}", Format(契約期限日取得(toolSheet), "yyyy年mm月dd日"), _
        "{{作成日}}", Format(Now, "yyyy年mm月dd日"), _
        "{{借入限度額}}", Format(借入限度額取得(toolSheet), "#,##0"), _
        "{{ステータス}}", ローン口座ステータス取得(toolSheet), _
        "{{期失日}}", Format(期失日取得(toolSheet), "yyyy年mm月dd日"), _
        "{{期失理由}}", 期失理由取得(toolSheet), _
        "{{貸付金_計}}", Format(貸付金_計, "#,##0"), _
        "{{約定利息金_計}}", Format(約定利息金_計, "#,##0"), _
        "{{遅延損害金_計}}", Format(遅延損害金_計, "#,##0"), _
        "{{利息内容}}", 利息内容生成(toolSheet), _
        "{{遅延損害金内容}}", 遅延損害金内容生成(toolSheet) _
    )
    
    ' 配列の各ペアで置換を実行
    Dim i As Long
    For i = LBound(置換ペア配列) To UBound(置換ペア配列) Step 2
        If i + 1 <= UBound(置換ペア配列) Then
            Call ワード文字列置換(wordDoc, CStr(置換ペア配列(i)), CStr(置換ペア配列(i + 1)))
        End If
    Next i

    Call 利息明細生成(wordDoc, toolSheet, 出力データ)
    Call 遅延損害金明細生成(wordDoc, toolSheet, 出力データ)

    Exit Sub
    
ErrorHandler:
    ' エラーが発生した場合は無視して続行
    Resume Next
End Sub

' 出力データから利息明細文字列を生成
Private Function 利息明細生成(wordDoc As Object, toolSheet As Worksheet, 出力データ As Variant) As String
    Dim 利息明細内容 As String
    Dim 利息明細補足 As String
    Dim i As Long
    Dim 順番 As Long
    Dim 最大順番 As Long
    Dim 表示用順番 As Long
    Dim 順番文字 As String
    
    ' データが配列かつ要素が存在するかチェック
    If Not IsArray(出力データ) Or UBound(出力データ, 1) < 1 Then
        利息明細生成 = ""
        Exit Function
    End If
    
    利息明細内容 = ""
    順番 = 0
    最大順番 = 0
    表示用順番 = 0
    
    ' 順番数字を格納する配列を定義
    Dim 順番配列() As Long
    dim 内入れ順番配列() As Long
    
    ' 最大順番を取得（利息金額列と利息_返済額列にデータがあるレコードの数）
    For i = 1 To UBound(出力データ, 1)
        If (出力データ(i, 出力_利息金額列) <> "" And 出力データ(i, 出力_利息金額列) <> 0) Or _
           (出力データ(i, 出力_利息_返済額列) <> "" And 出力データ(i, 出力_利息_返済額列) <> 0) Then
            最大順番 = 最大順番 + 1
        End If
    Next i
    
    ' 順番配列のサイズを設定
    If 最大順番 > 0 Then
        ReDim 順番配列(1 To 最大順番)
        ReDim 内入れ順番配列(1 To 最大順番)
    End If
    
    ' 出力データをループして利息明細を生成
    For i = 1 To UBound(出力データ, 1) - 1 ' 最後の行は合計行なので除外
        
        ' 利息返済額列にデータがある場合
        If 出力データ(i, 出力_利息_返済額列) <> "" And 出力データ(i, 出力_利息_返済額列) <> 0 Then ' R列：利息_返済額
            表示用順番 = 表示用順番 + 1
            内入れ順番配列(順番) = 表示用順番
            
            ' 順番を①②③形式に変換（共通関数を使用）
            順番文字 = 順番文字変換(表示用順番)
            
            利息明細内容 = 順番文字 & "▲" & Format(出力データ(i, 出力_利息_返済額列), "#,##0") & "円" & "^p" & _
                        Format(出力データ(i, 出力_返済日列), "yyyy年mm月dd日") & "に上記" & 順番配列範囲変換(順番配列) & "利息の一部として内入返済" & "^p" & "{{利息明細" & CStr(表示用順番 + 1) & "}}"
            Call ワード文字列置換(wordDoc, "{{利息明細" & 表示用順番 & "}}", 利息明細内容)
        End If

        ' 利息金額列にデータがある場合
        If 出力データ(i, 出力_利息金額列) <> "" And 出力データ(i, 出力_利息金額列) <> 0 Then ' L列：利息金額
            順番 = 順番 + 1
            表示用順番 = 表示用順番 + 1
            順番配列(順番) = 表示用順番
            
            ' 順番を①②③形式に変換（共通関数を使用）
            順番文字 = 順番文字変換(表示用順番)
            
            利息明細内容 = 順番文字 & Format(出力データ(i, 出力_利息金額列), "#,##0") & "円" & "^p" & _
                        "貸付金" & Format(出力データ(i, 出力_対象元金列), "#,##0") & "円に対する" & _
                        Format(出力データ(i, 出力_計算期間開始日列), "yyyy年mm月dd日") & "から" & _
                        Format(出力データ(i, 出力_計算期間終了日列), "yyyy年mm月dd日") & "まで" & _
                        出力データ(i, 出力_計算日数列) & "日間、年" & Format(出力データ(i, 出力_利率列) * 100, "0.0") & "%の割合による利息" & "^p" & "{{利息明細" & CStr(表示用順番 + 1) & "}}"
            Call ワード文字列置換(wordDoc, "{{利息明細" & 表示用順番 & "}}", 利息明細内容)
        End If
    Next i
    
    Call ワード文字列置換(wordDoc, "{{利息明細" & CStr(表示用順番 + 1) & "}}", "")

    Dim 内入れ範囲文字列 As String
    内入れ範囲文字列 = 順番配列範囲変換(内入れ順番配列)
    
    If 内入れ範囲文字列 = "" Then
        利息明細補足 = 順番配列範囲変換(順番配列) & "の合計金額"
    Else
        利息明細補足 = 順番配列範囲変換(順番配列) & "の合計金額より内入" & 内入れ範囲文字列 & "を控除"
    End If
    Call ワード文字列置換(wordDoc, "{{利息明細補足}}", 利息明細補足)

    利息明細生成 = ""
End Function

' 利息内容生成関数
' 借入利率取得データのレコード数に応じて利息内容文字列を生成
Private Function 利息内容生成(toolSheet As Worksheet) As String
    Dim 借入利率データ As Variant
    Dim 利息内容 As String
    
    ' 借入利率データを取得
    借入利率データ = 借入利率取得(toolSheet)
    
    ' データが配列かつ要素が存在するかチェック
    If IsArray(借入利率データ) And UBound(借入利率データ, 1) >= 1 Then
        If UBound(借入利率データ, 1) = 1 Then
            ' 単一レコードの場合
            利息内容 = "利息年" & Format(借入利率データ(1, 1) * 100, "0.0") & "%"
        Else
            ' 複数レコードの場合
            利息内容 = "^p" & "利息年" & Format(借入利率データ(UBound(借入利率データ, 1), 1) * 100, "0.0") & "%（" & _
                      Format(借入利率データ(2, 2), "yyyy年mm月dd日") & "以前" & _
                      Format(借入利率データ(1, 1) * 100, "0.0") & "%）"
        End If
    Else
        ' データが存在しない場合のデフォルト値
        利息内容 = ""
    End If
    
    利息内容生成 = 利息内容
End Function

' 遅延損害金内容生成関数
' 遅延損害金利率取得データのレコード数に応じて遅延損害金内容文字列を生成
Private Function 遅延損害金内容生成(toolSheet As Worksheet) As String
    Dim 遅延損害金利率データ As Variant
    Dim 遅延損害金内容 As String
    
    ' 遅延損害金利率データを取得
    遅延損害金利率データ = 遅延損害金利率取得(toolSheet)
    
    ' データが配列かつ要素が存在するかチェック
    If IsArray(遅延損害金利率データ) And UBound(遅延損害金利率データ, 1) >= 1 Then
        If UBound(遅延損害金利率データ, 1) = 1 Then
            ' 単一レコードの場合
            遅延損害金内容 = "遅延損害金年" & Format(遅延損害金利率データ(1, 1) * 100, "0.0") & "%"

        ElseIf UBound(遅延損害金利率データ, 1) = 2 Then
            ' 2レコードの場合
            遅延損害金内容 = "^p" & "遅延損害金年" & Format(遅延損害金利率データ(UBound(遅延損害金利率データ, 1), 1) * 100, "0.0") & "%（" & _
                          Format(遅延損害金利率データ(2, 2), "yyyy年mm月dd日") & "以前" & _
                          Format(遅延損害金利率データ(1, 1) * 100, "0.0") & "%）" & "^p"
        ElseIf UBound(遅延損害金利率データ, 1) = 3 Then
            ' 3レコードの場合
            遅延損害金内容 = "^p" & "遅延損害金年" & Format(遅延損害金利率データ(UBound(遅延損害金利率データ, 1), 1) * 100, "0.0") & "%（" & _
                          Format(遅延損害金利率データ(2, 2), "yyyy年mm月dd日") & "以前" & _
                          Format(遅延損害金利率データ(1, 1) * 100, "0.0") & "%、" & _
                          Format(DateAdd("d", 1, 遅延損害金利率データ(2, 2)), "yyyy年mm月dd日") & "～" & _
                          Format(遅延損害金利率データ(3, 2), "yyyy年mm月dd日") & _
                          Format(遅延損害金利率データ(2, 1) * 100, "0.0") & "%）" & "^p"
        Else
            ' 3レコード以上の場合はエラーとする
            Err.Raise 9999, "遅延損害金内容生成", "遅延損害金利率データが3レコードを超えています。最大3レコードまでしか対応していません。"
        End If
    Else
        ' データが存在しない場合のデフォルト値
        遅延損害金内容 = ""
    End If
    
    遅延損害金内容生成 = 遅延損害金内容
End Function

' 出力データから遅延損害金明細文字列を生成
Private Function 遅延損害金明細生成(wordDoc As Object, toolSheet As Worksheet, 出力データ As Variant) As String
    Dim 遅延損害金明細内容 As String
    Dim 遅延損害金明細補足 As String
    Dim i As Long
    Dim 順番 As Long
    Dim 最大順番 As Long
    Dim 表示用順番 As Long
    Dim 順番文字 As String
    
    ' データが配列かつ要素が存在するかチェック
    If Not IsArray(出力データ) Or UBound(出力データ, 1) < 1 Then
        遅延損害金明細生成 = ""
        Exit Function
    End If
    
    遅延損害金明細内容 = ""
    順番 = 0
    最大順番 = 0
    
    ' 最大順番を取得（遅延損害金列と遅損金_返済額列にデータがあるレコードの数）
    For i = 1 To UBound(出力データ, 1)
        If (出力データ(i, 出力_遅延損害金列) <> "" And 出力データ(i, 出力_遅延損害金列) <> 0) Or _
           (出力データ(i, 出力_遅損金_返済額列) <> "" And 出力データ(i, 出力_遅損金_返済額列) <> 0) Then
            最大順番 = 最大順番 + 1
        End If
    Next i
    
    ' 順番数字を格納する配列を定義
    Dim 順番配列() As Long
    Dim 内入れ順番配列() As Long

    ' 順番配列のサイズを設定
    If 最大順番 > 0 Then
        ReDim 順番配列(1 To 最大順番)
        ReDim 内入れ順番配列(1 To 最大順番)
    End If

    ' L列（利息金額）とR列（利息_返済額）に数値がある件数を初期値として設定
    表示用順番 = 0
    For i = 1 To UBound(出力データ, 1) - 1 ' 最後の行は合計行なので除外
        If 出力データ(i, 出力_利息金額列) <> "" And 出力データ(i, 出力_利息金額列) <> 0 Then ' L列：利息金額
            表示用順番 = 表示用順番 + 1
        End If
        If 出力データ(i, 出力_利息_返済額列) <> "" And 出力データ(i, 出力_利息_返済額列) <> 0 Then ' R列：利息_返済額
            表示用順番 = 表示用順番 + 1
        End If
    Next i

    Call ワード文字列置換(wordDoc, "{{遅延損害金明細1}}", "{{遅延損害金明細" & (表示用順番 + 1) & "}}")

    
    ' 出力データをループして遅延損害金明細を生成
    For i = 1 To UBound(出力データ, 1) - 1 ' 最後の行は合計行なので除外

        ' 遅損金返済額列にデータがある場合
        If 出力データ(i, 出力_遅損金_返済額列) <> "" And 出力データ(i, 出力_遅損金_返済額列) <> 0 Then ' S列：遅損金_返済額

            表示用順番 = 表示用順番 + 1
            内入れ順番配列(順番) = 表示用順番
            
            ' 順番を①②③形式に変換（共通関数を使用）
            順番文字 = 順番文字変換(表示用順番)
            
            遅延損害金明細内容 = 順番文字 & "▲" & Format(出力データ(i, 出力_遅損金_返済額列), "#,##0") & "円" & "^p" & _
                        Format(出力データ(i, 出力_返済日列), "yyyy年mm月dd日") & "に上記" & 順番配列範囲変換(順番配列) & "に係る遅延損害金の合計として内入返済" & "^p" & "{{遅延損害金明細" & CStr(表示用順番 + 1) & "}}"
            Call ワード文字列置換(wordDoc, "{{遅延損害金明細" & 表示用順番 & "}}", 遅延損害金明細内容)
        End If

        ' 遅延損害金列にデータがある場合
        If 出力データ(i, 出力_遅延損害金列) <> "" And 出力データ(i, 出力_遅延損害金列) <> 0 And _
          (出力データ(i, 出力_ステータス列) <> 期失劣後ステータス文字列 And 出力データ(i, 出力_ステータス列) <> 期限切れ劣後ステータス文字列) Then ' M列：遅延損害金
            順番 = 順番 + 1
            表示用順番 = 表示用順番 + 1
            順番配列(順番) = 表示用順番
            
            ' 順番を①②③形式に変換（共通関数を使用）
            順番文字 = 順番文字変換(表示用順番)
            
            If 出力データ(i, 出力_利率列) > 0 Then
                遅延損害金明細内容 = 順番文字 & Format(出力データ(i, 出力_遅延損害金列), "#,##0") & "円" & "^p" & _
                            "貸付金" & Format(出力データ(i, 出力_対象元金列), "#,##0") & "円に対する" & _
                            Format(出力データ(i, 出力_計算期間開始日列), "yyyy年mm月dd日") & "から" & _
                            Format(出力データ(i, 出力_計算期間終了日列), "yyyy年mm月dd日") & "まで" & _
                            出力データ(i, 出力_計算日数列) & "日間、年" & Format(出力データ(i, 出力_利率列) * 100, "0.0") & "%の割合による遅延損害金" & "^p" & "{{遅延損害金明細" & CStr(表示用順番 + 1) & "}}"
            Else
                遅延損害金明細内容 = 順番文字 & Format(出力データ(i, 出力_遅延損害金列), "#,##0") & "円" & "^p" & _
                            "貸付金" & Format(出力データ(i, 出力_対象元金列), "#,##0") & "円に対する" & _
                            Format(出力データ(i, 出力_計算期間開始日列), "yyyy年mm月dd日") & "から" & _
                            Format(出力データ(i, 出力_計算期間終了日列), "yyyy年mm月dd日") & "まで" & _
                            出力データ(i, 出力_計算日数列) & "日間、遅延損害金免除" & "^p" & "{{遅延損害金明細" & CStr(表示用順番 + 1) & "}}"
            END If

            Call ワード文字列置換(wordDoc, "{{遅延損害金明細" & 表示用順番 & "}}", 遅延損害金明細内容)

        End If
        
    Next i
    
    if UBound(出力データ, 1) > 2 Then
        i = UBound(出力データ, 1) - 1 

        if 出力データ(i, 出力_ステータス列) = 期失劣後ステータス文字列 Or 出力データ(i, 出力_ステータス列) = 期限切れ劣後ステータス文字列 Then

            ' 劣後債権
            順番 = 順番 + 1
            表示用順番 = 表示用順番 + 1
            順番配列(順番) = 表示用順番

            ' 順番を①②③形式に変換（共通関数を使用）
            順番文字 = 順番文字変換(表示用順番)    
            
            遅延損害金明細内容 = 順番文字 & "額未定（劣後債権）" & "^p" & _
                        "貸付金" & Format(出力データ(i, 出力_対象元金列), "#,##0") & "円に対する" & _
                        Format(出力データ(i, 出力_計算期間開始日列), "yyyy年mm月dd日") & "から" & "完済まで" & _                
                        ”年" & Format(出力データ(i, 出力_利率列) * 100, "0.0") & "%の割合による遅延損害金" & "^p"
            Call ワード文字列置換(wordDoc, "{{遅延損害金明細" & 表示用順番 & "}}", 遅延損害金明細内容)
        End If
    End If

    
    Dim 内入れ範囲文字列 As String
    内入れ範囲文字列 = 順番配列範囲変換(内入れ順番配列)
    
    If 内入れ範囲文字列 = "" Then
        遅延損害金明細補足 = 順番配列範囲変換(順番配列) & "の合計金額"
    Else
        遅延損害金明細補足 = 順番配列範囲変換(順番配列) & "の合計金額より内入" & 内入れ範囲文字列 & "を控除"
    End If
    Call ワード文字列置換(wordDoc, "{{遅延損害金明細補足}}", 遅延損害金明細補足)

    遅延損害金明細生成 = ""
End Function

' Word文字列置換サブルーチン
' 引数: Wordドキュメント、検索文字列、置換文字列
Private Sub ワード文字列置換(wordDoc As Object, 検索文字列 As String, 置換文字列 As String)
    On Error GoTo ErrorHandler
    
    With wordDoc.Content.Find
        .Text = 検索文字列
        .Replacement.Text = 置換文字列
        .Forward = True
        .Wrap = 1 ' wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=2 ' wdReplaceAll
    End With
    
    Exit Sub
    
ErrorHandler:
    ' エラーが発生した場合は無視して続行
    Err.Raise 9999, "ワード文字列置換", "ワード文字列置換でエラーが発生しました。"
End Sub



' 順番文字変換共通関数
' 引数: 順番（数値）
' 戻り値: ①②③形式の文字列
Public Function 順番文字変換(順番 As Long) As String
    Dim 順番文字 As String
    Select Case 順番
        Case 1: 順番文字 = "①"
        Case 2: 順番文字 = "②"
        Case 3: 順番文字 = "③"
        Case 4: 順番文字 = "④"
        Case 5: 順番文字 = "⑤"
        Case 6: 順番文字 = "⑥"
        Case 7: 順番文字 = "⑦"
        Case 8: 順番文字 = "⑧"
        Case 9: 順番文字 = "⑨"
        Case 10: 順番文字 = "⑩"
        Case 11: 順番文字 = "⑪"
        Case 12: 順番文字 = "⑫"
        Case 13: 順番文字 = "⑬"
        Case 14: 順番文字 = "⑭"
        Case 15: 順番文字 = "⑮"
        Case 16: 順番文字 = "⑯"
        Case 17: 順番文字 = "⑰"
        Case 18: 順番文字 = "⑱"
        Case 19: 順番文字 = "⑲"
        Case 20: 順番文字 = "⑳"
        Case 21 To 35
            順番文字 = ChrW(&H3250 + (順番 - 20))  ' （２１） = U+3251
        Case 36 To 50
            順番文字 = ChrW(&H32B0 + (順番 - 35))  ' （３６） = U+32B1
        Case Else: 順番文字 = "(" & 順番 & ")"
    End Select
    
    順番文字変換 = 順番文字
End Function


' 順番配列を範囲文字列に変換する共通関数
' 引数: 順番配列（数値の配列）
' 戻り値: ①～③、⑤～⑥、⑧ 形式の文字列
Public Function 順番配列範囲変換(順番配列 As Variant) As String
    ' 配列かどうか、および有効な範囲を持っているかをチェック
    If Not IsArray(順番配列) Or UBound(順番配列) < LBound(順番配列) Then
        順番配列範囲変換 = ""
        Exit Function
    End If
 
    ' === ステップ1: 0を除外した新しい配列を作成 ===
    Dim filteredArray() As Long  ' 0以外の値を格納する配列
    Dim i As Long, j As Long
    Dim count As Long            ' 0以外の要素数をカウント

    ' 0以外の要素の個数をカウント
    count = 0
    For i = LBound(順番配列) To UBound(順番配列)
        ' 値が数値であり、かつ0でない場合にカウント
        If IsNumeric(順番配列(i)) Then
            If 順番配列(i) <> 0 Then
                count = count + 1
            End If
        End If
    Next i

    ' 有効なデータが一つもなければ、空文字を返して終了
    If count = 0 Then
        順番配列範囲変換 = ""
        Exit Function
    End If

    ' filteredArray を必要なサイズに調整（1始まり）
    ReDim filteredArray(1 To count)

    ' 実際の値を新しい配列にコピー（0はスキップ）
    j = 1
    For i = LBound(順番配列) To UBound(順番配列)
        If IsNumeric(順番配列(i)) Then
            If 順番配列(i) <> 0 Then
                filteredArray(j) = CLng(順番配列(i))  ' 数値に変換して格納
                j = j + 1
            End If
        End If
    Next i

    ' === ステップ2: 新しい配列を昇順にソート ===
    Dim temp As Long
    For i = 1 To UBound(filteredArray) - 1
        For j = i + 1 To UBound(filteredArray)
            If filteredArray(i) > filteredArray(j) Then
                ' 値の入れ替え
                temp = filteredArray(i)
                filteredArray(i) = filteredArray(j)
                filteredArray(j) = temp
            End If
        Next j
    Next i

    ' === ステップ3: 連続する数字を範囲としてまとめる ===
    Dim 結果文字列 As String      ' 最終的な結果文字列
    Dim 範囲開始 As Long          ' 範囲の開始番号
    Dim 範囲終了 As Long          ' 範囲の終了番号

    結果文字列 = ""
    範囲開始 = filteredArray(1)   ' 最初の値を範囲の開始とする
    範囲終了 = 範囲開始

    ' 2番目以降の要素をチェック
    For i = 2 To UBound(filteredArray)
        If filteredArray(i) = 範囲終了 + 1 Then
            ' 数字が連続している場合：範囲を拡張
            範囲終了 = filteredArray(i)
        Else
            ' 連続が途切れた場合：現在の範囲を結果に追加
            If 結果文字列 <> "" Then
                結果文字列 = 結果文字列 & "、"
            End If

            If 範囲開始 = 範囲終了 Then
                ' 単一の数字（例：③）
                結果文字列 = 結果文字列 & 順番文字変換(範囲開始)
            Else
                ' 範囲（例：②～⑤）
                結果文字列 = 結果文字列 & 順番文字変換(範囲開始) & "～" & 順番文字変換(範囲終了)
            End If

            ' 新しい範囲の開始
            範囲開始 = filteredArray(i)
            範囲終了 = 範囲開始
        End If
    Next i

    ' 最後の範囲を結果に追加
    If 結果文字列 <> "" Then
        結果文字列 = 結果文字列 & "、"
    End If

    If 範囲開始 = 範囲終了 Then
        ' 単一の数字
        結果文字列 = 結果文字列 & 順番文字変換(範囲開始)
    Else
        ' 範囲
        結果文字列 = 結果文字列 & 順番文字変換(範囲開始) & "～" & 順番文字変換(範囲終了)
    End If

    ' 関数の戻り値を設定
    順番配列範囲変換 = 結果文字列
End Function