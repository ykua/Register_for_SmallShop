Sub aggrigate()

'集計期間を取得するダイアログ

    '集計するシートを選択
    Sheets("MonsSales").Select

    '開始期間の変数をダイアログから取得
    Dim dayselectStart As Date

    'ダイアログの入力をIMEオフに設定
    If IMEStatus = vbIMEHiragana Then SendKeys "{Kanji}"
    '開始期間入力ダイアログ
    dayselectStart = Application.InputBox("絞り込みの開始期間を入力してください", Format(Date, "yyyyy/mm/dd"))

    '終了期間入力ダイアログ
    Dim dayselectEnd As Date
    'ダイアログの入力をIMEオフに設定
    If IMEStatus = vbIMEHiragana Then SendKeys "{Kanji}"
    '終了期間入力ダイアログ
    dayselectEnd = Application.InputBox("絞り込みの終了期間を入力してください", Format(Date, "yyyyy/mm/dd"))
    dayselectEnd = DateAdd("d", 1, dayselectEnd)
    


'日付条件で絞り込み
Range("A1").AutoFilter Field:=1, Criteria1:=(">=" & dayselectStart), Operator:=xlAnd, Criteria2:=("<" & dayselectEnd)

'絞り込みたい取引先コードをダイアログで取得
Dim supplierID As String
supplierID = Application.InputBox("絞り込みたい取引先のコードを入力してください")


'取引先コードで絞り込み
Range("A1").AutoFilter Field:=2, Criteria1:=(supplierID & "*")

'合計を計算する
result1 = WorksheetFunction.Subtotal(9, Sheets("MonsSales").Range("C:C"))


'売上点数を集計する
qty0 = WorksheetFunction.Subtotal(2, Sheets("MonsSales").Range("C:C"))

'計算結果をメッセージボックスに表示
'集計期間終了日の表示を算出日になるように調整
dayselectEndDisp = DateAdd("d", -1, dayselectEnd)
MsgBox "集計結果" & vbCrLf & "開始期間： " & dayselectStart & vbCrLf & "終了期間： " & dayselectEndDisp & vbCrLf & "取引先コード： " & supplierID & vbCrLf & vbCrLf & "販売数は " & qty0 & "点" & " です" & vbCrLf & "集計合計は " & result1 & "円" & " です"

    
'集計結果を「Result」シートに出力
    
    'コピー前にセルの内容をクリア
    Sheets("Result").Select
    Range("A2:C5").ClearContents
    
    '集計期間をResult A3セルに表示する
    Sheets("Result").Range("A3") = "SALES REPORT   " & dayselectStart & " to " & dayselectEndDisp
    
    '取引先をResult A4セルに表示する
    '何も入力していないとエラーが出るので無視する
    On Error Resume Next

    SupName = WorksheetFunction.VLookup(supplierID, Worksheets("Supplier").Range("A1:B100"), 2, Fales)
    Sheets("Result").Range("A4") = "Supplier   " & SupName & "(" & supplierID & ")"
    
    '書式を指定する
    Range("A2:C5").Font.Size = 9
    Range("A2:C5").Font.Name = "Futura Std Light"
    Range("B4").Font.Name = "AXIS Std L"

    
    
    'MonsSalesの値をResultシートにペースト
    
    'コピー前にResultシートの入力値(A6以降)をクリア
    Sheets("Result").Select
    Range("A6").Select
    Selection.CurrentRegion.Select
    Selection.ClearContents
    '値をコピー
    Sheets("MonsSales").Select
    Range("A1").CurrentRegion.copy Sheets("Result").Range("A6")
    
    
    '合計を追記する
        'Resultシートに合計・手数料・支払額を表示
        '合計をセルに入力
        Sheets("Result").Select
        Range("C7").Select
        nTotalR1 = Cells(Rows.Count, "C").End(xlUp).Row
        '合計・支払額項目 をセルに表示
        nTotal_printR1 = Cells(Rows.Count, "C").End(xlUp).Row + 1
        Cells(nTotal_printR1, 2) = "Total"
        'Cells(nTotal_printR1, 2).Offset(2, 0) = "Payment Total"
        
        
        '合計を求めてtoatal_valに格納
        total_val = WorksheetFunction.Sum(Range(Cells(7, 3), Cells(nTotal_printR1, 3)))
        '合計金額を表示
        Cells(nTotal_printR1, 3) = total_val
        
        '仕入先に対応する手数料率をSupplierシートから取得する
        'Set searchArea = Sheets("Supplier").Range("A1:C100")
        'supplierRate = Application.WorksheetFunction.VLookup(supplierID, searchArea, 3, Fales)
        
        
        '手数料項目をセルに表示
        'Cells(nTotal_printR1, 2).Offset(1, 0) = "Sales Charge(" & supplierRate & "%)"
        
        '手数料金額を表示
        'charge1 = total_val * supplierRate / 100
        'Cells(nTotal_printR1, 3).Offset(1, 0) = charge1
        
        
        '支払金額を表示
        'Cells(nTotal_printR1, 3).Offset(2, 0) = total_val - charge1

    
    
    'フォント指定をする
    '書式設定エリアを定義
    Sheets("Result").Select

    'Aカラムの範囲を定義
    Dim formatAreaA As Range
    Set formatAreaA = Range("A6", Cells(Rows.Count, 1).End(xlUp))
    
    'Bカラムの範囲を定義
    Dim formatAreaB As Range
    Set formatAreaB = Range("B6", Cells(Rows.Count, 2).End(xlUp))
    
    
    'Cカラムの範囲を定義
    Dim formatAreaC As Range
    Set formatAreaC = Range("C6", Cells(Rows.Count, 3).End(xlUp))
    
    
    'Aカラムの書式を設定
    formatAreaA.Font.Size = 9
    formatAreaA.Font.Name = "Futura Std Light"
    formatAreaA.HorizontalAlignment = xlLeft
    
    'Bカラムの書式を設定
    formatAreaB.Font.Size = 9
    formatAreaB.Font.Name = "Futura Std Light"
    formatAreaB.HorizontalAlignment = xlLeft

    'Cカラムの書式を設定
    formatAreaC.Font.Size = 9
    formatAreaC.Font.Name = "Futura Std Light"
    formatAreaC.HorizontalAlignment = xlRight

    'Cカラムの数値にカンマを追加
    formatAreaC.NumberFormatLocal = "###,##0"
    
    
    
    
    'レシートをPOSプリンタを指定して印刷する
    'Sheets("Result").PrintOut ActivePrinter:="POS-80"
   

'オートフィルタを解除する
Worksheets("MonsSales").AutoFilterMode = False

'アクティブシートをInputにもどす
Sheets("Input").Select
Range("A1").Select

End Sub
