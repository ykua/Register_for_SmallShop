Sub aggrigate_suppliers()

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


    
'集計結果を「Result」シートに出力
    
    'コピー前に7行目までのセルの内容をクリア
    Sheets("Result").Select
    Range("A2:C7").ClearContents
    
    '集計期間終了日の表示を算出日になるように調整
    dayselectEndDisp = DateAdd("d", -1, dayselectEnd)
    
    '集計期間をResult A3セルに表示する
    Sheets("Result").Range("A3") = "SALES REPORT   " & dayselectStart & " to " & dayselectEndDisp
    
'何も入力していないとエラーが出るので無視する
On Error Resume Next
    '取引先をResult A4セルに表示する
    SupName = WorksheetFunction.VLookup(supplierID, Worksheets("Supplier").Range("A1:B100"), 2, Fales)
    Sheets("Result").Range("A4") = "Supplier    " & SupName & "(" & supplierID & ")"
        'すべて表示する場合にALLと記載する
            If Range("A4") = "Supplier    ()" Then
            Range("A4") = "Supplier    ALL"
                End If
        
        
    '書式を指定する
    Range("A2:C7").Font.Size = 9
    Range("A2:C7").Font.Name = "Futura Std Light"
    Range("B4").Font.Name = "AXIS Std L"

    
    
    'MonsSalesの値をResultシートにペースト
    
    'コピー前にResultシートの入力値(A8以降)をクリア
    Sheets("Result").Select
    Range("A8").Select
    Selection.CurrentRegion.Select
    Selection.ClearContents
    '値をResultシートのA8以降にコピー
    Sheets("MonsSales").Select
    Range("A1").CurrentRegion.copy Sheets("Result").Range("A8")
    
    
'区切り線を追加
        Sheets("Result").Select
        nItems_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
        'Cells(nItems_print, 1) = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
        Cells(nItems_print, 2) = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
        Cells(nItems_print, 3) = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
    
    
    '合計を追記する
        'Resultシートに合計・手数料・支払額を表示
        '合計をセルに入力
        Sheets("Result").Select
        Range("C9").Select
        nTotalR1 = Cells(Rows.Count, "C").End(xlUp).Row
        '合計をセルに表示
        nTotal_printR1 = Cells(Rows.Count, "C").End(xlUp).Row + 1
        Cells(nTotal_printR1, 2) = "Sales Subtotal"
        Cells(nTotal_printR1, 2).Offset(2, 0) = "Payment Total"
        
            '取引先の入力がない場合は合計の表示を「Sales Subtotal」から「Total」にする
            If supplierID = "" Then
            Cells(nTotal_printR1, 2) = "Total"
                End If
            
        
        
        '合計を求めてtoatal_valに格納
        total_val = WorksheetFunction.Sum(Range(Cells(7, 3), Cells(nTotal_printR1, 3)))
        
        
    '売掛（物品）の合計を集計する

    '売掛対象を集計する
        Dim RcvYtotal As Long
            RcvYtotal = 0
            qtyY = 0
        
            '値の入力された最終行を取得「 ---- 」の行は無視
            nRcvtY = Cells(Rows.Count, "B").End(xlUp).Row - 1
            
            '売掛対象の合計を求めるループ処理
            For i = 8 To nRcvtY
            If Mid(Cells(i, 2), 3, 1) = "Y" Then
                RcvYtotal = RcvYtotal + Cells(i, 3)
                '物品売掛の回数をカウント
                qtyY = qtyY + 1
            
            End If
            
            Next
            
            
    '売掛（サービス）の合計を集計する

    '売掛対象を集計する
        Dim RcvZtotal As Long
            RcvZtotal = 0
            qtyZ = 0
        
            '値の入力された最終行を取得「 ---- 」の行は無視
            nRcvtZ = Cells(Rows.Count, "B").End(xlUp).Row - 1
            
            '売掛対象の合計を求めるループ処理
            For i = 8 To nRcvtZ
            If Mid(Cells(i, 2), 3, 1) = "Z" Then
                RcvZtotal = RcvZtotal + Cells(i, 3)
                'サービス売掛の回数をカウント
                qtyZ = qtyZ + 1
            End If
            
            Next
            
            
    '立替金の合計を集計する

    '立替対象を集計する
        Dim RcvWtotal As Long
            RcvWtotal = 0
            qtyW = 0
        
            '値の入力された最終行を取得「 ---- 」の行は無視
            nRcvtW = Cells(Rows.Count, "B").End(xlUp).Row - 1
            
            '立替対象の合計を求めるループ処理
            For i = 8 To nRcvtW
            If Mid(Cells(i, 2), 3, 1) = "W" Then
                RcvWtotal = RcvWtotal + Cells(i, 3)
                '立替の回数をカウント
                qtyW = qtyW + 1
            End If
            
            Next


    '立替対象を集計する
        Dim RcvXtotal As Long
            RcvXtotal = 0
            qtyX = 0
        
            '値の入力された最終行を取得「 ---- 」の行は無視
            nRcvtX = Cells(Rows.Count, "B").End(xlUp).Row - 1
            
            '立替対象の合計を求めるループ処理
            For i = 8 To nRcvtX
            If Mid(Cells(i, 2), 3, 1) = "X" Then
                RcvXtotal = RcvXtotal + Cells(i, 3)
                '立替の回数をカウント
                qtyX = qtyX + 1
            End If
            
            Next
            

'売掛・立替を除いた正味の販売個数を計算
qty1 = qty0 - qtyY - qtyZ - qtyW - qtyX

'正味の販売アイテム数をA5セルに表示する
Sheets("Result").Range("A5") = "Sales qty   " & qty1 & " pcs."



        '合計金額を表示
            '売掛合計を正の値にして売り上げ合計を合わせる
            RcvYtotal = RcvYtotal * -1
            RcvZtotal = RcvZtotal * -1
            RcvWtotal = RcvWtotal * -1
            
        '売掛相殺と立替を除いた販売合計を計算する
        total_val = total_val + RcvYtotal + RcvZtotal + RcvWtotal
            
        Cells(nTotal_printR1, 3) = total_val
        
        '仕入先に対応する手数料率をSupplierシートから取得する
        supplierRate = WorksheetFunction.VLookup(supplierID, Worksheets("Supplier").Range("A1:C100"), 3, Fales)
        
        
        '手数料項目をセルに表示
        Cells(nTotal_printR1, 2).Offset(1, 0) = "Commission(" & supplierRate & "%)"
        
        '手数料金額を表示
        charge1 = total_val * supplierRate / 100
        '手数料金額を小数点以下を繰り上げ
        charge2 = WorksheetFunction.RoundUp(charge1, 0) * -1
        Cells(nTotal_printR1, 3).Offset(1, 0) = charge2
        
        
        '最下段セルの位置を取得
        nTotal_printR2 = Cells(Rows.Count, "C").End(xlUp).Row + 1
        '売掛（物品）の合計額を表示
        Cells(nTotal_printR2, 2) = "A/R(Goods)"
        Cells(nTotal_printR2, 3) = RcvYtotal * -1
            '売掛がない場合は項目を削除
            If RcvYtotal = 0 Then
                Range(Cells(nTotal_printR2, 2), Cells(nTotal_printR2, 3)).Clear
            End If
            
            
        '最下段セルの位置を取得
        nTotal_printR3 = Cells(Rows.Count, "C").End(xlUp).Row + 1
        '売掛（サービス）の合計額を表示
        Cells(nTotal_printR3, 2) = "A/R(Services)"
        Cells(nTotal_printR3, 3) = RcvZtotal * -1
            '売掛がない場合は項目を削除
            If RcvZtotal = 0 Then
                Range(Cells(nTotal_printR3, 2), Cells(nTotal_printR3, 3)).Clear
            End If
            
            
        '最下段セルの位置を取得
        nTotal_printAdP1 = Cells(Rows.Count, "C").End(xlUp).Row + 1
        '立替の合計額を表示
        Cells(nTotal_printAdP1, 2) = "Adv. Paid"
        Cells(nTotal_printAdP1, 3) = RcvWtotal * -1
            '立替がない場合は項目を削除
            If RcvWtotal = 0 Then
                Range(Cells(nTotal_printAdP1, 2), Cells(nTotal_printAdP1, 3)).Clear
            End If
           
        
        '最下段セルの位置を取得
        nTotal_printR4 = Cells(Rows.Count, "C").End(xlUp).Row + 1
        '支払金額を表示
        Cells(nTotal_printR4, 2) = "Payment Total"
        Cells(nTotal_printR4, 3) = total_val + charge2 - RcvYtotal - RcvZtotal - RcvWtotal
        
        
        '全体売り上げの場合に手数料と支払合計の表示を消去する
            If supplierID = "" Then
            nTotal_Clear = Cells(Rows.Count, "C").End(xlUp).Row
            Range(Cells(nTotal_printR1 + 1, 1), Cells(nTotal_Clear, 3)).Clear
                End If

    
    
    'フォント指定をする
    '書式設定エリアを定義
    Sheets("Result").Select

    'Aカラムの範囲を定義
    Dim formatAreaA As Range
    Set formatAreaA = Range("A8", Cells(Rows.Count, 1).End(xlUp))
    
    'Bカラムの範囲を定義
    Dim formatAreaB As Range
    Set formatAreaB = Range("B8", Cells(Rows.Count, 2).End(xlUp))
    
    
    'Cカラムの範囲を定義
    Dim formatAreaC As Range
    Set formatAreaC = Range("C8", Cells(Rows.Count, 3).End(xlUp))
    
    
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
    Sheets("Result").PrintOut ActivePrinter:="POS-80"
   

'オートフィルタを解除する
Worksheets("MonsSales").AutoFilterMode = False


'計算結果をメッセージボックスに表示
MsgBox "集計結果" & vbCrLf & "開始期間： " & dayselectStart & vbCrLf & "終了期間： " & dayselectEndDisp & vbCrLf & "取引先コード： " & supplierID & vbCrLf & vbCrLf & "販売数は " & qty1 & "点" & " です" & vbCrLf & "集計合計は " & total_val & "円" & " です"


'アクティブシートをInputにもどす
Sheets("Input").Select
Range("A1").Select


End Sub

