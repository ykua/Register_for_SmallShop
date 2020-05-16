Sub aggrigate_today()


'集計期間を取得するダイアログ

'集計するシートを選択
Sheets("MonsSales").Select

'開始期間の変数を定義
Dim dayselectStart As Date

'開始期間入力ダイアログ
dayselectStart = Date

'終了期間の変数を定義
Dim dayselectEnd As Date
   
'終了期間の値を正しい結果が得られるように調整
dayselectEnd = DateAdd("d", 1, Date)
    


'日付条件で絞り込み
Range("A1").AutoFilter Field:=1, Criteria1:=(">=" & dayselectStart), Operator:=xlAnd, Criteria2:=("<" & dayselectEnd)


'絞り込みたい取引先コードをダイアログで取得
Dim supplierID As String
If IMEStatus = vbIMEHiragana Then SendKeys "{Kanji}"
supplierID = Application.InputBox("絞り込みたい取引先のコードを入力してください")


'取引先コードで絞り込み
Range("A1").AutoFilter Field:=2, Criteria1:=(supplierID & "*")

'合計を計算する
result2 = WorksheetFunction.Subtotal(9, Sheets("MonsSales").Range("C:C"))

'売上点数を集計する
qty1 = WorksheetFunction.Subtotal(2, Sheets("MonsSales").Range("C:C"))


'計算結果をメッセージボックスに表示
'集計期間終了日の表示を算出日になるように調整
MsgBox "集計結果" & vbCrLf & "集計対象日： " & Date & vbCrLf & "取引先コード： " & supplierID & vbCrLf & vbCrLf & "販売数は " & qty1 & "点" & " です" & vbCrLf & "集計合計は " & result2 & "円" & " です"

    
'集計結果を「Result」シートに出力
    
    
'コピー前にセルの内容をクリア
Sheets("Result").Select
Range("A2:A5").ClearContents
    
'集計期間をResult A2セルに表示する
Sheets("Result").Range("A2") = Date & " to " & Date
'取引先をResult A3セルに表示する
Sheets("Result").Range("A3") = "Supplier ID  " & supplierID
    
    
'MonsSalesの値をResultシートにペースト
    
'コピー前にResultシートの入力値をクリア
Range("A6").Select
Selection.CurrentRegion.Select
Selection.ClearContents
'値をコピー
Sheets("MonsSales").Select
Range("A1").CurrentRegion.copy Sheets("Result").Range("A6")
    
    
'Resultシートに合計・手数料・支払額を表示
    '合計をセルに入力
    Sheets("Result").Select
    Range("C7").Select
    nTotalR1 = Cells(Rows.Count, "C").End(xlUp).Row
    '合計をセルに表示
    nTotal_printR1 = Cells(Rows.Count, "C").End(xlUp).Row + 1
    Cells(nTotal_printR1, 2) = "Total"
        
        
    '合計を求めてtoatal_valに格納
    total_val = WorksheetFunction.Sum(Range(Cells(7, 3), Cells(nTotal_printR1, 3)))
    '合計金額を表示
    Cells(nTotal_printR1, 3) = total_val
        
       
        
    
        


'オートフィルタを解除する
Worksheets("MonsSales").AutoFilterMode = False

'アクティブシートをInputにもどしA1セルをセレクトする
Sheets("Input").Select
Range("A1").Select

End Sub
