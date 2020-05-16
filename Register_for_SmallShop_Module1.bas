'通常課税の税率を定義
Public Const StdTaxRate = 1.1
''通常課税の税率を定義(%)
Public Const StdTax = 10

'軽減課税の税率を定義
Public Const ReduceTaxRate = 1.08
'軽減課税の税率を定義（％）
Public Const ReduceTax = 8



Sub checkout()


    
    'ブック保護からVBAアクセスを除外
    'ThisBook.Worksheets("Input").Protect.UserInterfaceOnly = True
    

    'シートの名前に重複がないか確認し，なければ新しく月のシートを作成する
    'Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = Format(Now, "YYYYMM")
   
    '合計金額の計算
    Sheets("Input").Select
    Dim Total As Long
   
    Range("C1").Select
    nBtotal = Cells(Rows.Count, "C").End(xlUp).Row + 1
    Total = WorksheetFunction.Sum(Range(Cells(1, 3), Cells(nBtotal, 3)))
    
    '合計金額をダイアログに表示し，お預かり金額の入力する
    Dim buf1 As String
    'ダイアログの入力をIMEオフにする
    If IMEStatus = vbIMEHiragana Then SendKeys "{Kanji}"
    
    '合計金額表示とお預かり金額入力ダイアログ
    buf1 = InputBox("合計金額は" & Total & "円です。" & vbCrLf & "お預かり金額を入力してください。")
    
    'お釣りを算出し，ダイアログに表示
    Dim buf2 As String
        '合計がマイナスの場合の処理
        If Total < 0 Then
        turi = buf1 + Total
        
        Else
            turi = buf1 - Total
            
        End If
        
    buf2 = MsgBox("おつりは" & turi & "円です")
    
    
    'InputシートのAカラム（商品コード）の入力値をコピー
    Sheets("Input").Select
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.copy
    
    'Inputシートからコピーされた値をMonsSalesシートのBカラムの入力されているセルの下にペースト
    Sheets("MonsSales").Select
    Range("B1").Select
    nBcopy = Cells(Rows.Count, "B").End(xlUp).Row + 1
    Range("B" & nBcopy).Select
    ActiveSheet.Paste
    
    
    'InputシートのCカラムの入力値（価格）をコピー
    Sheets("Input").Select
    Range("C1").Select
    Selection.CurrentRegion.Select
    Selection.copy
    
    'Inputシートからコピーされた値をMonsSalesシートのBカラムの入力されているセルの下にペースト
    Sheets("MonsSales").Select
    Range("C1").Select
    'Selection.End(xlDown).Offset(1, 0).Select ← コンパイルエラーになる
    nCcopy = Cells(Rows.Count, "C").End(xlUp).Row + 1
    Range("C" & nCcopy).Select
    ActiveSheet.Paste
    
    'MonsSalesシートのAカラムに日時データを入力
    Sheets("MonsSales").Select
    
    Range("B1").Select
    nAtime = Cells(Rows.Count, "A").End(xlUp).Row + 1
    nBtime = Cells(Rows.Count, "B").End(xlUp).Row
        
    
    Range(Cells(nAtime, 1), Cells(nBtime, 1)).Value = Now
    
    
    'レシート印刷用のシートに値をコピー
       
        'Receiptシートの前回入力値の4行目と入力行目以降をクリア
        Sheets("Receipt").Select
        Range("A6").Select
        Selection.CurrentRegion.Select
        Selection.ClearContents
        Range("A4").Select
        Selection.CurrentRegion.Select
        Selection.ClearContents
        
        'InputシートのEカラム（商品カテゴリ＋仕入先）の入力値をコピー
        Sheets("Input").Select
        Range("E1").Select
        Selection.CurrentRegion.Select
        Selection.copy
        

        'Receiptシートに貼り付け
        Sheets("Receipt").Select
        Range("A6").Select
        ActiveSheet.Paste


        'Inputシートの価格をコピー
        Sheets("Input").Select
        Range("C1").Select
        Selection.CurrentRegion.Select
        Selection.copy

        'Receiptシートに貼り付け
        Sheets("Receipt").Select
        Range("B6").Select
        ActiveSheet.Paste
        

        'レシートに日付を表示
        Sheets("Receipt").Select
        Range("A4") = "Date    " & Format(Now, "MMM DD YYYY HH:MM")
        
        
        'フォントを指定
        Range("A2:B5").Font.Size = 10
        Range("A2:B5").Font.Name = "Futura Std Light"

        
        '区切り線を追加
        Sheets("Receipt").Select
        nItems_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
        Cells(nItems_print, 1) = "-------------------------------------------------"
        Cells(nItems_print, 2) = "-------------------------------------------------"
        
'レシートの内容の売掛の合計を集計する

        '売掛対象を集計する
        Dim RcvYtotal As Long
            RcvYtotal = 0
        
            '値の入力された最終行を取得「 ---- 」の行は無視
            nRcvt = Cells(Rows.Count, "B").End(xlUp).Row - 1
            
            '売掛対象の合計を求めるループ処理
            For i = 6 To nRcvt
            If Mid(Cells(i, 1), 3, 2) = "売掛" Then
                RcvYtotal = RcvYtotal + Cells(i, 2)
                RcvChange = Cells(i, 2) * -1
                Cells(i, 2) = RcvChange
            
            End If
            
            Next
            
'レシートの内容の立替の合計を集計する

        '立替対象を集計する
        Dim RcvWtotal As Long
            RcvWtotal = 0
        
            '値の入力された最終行を取得「 ---- 」の行は無視
            nRcvtW = Cells(Rows.Count, "B").End(xlUp).Row - 1
            
            '立替対象の合計を求めるループ処理
            For i = 6 To nRcvtW
            If Mid(Cells(i, 1), 3, 3) = "立替金" Then
                RcvWtotal = RcvWtotal + Cells(i, 2)
                RcvChangeW = Cells(i, 2) * -1
                Cells(i, 2) = RcvChangeW
            
            End If
            
            Next


'Receiptシートに合計を表示


'課税計算をする
            
        Sheets("Receipt").Select

        '軽減課税対象を集計する
        Dim RTtotal As Long
            RTtotal = 0
        
            '値の入力された最終行を取得「 ---- 」の行は無視
            nRTt = Cells(Rows.Count, "B").End(xlUp).Row - 1
            
            '軽減課税対象の合計を求めるループ処理
            For i = 6 To nRTt
            If Left(Cells(i, 1), 1) = "R" Then
                RTtotal = RTtotal + Cells(i, 2)
            
            End If
            
            Next
            
        
        '軽減課税対象額を表示
            '軽減課税対象をセルに入力
            Sheets("Receipt").Select
            nTotalRT_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
            Cells(nTotalRT_print, 1) = "通常課税対象 R 小計"
            Cells(nTotalRT_print, 2) = RTtotal
 
        '軽減課税対象の課税額を表示
            '軽減税の合計額を計算
            RTc = RTtotal / ReduceTaxRate
            RT = RTtotal - RTc
            '小数点以下を切り捨て
            RT = WorksheetFunction.RoundDown(RT, 0)
            'セルに軽減税額の計算値を入力
            nTotalRTP_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
            Cells(nTotalRTP_print, 1) = "内消費税（" & ReduceTax & "%）"
            Cells(nTotalRTP_print, 2) = RT
                
                '課税対象がなければセルを削除
                nRTt = Cells(Rows.Count, "B").End(xlUp).Row - 1
                                
                If Cells(nRTt, 2) = 0 Then
                    Range(Cells(nRTt, 2), Cells(nRTt + 1, 1)).Clear
                
                End If
                    
                    
                    
                    
    '通常課税対象を集計する
        Dim STtotal As Long
            STtotal = 0
        
            '値の入力された最終行を取得「 ---- 」の行は無視
            nSTt = Cells(Rows.Count, "B").End(xlUp).Row - 1
            
            '軽減課税対象の合計を求めるループ処理
            For i = 6 To nSTt
            If Left(Cells(i, 1), 1) = "S" Then
                STtotal = STtotal + Cells(i, 2)
            
            End If
            
            Next
            
        
        '通常課税対象額を表示
            '通常課税対象をセルに入力
            Sheets("Receipt").Select
            nTotalST_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
            Cells(nTotalST_print, 1) = "通常課税対象 S 小計"
            Cells(nTotalST_print, 2) = STtotal
 
        '通常課税対象の課税額を表示
            '通常消費税の合計額を計算
            STc = STtotal / StdTaxRate
            ST = STtotal - STc
            '小数点以下を切り捨て
            ST = WorksheetFunction.RoundDown(ST, 0)
            'セルに軽減税額の計算値を入力
            nTotalSTP_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
            Cells(nTotalSTP_print, 1) = "内消費税（" & StdTax & "%）"
            Cells(nTotalSTP_print, 2) = ST
                
            '課税対象がなければセルを削除
                nSTt = Cells(Rows.Count, "B").End(xlUp).Row - 1
                               
                If STtotal = 0 Then
                    Range(Cells(nSTt, 2), Cells(nSTt + 1, 1)).Clear
                    
                    End If




'非課税計算をする
            
        Sheets("Receipt").Select

        '非課税対象を集計する
        Dim NTtotal As Long
            NTtotal = 0
        
            '値の入力された最終行を取得「 ---- 」の行は無視
            nNTt = Cells(Rows.Count, "B").End(xlUp).Row - 1
            
            '軽減課税対象の合計を求めるループ処理
            For i = 6 To nNTt
            If Left(Cells(i, 1), 1) = "N" Then
                NTtotal = NTtotal + Cells(i, 2)
            
            End If
            
            Next
            
        
    '非課税対象額を表示
            '非課税対象をセルに入力
            Sheets("Receipt").Select
            nTotalNT_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
            Cells(nTotalNT_print, 1) = "非課税対象 N 小計"
            Cells(nTotalNT_print, 2) = NTtotal
 
       
           
                
            '課税対象がなければセルを削除
                nNTt = Cells(Rows.Count, "B").End(xlUp).Row
                
                If NTtotal = 0 Then
                    Range(Cells(nNTt, 2), Cells(nNTt, 1)).Clear
                
                End If
                
                
                
'不課税計算をする
            
        Sheets("Receipt").Select

        '不課税対象を集計する
        Dim UTtotal As Long
            UTtotal = 0
        
            '値の入力された最終行を取得「 ---- 」の行は無視
            nUTt = Cells(Rows.Count, "B").End(xlUp).Row - 1
            
            '不課税対象の合計を求めるループ処理
            For i = 6 To nUTt
            If Left(Cells(i, 1), 1) = "U" Then
                UTtotal = UTtotal + Cells(i, 2)
            
            End If
            
            Next
            
        
    '不課税対象額を表示
            '不課税対象をセルに入力
            Sheets("Receipt").Select
            nTotalUT_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
            Cells(nTotalUT_print, 1) = "不課税対象 U 小計"
            Cells(nTotalUT_print, 2) = UTtotal
 
       
           
                
            '課税対象がなければセルを削除
                nUTt = Cells(Rows.Count, "B").End(xlUp).Row
                
                If UTtotal = 0 Then
                    Range(Cells(nUTt, 2), Cells(nUTt, 1)).Clear
                
                End If
                
                
                
'区切り線を追加
        Sheets("Receipt").Select
        nSubTotal_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
        Cells(nSubTotal_print, 1) = "-------------------------------------------------"
        Cells(nSubTotal_print, 2) = "-------------------------------------------------"

        
            
    '合計をセルに入力
        Sheets("Receipt").Select
        nTotal = Cells(Rows.Count, "B").End(xlUp).Row
        '合計表示位置
        nTotal_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
        Cells(nTotal_print, 1) = "合計"
        '合計を表示
        Cells(nTotal_print, 2) = Total
        
        'お預かり金額を表示
        nCash_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
        Cells(nCash_print, 1) = "現金"
        Cells(nCash_print, 2) = buf1
        
        'お釣り金額を表示
        nTuri_print = Cells(Rows.Count, "B").End(xlUp).Row + 1
        Cells(nTuri_print, 1) = "おつり"
        Cells(nTuri_print, 2) = turi

        
        
        'フォント指定をする
        Range("A6:B200").Font.Size = 9
        Range("A6:B200").Font.Name = "AXIS Std L"
        
        '数値にカンマを追加
        Range("B6:B200").NumberFormatLocal = "###,##0"
        
        '合計以下の数値にJPYを追記
        Sheets("Receipt").Select
        nJPYt = Cells(Rows.Count, "B").End(xlUp).Row 'お釣りのセルを取得
        nJPYp = Cells(Rows.Count, "B").End(xlUp).Row - 1 'お預かりのセルを取得
        nJPYc = Cells(Rows.Count, "B").End(xlUp).Row - 2  '合計のセルを取得
        'Cells(nJPYt, 2).Value = "JPY " & Format(Cells(nJPYt, 2).Value, "###,###")
        'Cells(nJPYp, 2).Value = "JPY" & Format(Cells(nJPYp, 2).Value, "###,###")
        Cells(nJPYc, 2).Value = "JPY" & Format(Cells(nJPYc, 2).Value, "###,###")
        
        
        '表示を右詰めにする
        Range("B6:B200").HorizontalAlignment = xlRight
        
        
 '売掛対象がが合計と同額なら「合計」表示を売掛合計として「現金」，「お釣り」を表示しない
                nRcvtE = Cells(Rows.Count, "B").End(xlUp).Row
                               
                If RcvYtotal + RcvWtotal = Total Then
                    Range(Cells(nRcvtE, 2), Cells(nRcvtE - 1, 1)).Clear
                    Cells(nRcvtE - 2, 1) = "合計"
                    Cells(nRcvtE - 2, 2) = RcvYtotal * -1 + RcvWtotal * -1
                    Cells(nRcvtE - 2, 2).Value = "JPY" & Format(Cells(nRcvtE - 2, 2).Value, "###,###")
                End If

                
        
        
        'レシートをPOSプリンタを指定して印刷する
        Sheets("Receipt").PrintOut ActivePrinter:="POS-80"
        
             
    
    'Inputシートにもどり商品コードをクリアしA1セルをセレクトする
    Sheets("Input").Select
    Worksheets("Input").Cells.Clear
    Range("A1").Select
    
    'ワークブックを保存
    ActiveWorkbook.Save
    
    
    
End Sub
