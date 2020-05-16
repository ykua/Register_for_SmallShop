Sub Worksheet_Change(ByVal Target As Range)


'シートのAカラムに入力された値を値段に変換する処理


If Intersect(Target, Range("A1:A10000")) Is Nothing Then
        Exit Sub
    Else

    '最新のバーコードの値が入力されたセルに移動
    Sheets("Input").Select
    nAinput = Cells(Rows.Count, "A").End(xlUp).Row
    Cells(nAinput, 1).Select

    'バーコード値の右から6桁を取得してPriceに代入
    input_a = Cells(nAinput, 1).Value
    cal_v = Right(input_a, 6)
    Price = Val(cal_v)
             
    
    'CatNにバーコード値の前から3文字目を抽出してカテゴリを定義
    CatN = Mid(input_a, 3, 1)
    
    'TaxCにバーコード値の前から3文字目を抽出して課税コードを定義
    TaxN = Mid(input_a, 4, 1)