// ListFeild
(table_, index_, feilds_, values_) => 
let
    ret =if( index_< List.Count(feilds_)) then
        // フィールド数index_が指定以下ならカラム追加処理
            let 
                res =Table.ExpandRecordColumn(table_, feilds_{index_}, {"value"}, {feilds_{index_}}),
                textValue =Text.From(values_{index_}),

                rtn =if( textValue="") then
                        // 単一フィールド
                        @ListFeild(res, index_ +1, feilds_,values_)
                    else if( textValue="value") then
                        // 選択フィールド
                         let
                            dep1= Table.ExpandListColumn(res, feilds_{index_}),
                            dep2= @ListFeild(dep1, index_ +1, feilds_,values_)            
                         in
                            dep2
                    else
                        // ユーザなどのフィールドの場合、展開を入れる
                         let
                            dep1= Table.ExpandListColumn(res, feilds_{index_}),
                            dep2= Table.ExpandRecordColumn(dep1, feilds_{index_}, {textValue}, {feilds_{index_}&"."&textValue}),
                            dep3= @ListFeild(dep2, index_ +1, feilds_,values_)            
                         in
                            dep3
            in 
                rtn
        else
            // フィールド数index_が指定数になったため再帰終了
            table_
in
    ret

// ExcelFromKintone
(offset_)=> 
let
    // Excelデータ
    ParamData = Excel.CurrentWorkbook(){[Name="基本情報"]}[Content][パラメータ情報],
    FeildData = Excel.CurrentWorkbook(){[Name="表示フィールド"]}[Content][フィールド名称],
    // nullの置き換え
    ValueTemp = Excel.CurrentWorkbook(){[Name="表示フィールド"]}[Content][形式],
    ValueData = List.ReplaceMatchingItems(ValueTemp,{{null,""}}),

    // Web Url
    url =Text.From(ParamData {0})&"k/v1/records.json" & "?app="&Text.From(ParamData {2}) & "&query=limit 500 offset "&Text.From(offset_),
    // ヘッダーと呼び出し
    web = Json.Document(Web.Contents(url, [Headers=[#"X-Cybozu-API-Token"=Text.From(ParamData {1})]])),
    // 必要なデータの取り出し
    TableWeb =web[records],
    
    // データの展開
    TableFromList = Table.FromList(TableWeb , Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    TableExpand   = Table.ExpandRecordColumn(TableFromList , "Column1", FeildData ,FeildData ),
    // フィールド(カラム)の展開
    TableData     = ListFeild(TableExpand , 0, FeildData ,ValueData ),
    
    ret =if( Table.RowCount(TableData     )=500) then 
            // データ数がMaxの呼び出しの場合、オフセットを500ずらして再帰
            Table.Combine({TableData,@ExcelFromKintone(offset_+500)})
        else
            // 500未満の場合は再帰せず終了
            TableData     
in
    ret

// Kintone情報
let
    ソース = ExcelFromKintone(0)
in
    ソース