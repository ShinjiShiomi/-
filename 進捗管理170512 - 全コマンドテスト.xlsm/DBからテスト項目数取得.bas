Attribute VB_Name = "DBからテスト項目数取得"
Option Explicit
'DB内からデータを抽出し、[抽出用]シートのリストオブジェクトに貼り付ける。


'LOの変数を定義
'Dim TableName As String     'テーブル名
Dim clmName() As Variant    '列名の配列
'Dim fstRow As Integer       '開始行（列名のある行）
'Dim lstRow As Integer       '最終行
'Dim fstClm As Integer       '開始列
'Dim lstClm As Integer       '最終列
'Dim rowCnt As Integer       '行数（列名の行は含まない）
'Dim clmCnt As Integer       '列数


'カテゴリ、担当者ごとに集計をする
Dim CategoryClm As Integer      '[カテゴリ]列の列数
Dim PersonClm As Integer        '[担当者]列の列数
Dim CommandClm As Integer       '[コマンド]の列数
Dim OptionClm As Integer        '[オプション数]の列数
Dim DialogClm As Integer        '[ダイアログ数]の列数


Sub 抽出用LO情報取得()
'[抽出用]シートの情報を取得。リストオブジェクトの開始行、開始列、行数、列数、最終行、最終列、列名（配列）
    Worksheets("抽出用").Activate
        
    Call テーブル情報取得
'
'    With Worksheets("全コマンド").ListObjects("" & tblName & "")
'        fstRow = .Range.Cells(1, 1).Row
'        fstClm = .Range.Cells(1, 1).Column
'        rowCnt = .ListRows.Count
'        clmCnt = .ListColumns.Count
'        lstRow = .Range.Cells(rowCnt + 1, 1).Row
'        lstClm = .Range.Cells(1, clmCnt).Column
        
        Dim i As Integer
        Dim j As Integer
        j = 1
        
        For i = fstClm To lstClm
            ReDim Preserve clmName(j)
            clmName(j) = Worksheets("抽出用").Cells(hdrRow, i).Value
            
            j = j + 1
            'jをインクリメント
        Next
      
'    End With

    Worksheets("進捗管理").Activate

End Sub

Sub 抽出用LOに貼り付け()
    Call 抽出用LO情報取得
    Worksheets("抽出用").Range("抽出用テーブル").Clear
    Worksheets("抽出用").ListObjects(1).Resize Range(Cells(hdrRow, fstClm), Cells(hdrRow + 1, lstClm))
    Call データベース接続
    Call 選択バージョン取得
    Call 選択重要度取得

    Dim i As Integer
    Dim SQL1 As String
        For i = 1 To lstClm
            SQL1 = SQL1 & clmName(i) & ","
        Next
    '抽出するフィールド名のSQL文を準備
        
    SQL1 = Left(SQL1, Len(SQL1) - 1)
    'カンマを除去

    Dim selSql As String
    selSql = "SELECT " & SQL1 & " FROM 全コマンドリスト WHERE 重要度 ='" & Importance & "' and PRO ='" & Version & "' and MECH ='○' ORDER BY 'カテゴリ','担当者';"
    '[カテゴリ]を基準に抽出

    Set adoRs = adoCon.Execute(selSql)
    Sheets("抽出用").ListObjects(1).Range.Cells(2, 1).CopyFromRecordset adoRs


End Sub



Sub 各列名の列番号を取得()

    Worksheets("抽出用").Activate

    Call テーブル情報取得

    CategoryClm = Worksheets("抽出用").Range("" & tblName & "[カテゴリ]").Column
    PersonClm = Worksheets("抽出用").Range("" & tblName & "[担当者]").Column
    CommandClm = Worksheets("抽出用").Range("" & tblName & "[コマンド]").Column
    OptionClm = Worksheets("抽出用").Range("" & tblName & "[オプション数]").Column
    DialogClm = Worksheets("抽出用").Range("" & tblName & "[ダイアログ数]").Column



    Worksheets("進捗管理").Activate


'    CategoryClm = Worksheets("抽出用").Range("抽出用テーブル[カテゴリ]").Column
'    PersonClm = Worksheets("抽出用").Range("抽出用テーブル[担当者]").Column
'    CommandClm = Worksheets("抽出用").Range("抽出用テーブル[コマンド]").Column
'    OptionClm = Worksheets("抽出用").Range("抽出用テーブル[オプション数]").Column
'    DialogClm = Worksheets("抽出用").Range("抽出用テーブル[ダイアログ数]").Column

End Sub



Sub カテゴリ別に重複しないようにデータを貼り付け()
    Call 抽出用LO情報取得
    Call 各列名の列番号を取得

    With Worksheets("抽出用")
    Dim i, j As Integer
    Dim c, o, d As Integer
    'c：コマンド数、o：オプション数、d：ダイアログ数
    j = 2
    c = 1
    o = 0
    d = 0
        For i = fstRow + 1 To lstRow
            
            If .Cells(i, CategoryClm).Value = .Cells(i + 1, CategoryClm).Value And _
                .Cells(i, PersonClm).Value = .Cells(i + 1, PersonClm).Value Then
                 c = c + 1
            End If
            
                o = o + .Cells(i, OptionClm).Value
                d = d + .Cells(i, DialogClm).Value
            
            If .Cells(i, CategoryClm).Value <> .Cells(i + 1, CategoryClm) Or .Cells(i, PersonClm).Value <> .Cells(i + 1, PersonClm).Value Then
                Worksheets("進捗管理").ListObjects("進捗管理").Range.Cells(j, 1).Value = .Cells(i, CategoryClm).Value
                Worksheets("進捗管理").ListObjects("進捗管理").Range.Cells(j, 2).Value = .Cells(i, PersonClm).Value
                Worksheets("進捗管理").ListObjects("進捗管理").Range.Cells(j, 3).Value = c
                Worksheets("進捗管理").ListObjects("進捗管理").Range.Cells(j, 4).Value = o
                Worksheets("進捗管理").ListObjects("進捗管理").Range.Cells(j, 5).Value = d
                Worksheets("進捗管理").ListObjects("進捗管理").Range.Cells(j, 6).Value = "=SUM(D" & j + 5 & ":E" & j + 5 & ")"
                
                j = j + 1
                c = 1
                o = 0
                d = 0

            End If
        Next
    End With

End Sub



Sub 担当者別に重複しないようにデータを貼り付け()
    Call 抽出用LO情報取得
    Call 各列名の列番号を取得

    
    With Worksheets("抽出用")
    Dim Dic, i As Long, buf As String, Keys
    Set Dic = CreateObject("Scripting.Dictionary")
    For i = fstRow + 1 To lstRow
        buf = .Cells(i, PersonClm).Value
        If Not Dic.Exists(buf) Then
            Dic.Add buf, buf
        End If
    Next i
    
    '出力
    Keys = Dic.Keys
    Dim j, c, o, d As Integer
    'c：コマンド数、o：オプション数、d：ダイアログ数
    c = 0
    o = 0
    d = 0
    
    For j = 0 To Dic.Count - 1
        Worksheets("進捗管理").Cells(j + 37, 1) = Keys(j)
        Worksheets("進捗管理").Cells(j + 37, 2) = Importance
        
        For i = fstRow + 1 To lstRow
            If .Cells(i, PersonClm).Value = Keys(j) Then
                c = c + 1
                o = o + .Cells(i, OptionClm).Value
                d = d + .Cells(i, DialogClm).Value
            End If
        Next
  
        Worksheets("進捗管理").Cells(j + 37, 3) = c
        Worksheets("進捗管理").Cells(j + 37, 4) = o
        Worksheets("進捗管理").Cells(j + 37, 5) = d
        Worksheets("進捗管理").Cells(j + 37, 6) = "=SUM(D" & j + 37 & ":E" & j + 37 & ")"
        
    c = 0
    o = 0
    d = 0

    Next j
 
    Set Dic = Nothing
    End With
End Sub

Sub 空行を非表示()
    Dim ListLast            '"カテゴリ集計"テーブルの最終行番号
    ListLast = Worksheets("進捗管理").ListObjects("進捗管理").ListRows.Count + 6

    Worksheets("進捗管理").Rows("" & ListLast + 4 & ":35").EntireRow.Hidden = True
    
End Sub

Sub 空行を再表示()
    Worksheets("進捗管理").Rows.Hidden = False
  
End Sub

Sub 集計行の追加()
    Dim ListLast1           '"カテゴリ集計"テーブルの最終行番号
    ListLast1 = Worksheets("進捗管理").ListObjects("進捗管理").ListRows.Count + 6

    Dim ListLast2           '"担当者集計"テーブルの最終行番号
    ListLast2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count + 36

    Worksheets("進捗管理").Cells(ListLast1 + 1, 1).Value = "合計"
    Worksheets("進捗管理").Cells(ListLast1 + 1, 3).Value = "=SUM(C7:C" & ListLast1 & ")"
    Worksheets("進捗管理").Cells(ListLast1 + 1, 4).Value = "=SUM(D7:D" & ListLast1 & ")"
    Worksheets("進捗管理").Cells(ListLast1 + 1, 5).Value = "=SUM(E7:E" & ListLast1 & ")"
    Worksheets("進捗管理").Cells(ListLast1 + 1, 6).Value = "=SUM(F7:F" & ListLast1 & ")"
    Worksheets("進捗管理").Range("A" & ListLast1 + 1 & ":F" & ListLast1 + 1).Interior.Color = 5296274       '集計行に色をつける

    Worksheets("進捗管理").Cells(ListLast2 + 1, 1).Value = "合計"
    Worksheets("進捗管理").Cells(ListLast2 + 1, 3).Value = "=SUM(C51:C" & ListLast2 & ")"
    Worksheets("進捗管理").Cells(ListLast2 + 1, 4).Value = "=SUM(D51:D" & ListLast2 & ")"
    Worksheets("進捗管理").Cells(ListLast2 + 1, 5).Value = "=SUM(E51:E" & ListLast2 & ")"
    Worksheets("進捗管理").Cells(ListLast2 + 1, 6).Value = "=SUM(D" & ListLast2 + 1 & ":E" & ListLast2 + 1 & ")"
    Worksheets("進捗管理").Range("A" & ListLast2 + 1 & ":F" & ListLast2 + 1).Interior.Color = 49407         '集計行に色をつける


End Sub


Sub 担当変更ボタン()
'「担当変更」ボタン

'
'    Worksheets("進捗管理").Range("進捗管理").Clear
'
'    Worksheets("進捗管理").Range("担当者別").Clear

    Range("進捗管理[[カテゴリ]:[項目合計]]").ClearContents
    Range("進捗管理[[項目数1]:[進捗]]").ClearContents



    Call 抽出用LOに貼り付け
    Call カテゴリ別に重複しないようにデータを貼り付け
    Call 担当者別に重複しないようにデータを貼り付け
'    Call 集計行の追加
'    Call 空行を再表示
'    Call 空行を非表示

    Call 表をすべて表示
    Call 表の表示調整

'    Call 更新日取得
    Call データベース切断



End Sub
