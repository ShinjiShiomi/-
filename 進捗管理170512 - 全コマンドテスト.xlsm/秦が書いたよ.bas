Attribute VB_Name = "秦が書いたよ"
'Option Explicit
'
'Dim adoCon As ADODB.Connection
'
'Public tblName As String
'Public rowCnt As Long
'Public clmCnt As Long
'Public hdrRow As Long
'Public fstRow As Long
'Public fstClm As Long
'Public lstRow As Long
'Public lstClm As Long
'
'Const juyodo As String = "C"
'
'Sub テーブル情報取得()
'    tblName = ActiveSheet.ListObjects(1).Name
'    rowCnt = ActiveSheet.ListObjects(1).ListRows.Count
'    clmCnt = ActiveSheet.ListObjects(1).ListColumns.Count
'    hdrRow = ActiveSheet.ListObjects(1).HeaderRowRange.Row
'    fstRow = hdrRow + 1
'    fstClm = ActiveSheet.ListObjects(1).Range.Cells(1, 1).Column
'    lstRow = ActiveSheet.ListObjects(1).Range.Cells(rowCnt, 1).Row + 1
'    lstClm = ActiveSheet.ListObjects(1).Range.Cells(1, clmCnt).Column
'End Sub
'
'Sub データベース接続()
'    If adoCon Is Nothing Then
'        Set adoCon = New ADODB.Connection
'    End If
'
'    If adoCon.State = adStateClosed Then
'        adoCon.Open "Provider=SQLOLEDB; Data Source=192.168.0.46; Initial Catalog=不具合リスト; User ID=sa; Password=smxHonda"
'    End If
'End Sub
'
'Sub データベース切断()
'    If adoCon.State = adStateOpen Then
'        adoCon.Close
'    End If
'End Sub
'
'Sub テーブルクリア()
'    Call テーブル情報取得
'    Range(tblName).Clear
'    ActiveSheet.ListObjects(1).Resize Range(Cells(hdrRow, fstClm), Cells(fstRow, lstClm))
'End Sub
'
'Sub 全コマンドシート再表示()
'    Application.ScreenUpdating = False
'
'    'テスター名を格納
'    Dim tblName2 As String
'    Dim rowCnt2 As Long
'    Dim clmCnt2 As Long
'    Dim hdrRow2 As Long
'    Dim fstRow2 As Long
'    Dim fstClm2 As Long
'    Dim lstRow2 As Long
'    Dim lstClm2 As Long
'
'    tblName2 = ActiveSheet.ListObjects(2).Name
'    rowCnt2 = ActiveSheet.ListObjects(2).ListRows.Count
'    clmCnt2 = ActiveSheet.ListObjects(2).ListColumns.Count
'    hdrRow2 = ActiveSheet.ListObjects(2).HeaderRowRange.Row
'    fstRow2 = hdrRow2 + 1
'    fstClm2 = ActiveSheet.ListObjects(2).Range.Cells(1, 1).Column
'    lstRow2 = ActiveSheet.ListObjects(2).Range.Cells(rowCnt2, 1).Row + 1
'    lstClm2 = ActiveSheet.ListObjects(2).Range.Cells(1, clmCnt2).Column
'
'    Dim j As Integer, n As Integer, tester() As String
'    n = 0
'    For j = fstRow2 To lstRow2
'        If Cells(j, Range(tblName2 & "[担当者]").Column).Value <> "" Then
'            ReDim Preserve tester(n)
'            tester(n) = Cells(j, Range(tblName2 & "[担当者]").Column).Value
'            n = n + 1
'        End If
'    Next
'
'    'SQL文を作成
'    Dim selSql As String
'
'    If UBound(tester) = 0 Then
'        selSql = "担当者 = '" & tester(j) & "'"
'    Else
'        For j = 0 To UBound(tester)
'            selSql = selSql & " OR 担当者 = '" & tester(j) & "'"
'        Next
'
'        selSql = Mid(selSql, 5)
'    End If
'
'    selSql = "SELECT * FROM V_全コマンドリスト WHERE 重要度 = '" & juyodo & "' AND (" & selSql & ")"
'
'    MsgBox selSql
'
'    Worksheets("全コマンド").Activate
'
'    Call データベース接続
'    Call テーブルクリア
'    Call テーブル情報取得
'
'    'データベースから情報を取得
'    Dim adoRs As ADODB.Recordset
'    Set adoRs = adoCon.Execute(selSql)
'
'    Dim i As Integer
'    n = 0
'    For i = fstClm To lstClm
'        Cells(hdrRow, i).Value = adoRs.Fields(n).Name
'        n = n + 1
'    Next
'
'    'データ貼り付け
'    ActiveSheet.Cells(fstRow, fstClm).CopyFromRecordset adoRs
'
'    Worksheets("進捗管理").Activate
'
'    Application.ScreenUpdating = True
'End Sub
'
'Sub オプションとダイアログの数を集計()
'    Application.ScreenUpdating = False
'
'    Call データベース接続
'    Call テーブル情報取得
'
'    '日にちを取得
'    Dim dayNum As Long
'    dayNum = DateValue(Range("F2").Value)
'
'    'シートの情報を取得
'    Dim komokuClm As Integer, kosuClm As Integer
'    Dim catRow As Integer
'    Dim i As Integer, n As Integer
'
'    '今日の項目数列と工数列を把握
'    For i = fstClm To lstClm
'        If Cells(hdrRow - 1, i).Value <> "" Then
'            If Cells(hdrRow - 1, i).Value = dayNum Then
'                komokuClm = i
'                kosuClm = i + 1
'            End If
'        End If
'    Next
'
'    '今日の項目数を初期化
'    For i = fstRow To lstRow
'        Cells(i, komokuClm).Value = ""
'    Next
'
'    'データベースから情報を取得
'    Dim adoRs As New ADODB.Recordset, adoRsCnt As New ADODB.Recordset
'    Dim selSql As String, cntSql As String
'    selSql = "SELECT カテゴリ,オプション数,ダイアログ数,変更日 FROM V_全コマンドリスト " & _
'    "WHERE 重要度 = '" & juyodo & "' AND 進捗度 = 1 AND 変更日 >= '" & Format(dayNum, "yyyy/mm/dd") & "' AND 変更日 <= '" & Format(dayNum + 1, "yyyy/mm/dd") & "'"
'    adoRs.Open selSql, adoCon, adOpenDynamic, adLockOptimistic
'
'    'レコードがなければメソッド終了
'    If adoRs.EOF Then
'        Exit Sub
'    Else
'        adoRs.MoveFirst
'    End If
'
'    cntSql = "SELECT COUNT(*) AS Count FROM V_全コマンドリスト " & _
'    "WHERE 重要度 = '" & juyodo & "' AND 進捗度 = 1 AND 変更日 >= '" & Format(dayNum, "yyyy/mm/dd") & "' AND 変更日 <= '" & Format(dayNum + 1, "yyyy/mm/dd") & "'"
'    Set adoRsCnt = adoCon.Execute(cntSql)
'
'    For i = 1 To adoRsCnt!Count
'        If DateValue(adoRs.Fields("変更日").Value) = dayNum Then
'            'カテゴリの行を把握
'            Dim j As Integer
'            For j = fstRow To lstRow
'                If Cells(j, Range(tblName & "[カテゴリ]").Column).Value = adoRs.Fields("カテゴリ").Value Then
'                    catRow = j
'                    Exit For
'                End If
'            Next
'
'            Cells(catRow, komokuClm).Value = Cells(catRow, komokuClm).Value + adoRs.Fields("オプション数").Value + adoRs.Fields("ダイアログ数").Value
'        End If
'
'        adoRs.MoveNext
'    Next
'
'    Set adoRs = Nothing
'    Set adoRsCnt = Nothing
'
'    Application.ScreenUpdating = True
'End Sub
