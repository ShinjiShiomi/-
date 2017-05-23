Attribute VB_Name = "DBとの同期"
Option Explicit

Public adoCon As ADODB.Connection
Public adoRs As New ADODB.Recordset, adoRsCnt As New ADODB.Recordset
Public selSql As String, cntSql As String

Public tblName As String
Public rowCnt As Long
Public clmCnt As Long
Public hdrRow As Long
Public fstRow As Long
Public fstClm As Long
Public lstRow As Long
Public lstClm As Long

Public tblName2 As String
Public rowCnt2 As Long
Public clmCnt2 As Long
Public hdrRow2 As Long
Public fstRow2 As Long
Public fstClm2 As Long
Public lstRow2 As Long
Public lstClm2 As Long

'Public Const Importance As String = "C"
Public Const lastDay As Integer = 10
Public dayNum(lastDay)  'データの入力日数のデフォルト記入日数

Sub 指定日を取得()
    
    '指定日を取得
        Erase dayNum()
        
        Dim bgnClm
        bgnClm = 8
        Dim d
        For d = 1 To lastDay
            If Cells(5, bgnClm + (d - 1) * 3).Value <> "" Then
                dayNum(d) = DateValue(Cells(5, bgnClm + (d - 1) * 3).Value)
            End If
        Next
                  
End Sub



Sub 更新日取得()
'更新日の選択を制限
    
    Call 指定日を取得

    Dim d As Integer
    Dim strD As String
    
    d = 1
    Do While dayNum(d) <> ""
        strD = strD + "," + CStr(dayNum(d))
        d = d + 1
        
        If d = 11 Then Exit Do
        
    Loop
    
    strD = Right(strD, (Len(strD) - 1))


  With Worksheets("進捗管理").Range("F1").Validation
    .Delete
    .Add _
        Type:=xlValidateList, _
        Formula1:="" & strD & ""
  End With


End Sub



Sub テーブル情報取得()
'「進捗管理」テーブルの情報を取得
    tblName = ActiveSheet.ListObjects(1).Name
    rowCnt = ActiveSheet.ListObjects(1).ListRows.Count
    clmCnt = ActiveSheet.ListObjects(1).ListColumns.Count
    hdrRow = ActiveSheet.ListObjects(1).HeaderRowRange.Row
    fstRow = hdrRow + 1
    fstClm = ActiveSheet.ListObjects(1).Range.Cells(1, 1).Column
    lstRow = ActiveSheet.ListObjects(1).Range.Cells(rowCnt, 1).Row + 1
    lstClm = ActiveSheet.ListObjects(1).Range.Cells(1, clmCnt).Column
End Sub

Sub テーブル情報取得2()
'「担当者別」テーブルの情報を取得
    tblName2 = ActiveSheet.ListObjects(2).Name
    rowCnt2 = ActiveSheet.ListObjects(2).ListRows.Count
    clmCnt2 = ActiveSheet.ListObjects(2).ListColumns.Count
    hdrRow2 = ActiveSheet.ListObjects(2).HeaderRowRange.Row
    fstRow2 = hdrRow2 + 1
    fstClm2 = ActiveSheet.ListObjects(2).Range.Cells(1, 1).Column
    lstRow2 = ActiveSheet.ListObjects(2).Range.Cells(rowCnt2, 1).Row + 1
    lstClm2 = ActiveSheet.ListObjects(2).Range.Cells(1, clmCnt2).Column
End Sub

Sub データベース接続()
    If adoCon Is Nothing Then
        Set adoCon = New ADODB.Connection
    End If

    If adoCon.State = adStateClosed Then
        adoCon.Open "Provider=SQLOLEDB; Data Source=N-74\SQLEXPRESS; Initial Catalog=不具合リスト; Integrated Security=SSPI"

        MsgBox "データベースに接続されました"

    End If
      

'    If adoCon.State = adStateClosed Then
'        adoCon.Open "Provider=SQLOLEDB; Data Source=192.168.0.46; Initial Catalog=不具合リスト; User ID=sa; Password=smxHonda"
'    End If
End Sub

Sub データベース切断()
    If adoCon.State = adStateOpen Then
        adoCon.Close
    End If
End Sub

Sub テーブルクリア()
    Call テーブル情報取得
    Range(tblName).Clear
    ActiveSheet.ListObjects(1).Resize Range(Cells(hdrRow, fstClm), Cells(fstRow, lstClm))
End Sub

Sub 全コマンドシート再表示()

    Application.ScreenUpdating = False

    Call 選択重要度取得
    Call 選択バージョン取得

    Call テーブル情報取得2

    Dim j As Integer, n As Integer, tester() As String
    n = 0
    For j = fstRow2 To lstRow2
        If Cells(j, Range(tblName2 & "[担当者]").Column).Value <> "" Then
            ReDim Preserve tester(n)
            tester(n) = Cells(j, Range(tblName2 & "[担当者]").Column).Value
            n = n + 1
        End If
    Next

    'SQL文を作成
    Dim selSql As String

    If UBound(tester) = 0 Then
        selSql = "報告者 = '" & tester(j) & "'"
    Else
        For j = 0 To UBound(tester)
            selSql = selSql & " OR 報告者 = '" & tester(j) & "'"
        Next

        selSql = Mid(selSql, 5)
    End If

    selSql = "SELECT * FROM V_全コマンドリスト WHERE PRO = '" & Version & "' AND 重要度 = '" & Importance & "' AND (" & selSql & ")"

    MsgBox selSql

    Worksheets("全コマンド").Activate

    Call データベース接続
    Call テーブルクリア
    Call テーブル情報取得

    'データベースから情報を取得
    Dim adoRs As ADODB.Recordset
    Set adoRs = adoCon.Execute(selSql)

    Dim i As Integer
    n = 0
    For i = fstClm To lstClm
        Cells(hdrRow, i).Value = adoRs.Fields(n).Name
        n = n + 1
    Next

    'データ貼り付け
    ActiveSheet.Cells(fstRow, fstClm).CopyFromRecordset adoRs

    Worksheets("進捗管理").Activate

    Application.ScreenUpdating = True
End Sub


    
Sub オプションとダイアログの数を集計()


    Application.ScreenUpdating = False

    Call 選択重要度取得
    Call データベース接続
    Call テーブル情報取得

    Call 指定日を取得

    'シートの情報を取得
    Dim komokuClm(lastDay) As Integer, kosuClm(lastDay) As Integer
    Dim catRow As Integer
    Dim i As Integer, n As Integer, d As Integer

    '指定日の項目数列と工数列を把握
    For d = 1 To lastDay

        If dayNum(d) = "" Then


        Else
        '指定日がない場合は次へ

            For i = fstClm To lstClm
                If Cells(hdrRow - 1, i).Value <> "" Then
                    If Cells(hdrRow - 1, i).Value = dayNum(d) Then
                        komokuClm(d) = i
                        kosuClm(d) = i + 1
                    End If
                End If
            Next



        '指定日の項目数を初期化
            For i = fstRow To lstRow
                Cells(i, komokuClm(d)).Value = ""
            Next

        End If
    Next

    'データベースから情報を取得

    For d = 1 To lastDay
        selSql = "SELECT カテゴリ,オプション数,ダイアログ数,変更日 FROM V_全コマンドリスト " & _
        "WHERE 重要度 = '" & Importance & "' AND 進捗度 = 1 AND 変更日 >= '" & Format(dayNum(d), "yyyy/mm/dd") & "' AND 変更日 < '" & Format(dayNum(d) + 1, "yyyy/mm/dd") & "'"
        adoRs.Open selSql, adoCon, adOpenDynamic, adLockOptimistic

'        MsgBox selSql

    'レコードがなければメソッド終了
'    If adoRs.EOF Then
'        Exit Sub
'    Else
'        adoRs.MoveFirst
'    End If

        cntSql = "SELECT COUNT(*) AS Count FROM V_全コマンドリスト " & _
        "WHERE 重要度 = '" & Importance & "' AND 進捗度 = 1 AND 変更日 >= '" & Format(dayNum(d), "yyyy/mm/dd") & "' AND 変更日 < '" & Format(dayNum(d) + 1, "yyyy/mm/dd") & "'"
        Set adoRsCnt = adoCon.Execute(cntSql)


        For i = 1 To adoRsCnt!Count
            If DateValue(adoRs.Fields("変更日").Value) = dayNum(d) Then
                'カテゴリの行を把握
                Dim j As Integer
                For j = fstRow To lstRow
                    If Cells(j, Range(tblName & "[カテゴリ]").Column).Value = adoRs.Fields("カテゴリ").Value Then
                        catRow = j
                        Exit For
                    End If
                Next

                Cells(catRow, komokuClm(d)).Value = Cells(catRow, komokuClm(d)).Value + adoRs.Fields("オプション数").Value + adoRs.Fields("ダイアログ数").Value
            End If

            adoRs.MoveNext
        Next
        adoRs.Close
    Next

    Set adoRs = Nothing
    Set adoRsCnt = Nothing

    Application.ScreenUpdating = True
End Sub


