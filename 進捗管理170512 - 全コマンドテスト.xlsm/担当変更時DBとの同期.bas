Attribute VB_Name = "担当変更時DBとの同期"
Sub カテゴリ選択()
'カテゴリ列の選択を制限

  With Worksheets("進捗管理").ListObjects("進捗管理").ListColumns("カテゴリ").DataBodyRange.Validation
    .Delete
    .Add _
        Type:=xlValidateList, _
        Formula1:="作成,修正,ファイル,寸法,挿入,カスタマイズ,ツール,形式,表示,編集,ブロック,作図補助,ウィンドウ,コントロール,設定,拡張ツール,PLUSツール,ヘルプ,システム変数,例外,その他,なし"
  End With

End Sub

Sub 担当者選択()
'担当者列の選択を制限

  With Worksheets("進捗管理").ListObjects("進捗管理").ListColumns("担当者").DataBodyRange.Validation
    .Delete
    .Add _
        Type:=xlValidateList, _
        Formula1:="戸家,野田,山下,大平,塩見"
  End With

    
  With Worksheets("進捗管理").ListObjects("担当者別").ListColumns("担当者").DataBodyRange.Validation
    .Delete
    .Add _
        Type:=xlValidateList, _
        Formula1:="戸家,野田,山下,大平,塩見"
  End With


End Sub

Sub 重要度貼り付け()

'
'  With Worksheets("進捗管理").ListObjects("担当者別").ListColumns("重要度").DataBodyRange.Validation
'    .Delete
'    .Add _
'        Type:=xlValidateList, _
'        Formula1:="A,B,C"
'  End With

    Call 選択重要度取得

Dim RMax As Long, i As Integer

    With Worksheets("進捗管理").ListObjects("担当者別")
        RMax = .ListRows.Count
        
        For i = 1 To RMax
            .ListColumns("重要度").DataBodyRange(i) = Importance
        Next
        

    End With

End Sub

Sub 表をすべて表示()
    Rows.Hidden = False
    Columns.Hidden = False
    
End Sub

Sub 表の表示調整()
'コマンド数が０か空白の場合、非表示にする
    
    Call テーブル情報取得
    Call テーブル情報取得2
    
Dim i, j As Integer
    For i = fstRow To lstRow
        If Cells(i, Range(tblName & "[コマンド]").Column).Value = "" Or Cells(i, Range(tblName & "[コマンド]").Column).Value = 0 Then
             Rows(i).Hidden = True
        End If
    Next
    
    For j = fstRow2 To lstRow2
        If Cells(j, Range(tblName2 & "[担当者]").Column).Value = "" Then
             Rows(j).Hidden = True
        End If
    Next

End Sub


Sub 担当更新()
Attribute 担当更新.VB_ProcData.VB_Invoke_Func = " \n14"

    
    Call 全コマンドシート再表示

    Call コマンド数
    Call オプション
    Call ダイアログ
    Call 項目合計
    Call 担当者別合計
       
    Call 重要度貼り付け
    Call 表の表示調整
   
End Sub


Sub コマンド数()
    
    Dim i As Long, RMax As Long, cnt As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long
        Rmax2 = Worksheets("進捗管理").ListObjects("進捗管理").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("進捗管理")
            Set R1 = .ListColumns("カテゴリ").DataBodyRange(i2)
            Set R2 = .ListColumns("コマンド").DataBodyRange(i2)
        End With
            
                With Worksheets("全コマンド").ListObjects(1)
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("カテゴリ").DataBodyRange(i) = R1 And .ListColumns("PRO").DataBodyRange(i) <> "×" Then
                        cnt = cnt + 1
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = cnt
            End If
            
            cnt = 0
            
        Next i2
                
End Sub

Sub オプション()

    Dim i As Long, RMax As Long, com As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long
        Rmax2 = Worksheets("進捗管理").ListObjects("進捗管理").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("進捗管理")
            Set R1 = .ListColumns("カテゴリ").DataBodyRange(i2)
            Set R2 = .ListColumns("オプション").DataBodyRange(i2)
        End With
            
                With Worksheets("全コマンド").ListObjects("全コマンド")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("カテゴリ").DataBodyRange(i) = R1 And .ListColumns("PRO").DataBodyRange(i) <> "×" Then
                        com = com + .ListColumns("オプション数").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
            com = 0
            
        Next i2
                
End Sub

Sub ダイアログ()

    Dim i As Long, RMax As Long, com As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long
        Rmax2 = Worksheets("進捗管理").ListObjects("進捗管理").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("進捗管理")
            Set R1 = .ListColumns("カテゴリ").DataBodyRange(i2)
            Set R2 = .ListColumns("ダイアログ").DataBodyRange(i2)
        End With
            
                With Worksheets("全コマンド").ListObjects("全コマンド")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("カテゴリ").DataBodyRange(i) = R1 And .ListColumns("PRO").DataBodyRange(i) <> "×" Then
                        com = com + .ListColumns("ダイアログ数").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
            com = 0
            
        Next i2
                
End Sub

Sub 項目合計()


    Dim i As Long, RMax As Long, R1 As Range, R2 As Range
    
        With Worksheets("進捗管理").ListObjects("進捗管理")
        
            RMax = .ListRows.Count
    
            For i = 1 To RMax

                Set R1 = .ListColumns("オプション").DataBodyRange(i)
                Set R2 = .ListColumns("ダイアログ").DataBodyRange(i)
                
                If R1 <> "" Then
                    .ListColumns("項目合計").DataBodyRange(i) = R1 + R2
                End If
            
            Next i
            
        End With

End Sub

Sub クリアボタン()



    Range("進捗管理[[カテゴリ]:[項目合計]]").ClearContents
    Range("進捗管理[[項目数1]:[進捗]]").ClearContents
    Range("担当者別[[担当者]:[進捗]]").ClearContents
    Call 集計

    Call 表をすべて表示

End Sub

