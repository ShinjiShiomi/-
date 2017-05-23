Attribute VB_Name = "Module4"

Sub 更新()
'「集計更新」ボタン

    Call オプションとダイアログの数を集計
    'DB（IJCADシート）から進捗100％のオプション数、ダイアログ数を自動取得
    
    Call 集計


End Sub

Sub 集計()

    Call 完了
    Call 工数計
    Call 残り
    Call 進捗
    Call 値更新2
    
End Sub

Sub 完了()



    Dim i As Long, RMax As Long, R(lastDay) As Range
'「途中」項目を加味し、項目数を調整する
    Dim Rsum As Long, UN(lastDay) As Range
    
        With Worksheets("進捗管理").ListObjects("進捗管理")
        
            RMax = .ListRows.Count
    
            For i = 1 To RMax

                For d = 1 To lastDay
                    Set R(d) = .ListColumns("項目数" & d & "").DataBodyRange(i)
                    Set UN(d) = .ListColumns("途中" & d & "").DataBodyRange(i)
                Next
                    
                For d = 1 To lastDay - 1
                    If UN(d) <> 0 Then
                        R(d + 1) = R(d + 1) - UN(d)
                    End If
                Next
              
              Rsum = 0
                For d = 1 To lastDay
                    Rsum = Rsum + R(d) + UN(d)
                Next
               
                If Rsum <> "0" Then
                    .ListColumns("完了").DataBodyRange(i) = Rsum
                End If
            Next i
            
        End With

End Sub

Sub 工数計()


    Dim i As Long, RMax As Long, R(lastDay) As Range, d As Integer, d2 As Integer, Rsum As Long
    
        With Worksheets("進捗管理").ListObjects("進捗管理")
        
            RMax = .ListRows.Count
    
        
             For i = 1 To RMax
                 For d = 1 To lastDay
                     Set R(d) = .ListColumns("工数" & d & "").DataBodyRange(i)
                 Next
                 
                 For d2 = 1 To lastDay
                    Rsum = Rsum + R(d2)
                 Next
                    If Rsum <> "0" Then
                        .ListColumns("工数計").DataBodyRange(i) = Rsum
                    End If
            
             Next i
           
            
        End With

End Sub


Sub 残り()


    Dim i As Long, RMax As Long, R1 As Range, R2 As Range
    
        With Worksheets("進捗管理").ListObjects("進捗管理")
        
            RMax = .ListRows.Count
    
            For i = 1 To RMax

                Set R1 = .ListColumns("項目合計").DataBodyRange(i)
                Set R2 = .ListColumns("完了").DataBodyRange(i)
               
                
                If R1 <> "" Then
                    .ListColumns("残り").DataBodyRange(i) = R1 - R2
                End If
            Next i
            
        End With

End Sub


Sub 進捗()
    Call テーブル情報取得

    Dim i As Long, RMax As Long, R1, R2
    
    For i = fstRow To lstRow
        R1 = Cells(i, Range(tblName & "[項目合計]").Column).Value
        R2 = Cells(i, Range(tblName & "[完了]").Column).Value
        
        If R1 <> 0 Then
            Cells(i, Range(tblName & "[進捗]").Column).Value = R2 / R1
        Else
            Cells(i, Range(tblName & "[進捗]").Column).Value = ""
        End If
    Next i
End Sub

Sub 進捗2()

    Dim i As Long, RMax As Long, R1 As Range, R2 As Range
        With Worksheets("進捗管理").ListObjects("進捗管理")
        
            RMax = .ListRows.Count
    
            For i = 1 To RMax
                R1 = Cells()
                
                
                
                Set R1 = .ListColumns("項目合計").DataBodyRange(i)
                Set R2 = .ListColumns("完了").DataBodyRange(i)
               
                
                If R1 <> 0 Then
                    .ListColumns("進捗").DataBodyRange(i) = R2 / R1
                End If
                
            Next i
            
        End With

End Sub





