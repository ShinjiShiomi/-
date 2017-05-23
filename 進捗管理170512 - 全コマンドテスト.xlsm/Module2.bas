Attribute VB_Name = "Module2"
Sub 値更新2()

    Call 担当者コマンド数
    Call 担当者オプション
    Call 担当者ダイアログ
    Call 担当者項目合計
    Call 担当者予定工数
    Call 項目数工数反映
    Call 担当者完了
    Call 担当者工数計
    Call 担当者残り
    Call 担当者進捗

End Sub
Sub 担当者コマンド数()

    Dim i As Long, RMax As Long, com As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long
        Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("担当者別")
            Set R1 = .ListColumns("担当者").DataBodyRange(i2)
            Set R2 = .ListColumns("コマンド").DataBodyRange(i2)
        End With
            
                With Worksheets("進捗管理").ListObjects("進捗管理")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                        com = com + .ListColumns("コマンド").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
            com = 0
            
        Next i2
                
End Sub
Sub 担当者オプション()

    Dim i As Long, RMax As Long, com As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long
        Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("担当者別")
            Set R1 = .ListColumns("担当者").DataBodyRange(i2)
            Set R2 = .ListColumns("オプション").DataBodyRange(i2)
        End With
            
                With Worksheets("進捗管理").ListObjects("進捗管理")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                        com = com + .ListColumns("オプション").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
            com = 0
            
        Next i2
                
End Sub

Sub 担当者ダイアログ()

    Dim i As Long, RMax As Long, com As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long
        Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("担当者別")
            Set R1 = .ListColumns("担当者").DataBodyRange(i2)
            Set R2 = .ListColumns("ダイアログ").DataBodyRange(i2)
        End With
            
                With Worksheets("進捗管理").ListObjects("進捗管理")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                        com = com + .ListColumns("ダイアログ").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
            com = 0
            
        Next i2
                
End Sub

Sub 担当者項目合計()

    Dim i As Long, RMax As Long, com As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long
        Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("担当者別")
            Set R1 = .ListColumns("担当者").DataBodyRange(i2)
            Set R2 = .ListColumns("項目合計").DataBodyRange(i2)
        End With
            
                With Worksheets("進捗管理").ListObjects("進捗管理")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                        com = com + .ListColumns("項目合計").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
            com = 0
            
        Next i2
                
End Sub
Sub 担当者予定工数()

    Dim i As Long, RMax As Long, com As Single, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long
        Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        com = 0
        With Worksheets("進捗管理").ListObjects("担当者別")
            Set R1 = .ListColumns("担当者").DataBodyRange(i2)
            Set R2 = .ListColumns("予定工数").DataBodyRange(i2)
        End With
                       
                With Worksheets("進捗管理").ListObjects("進捗管理")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax

                        If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                        com = com + .ListColumns("予定工数").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
                        
        Next i2
                
End Sub
Sub 項目数工数反映()

    Call 担当者別項目数
    Call 担当者別途中
    Call 担当者別工数

End Sub

Sub 担当者完了()

    Dim i As Long, RMax As Long, com As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long
        Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("担当者別")
            Set R1 = .ListColumns("担当者").DataBodyRange(i2)
            Set R2 = .ListColumns("完了").DataBodyRange(i2)
        End With
            
                With Worksheets("進捗管理").ListObjects("進捗管理")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                        com = com + .ListColumns("完了").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
            com = 0
            
        Next i2
                
End Sub

Sub 担当者工数計()

    Dim i As Long, RMax As Long, com As Single, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long
        Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("担当者別")
            Set R1 = .ListColumns("担当者").DataBodyRange(i2)
            Set R2 = .ListColumns("工数計").DataBodyRange(i2)
        End With
            
                With Worksheets("進捗管理").ListObjects("進捗管理")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                        com = com + .ListColumns("工数計").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
            com = 0
            
        Next i2
                
End Sub

Sub 担当者残り()

    Dim i As Long, RMax As Long, com As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long
        Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("担当者別")
            Set R1 = .ListColumns("担当者").DataBodyRange(i2)
            Set R2 = .ListColumns("残り").DataBodyRange(i2)
        End With
            
                With Worksheets("進捗管理").ListObjects("進捗管理")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                        com = com + .ListColumns("残り").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
            com = 0
            
        Next i2
                
End Sub

Sub 担当者進捗()
    Call テーブル情報取得2

    Dim i As Long, RMax As Long, R1, R2
    
    For i = fstRow2 To lstRow2
        R1 = Cells(i, Range(tblName2 & "[項目合計]").Column).Value
        R2 = Cells(i, Range(tblName2 & "[完了]").Column).Value
        
        If R1 <> 0 Then
            Cells(i, Range(tblName2 & "[進捗]").Column).Value = R2 / R1
        Else
            Cells(i, Range(tblName2 & "[進捗]").Column).Value = ""
        End If
    Next i
End Sub

Sub 担当者進捗2()


    Dim i As Long, RMax As Long, R1 As Range, R2 As Range
    
        With Worksheets("進捗管理").ListObjects("担当者別")
        
            RMax = .ListRows.Count
    
            For i = 1 To RMax

                Set R1 = .ListColumns("項目合計").DataBodyRange(i)
                Set R2 = .ListColumns("完了").DataBodyRange(i)
               
                
                If R1 <> "0" Then
                    .ListColumns("進捗").DataBodyRange(i) = R2 / R1
                End If
            Next i
            
        End With

End Sub



