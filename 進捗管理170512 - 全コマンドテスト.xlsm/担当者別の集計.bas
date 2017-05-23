Attribute VB_Name = "担当者別の集計"



Sub 担当者別項目数()

    Dim i As Long, RMax As Long, com As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long, d As Integer
    
    For d = 1 To lastDay
            Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
        
            For i2 = 1 To Rmax2
            
            With Worksheets("進捗管理").ListObjects("担当者別")
                Set R1 = .ListColumns("担当者").DataBodyRange(i2)
                Set R2 = .ListColumns("項目数" & d & "").DataBodyRange(i2)
            End With
                
                    With Worksheets("進捗管理").ListObjects("進捗管理")
            
                        RMax = .ListRows.Count
                          
                        For i = 1 To RMax
                            If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                            com = com + .ListColumns("項目数" & d & "").DataBodyRange(i).Value
                            End If
                        Next i
                        
                    End With
                
                If R1 <> "" Then
                    R2 = com
                End If
                
                com = 0
                
            Next i2
    Next
                
End Sub
Sub 担当者別途中()

    Dim i As Long, RMax As Long, com As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long, d As Integer
    
    For d = 1 To lastDay
        Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("担当者別")
            Set R1 = .ListColumns("担当者").DataBodyRange(i2)
            Set R2 = .ListColumns("途中" & d & "").DataBodyRange(i2)
        End With
            
                With Worksheets("進捗管理").ListObjects("進捗管理")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                        com = com + .ListColumns("途中" & d & "").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
            com = 0
            
        Next i2
    Next
                
End Sub
Sub 担当者別工数()

    Dim i As Long, RMax As Long, com As Single, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long, d As Integer
    
    For d = 1 To lastDay
        Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
    
        For i2 = 1 To Rmax2
        
        With Worksheets("進捗管理").ListObjects("担当者別")
            Set R1 = .ListColumns("担当者").DataBodyRange(i2)
            Set R2 = .ListColumns("工数" & d & "").DataBodyRange(i2)
        End With
            
                With Worksheets("進捗管理").ListObjects("進捗管理")
        
                    RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                        com = com + .ListColumns("工数" & d & "").DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
            
            If R1 <> "" Then
                R2 = com
            End If
            
            com = 0
            
        Next i2
    Next
                
End Sub

Sub 担当者別合計()

    Dim i As Long, RMax As Long, com As Long, R1 As Range, R2 As Range, i2 As Long, Rmax2 As Long, Goukei(4) As String, k As Integer
    
    Goukei(1) = "コマンド"
    Goukei(2) = "オプション"
    Goukei(3) = "ダイアログ"
    Goukei(4) = "項目合計"
    
    
    
    For k = 1 To 4
            Rmax2 = Worksheets("進捗管理").ListObjects("担当者別").ListRows.Count
        
            For i2 = 1 To Rmax2
            
                With Worksheets("進捗管理").ListObjects("担当者別")
                    Set R1 = .ListColumns("担当者").DataBodyRange(i2)
                    Set R2 = .ListColumns(Goukei(k)).DataBodyRange(i2)
                     
                End With
                    
                With Worksheets("進捗管理").ListObjects("進捗管理")
        
                RMax = .ListRows.Count
                      
                    For i = 1 To RMax
                        If .ListColumns("担当者").DataBodyRange(i) = R1 Then
                        com = com + .ListColumns(Goukei(k)).DataBodyRange(i).Value
                        End If
                    Next i
                    
                End With
                    
                    If R1 <> "" Then
                        R2 = com
                    End If
                    
                    com = 0
                
            Next i2
    Next
                
End Sub
