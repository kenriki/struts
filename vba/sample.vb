'期限超過チェック
Sub checkTaskDate()
    Worksheets("Sheet1").Activate
    Dim rowNum As Integer
    rowNum = 2
    
    '期限欄が空白なるまでループ
    Do Until Cells(rowNum, 5) = ""
        '今日日付よりも古い場合
        If Cells(rowNum, 5).Value < Format(Date, "yyyy/mm/dd") Then
            Cells(rowNum, 1).Interior.ColorIndex = 3
            Cells(rowNum, 2).Interior.ColorIndex = 3
            Cells(rowNum, 3).Interior.ColorIndex = 3
            Cells(rowNum, 4).Interior.ColorIndex = 3
            Cells(rowNum, 5).Interior.ColorIndex = 3
            Cells(rowNum, 6).Interior.ColorIndex = 3
            Cells(rowNum, 7).Value = "期限が遅れています"
            Cells(rowNum, 7).Font.ColorIndex = 3
        End If
        If Cells(rowNum, 3).Value = "△" Then
            Cells(rowNum, 1).Interior.ColorIndex = 7
            Cells(rowNum, 2).Interior.ColorIndex = 7
            Cells(rowNum, 3).Interior.ColorIndex = 7
            Cells(rowNum, 4).Interior.ColorIndex = 7
            Cells(rowNum, 5).Interior.ColorIndex = 7
            Cells(rowNum, 6).Interior.ColorIndex = 7
            Cells(rowNum, 7).Value = "保留中のタスクです"
            Cells(rowNum, 7).Font.ColorIndex = 13
        End If
        
        If Cells(rowNum, 3).Value = "〇" Then
            Cells(rowNum, 1).Interior.ColorIndex = 8
            Cells(rowNum, 2).Interior.ColorIndex = 8
            Cells(rowNum, 3).Interior.ColorIndex = 8
            Cells(rowNum, 4).Interior.ColorIndex = 8
            Cells(rowNum, 5).Interior.ColorIndex = 8
            Cells(rowNum, 6).Interior.ColorIndex = 8
            Cells(rowNum, 7).Value = ""
            
        End If
        
        If Cells(rowNum, 3) = "×" Then
            Cells(rowNum, 1).Interior.ColorIndex = 22
            Cells(rowNum, 2).Interior.ColorIndex = 22
            Cells(rowNum, 3).Interior.ColorIndex = 22
            Cells(rowNum, 4).Interior.ColorIndex = 22
            Cells(rowNum, 5).Interior.ColorIndex = 22
            Cells(rowNum, 6).Interior.ColorIndex = 22
            Cells(rowNum, 7).Value = "リスケ必要"
            Cells(rowNum, 7).Font.ColorIndex = 3
        End If
        '行数カウント
        rowNum = rowNum + 1
    Loop
    
    
End Sub

'タスク最新化
Sub checkvalue()

    Worksheets("Sheet1").Activate
    
    Dim rowNum As Integer
    Dim hantei_maru As Integer
    Dim hantei_batsu As Integer
    Dim hantei_pd As Integer
    Dim hantei_undigested As Integer
    Dim hantei_going As Integer
    
    hantei_maru = 0
    hantei_batsu = 0
    hantei_pd = 0
    hantei_going = 0
    hantei_undigested = 0
    rowNum = 2
    
    
    'タスク表
    Do Until Cells(rowNum, 2) = ""
    
        'Cells(rowNum, 6).Value = ""
        'Cells(rowNum, 7).Value = ""
        
        If Cells(rowNum, 3) = "〇" Then
            Cells(rowNum, 1).Interior.ColorIndex = 8
            Cells(rowNum, 2).Interior.ColorIndex = 8
            Cells(rowNum, 3).Interior.ColorIndex = 8
            Cells(rowNum, 4).Interior.ColorIndex = 8
            Cells(rowNum, 5).Interior.ColorIndex = 8
            Cells(rowNum, 6).Interior.ColorIndex = 8
            hantei_maru = hantei_maru + 1
        End If
        
        If Cells(rowNum, 3) = "△" Then
            Cells(rowNum, 1).Interior.ColorIndex = 6
            Cells(rowNum, 2).Interior.ColorIndex = 6
            Cells(rowNum, 3).Interior.ColorIndex = 6
            Cells(rowNum, 4).Interior.ColorIndex = 6
            Cells(rowNum, 5).Interior.ColorIndex = 6
            Cells(rowNum, 6).Interior.ColorIndex = 6
            hantei_pd = hantei_pd + 1
        End If
        
        If Cells(rowNum, 3) = "×" Then
            Cells(rowNum, 1).Interior.ColorIndex = 22
            Cells(rowNum, 2).Interior.ColorIndex = 22
            Cells(rowNum, 3).Interior.ColorIndex = 22
            Cells(rowNum, 4).Interior.ColorIndex = 22
            Cells(rowNum, 5).Interior.ColorIndex = 22
            Cells(rowNum, 6).Interior.ColorIndex = 22
            hantei_batsu = hantei_batsu + 1
        End If
        
        If Cells(rowNum, 3) = "★" Then
            Cells(rowNum, 1).Interior.ColorIndex = 4
            Cells(rowNum, 2).Interior.ColorIndex = 4
            Cells(rowNum, 3).Interior.ColorIndex = 4
            Cells(rowNum, 4).Interior.ColorIndex = 4
            Cells(rowNum, 5).Interior.ColorIndex = 4
            Cells(rowNum, 6).Interior.ColorIndex = 4
            hantei_going = hantei_going + 1
        End If
        
        If Cells(rowNum, 3) = "" Or Cells(rowNum, 3) = "-" Then
            Cells(rowNum, 3) = "-"
            Cells(rowNum, 1).Interior.ColorIndex = 35
            Cells(rowNum, 2).Interior.ColorIndex = 35
            Cells(rowNum, 3).Interior.ColorIndex = 35
            Cells(rowNum, 4).Interior.ColorIndex = 35
            Cells(rowNum, 5).Interior.ColorIndex = 35
            Cells(rowNum, 6).Interior.ColorIndex = 35
            hantei_undigested = hantei_undigested + 1
        End If
        
        
        rowNum = rowNum + 1
    Loop
    'MsgBox "完了した数は、「" & hantei_maru & "」です。"

    'まとめ表
    Cells(8, 13).Value = hantei_maru
    Cells(9, 13).Value = hantei_batsu
    Cells(10, 13).Value = hantei_pd
    Cells(11, 13).Value = hantei_undigested
    Cells(12, 13).Value = hantei_going
    Cells(13, 13).Value = hantei_maru + hantei_batsu + hantei_pd + hantei_going + hantei_undigested
End Sub
