# 【Excel VBA】すべてのシートに倍率100%、アクティブセルをA1に移動するマクロ

```vba
Sub Zoom100CursorA1()
   '#「s」という変数を「オブジェクト型」で定義
   Dim s As Object
   
   'フォントの統一
   With  s.Cells.Font
       .Name="MSゴシック"
       .Size=10
   End With

   '# 現在開いているブックのすべてのシートに対して順番に処理
   For Each s In ActiveWorkbook.Sheets
       s.Activate
       '# A1にカーソルを合わせる
       ActiveSheet.Range("A1").Select
       '# 倍率を「100%」にする
       ActiveWindow.Zoom = 100
       '# エクセルのマス目を非表示にする
       ActiveWindow.DisplayGridlines=False
   '# 次のシートに対して処理する
   Next s

   '図形を最後までループする
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        '図形を選択
        shp.Select Replace:=False
        shp.LockAspectRatio = msoTrue
        shp.Height= 300
    Next

   '# 一番左のシートを選択する
   Sheets(1).Select
End Sub
```

# 枠線をエクセルのショートカットで自動描画
```vba
Sub 図形の枠線書式を設定する()
   On Error GoTo ERR
   Dim Count
   Count = ActiveSheet.Shapes.Count

   Dim line_format As LineFormat

   Set line_format = Selection.ShapeRange.Line

   '複数選択
   If (Count > 1) Then
      With line_format
        .ForeColor.RGB = RGB(29,29,29)
        .Weight=1
        .Style = msoLineSingle
        .DashStyle = msoLineSolid
      End With
   End If

   '1つ選択
   If (Count =1) Then
      With line_format
        .ForeColor.RGB = RGB(255,0,0)
        .Weight=1
        .Style = msoLineSingle
        .DashStyle = msoLineSolid
      End With
   End If


Exit Sub

ERR:
     MsgBox  "オートシェイプが選択されていません"

End Sub

```

# 【Excel VBA】メール作成マクロ
```vba
Sub 日報メール作成()
'レポート部分の生成
Dim report As String: report = ""
Dim i As Long: i = 2
With Sheet2
    Do While .Cells(i, 1).Value <> ""
        report = report & .Cells(i, 1).Value & "／"
        report = report & .Cells(i, 2).Value & "／"
        report = report & .Cells(i, 3).Value & "<br>"
        i = i + 1
    Loop
End With
'メールの各要素の生成
With Sheet1
    Dim myTo As String: myTo = .Range("B2").Value
    Dim mySubject As String: mySubject = .Range("B3").Value
    Dim myBody As String: myBody = ""
    myBody = myBody & .Range("B4").Value & "<br>" '宛名"
    myBody = myBody & Range("B5").Value & "<br>" '書き出し
    myBody = myBody & report 'レポート
    myBody = myBody & Range("B6").Value '締め
End With
'下書き作成
Dim olApp As Outlook.Application
Set olApp = New Outlook.Application
Dim myMail As MailItem
Set myMail = olApp.CreateItem(olMailItem)
With myMail
    .To = myTo
    .Subject = mySubject
    .Display
    .HTMLBody = myBody & .HTMLBody
End With
End Sub

```
[参考](https://www.atmarkit.co.jp/ait/spv/1810/02/news004.html)

# データ転記サンプル
```vba
'データ転記の例
Sub Sample3()
    '①データの最終行を取得
    Dim maxRow As Long
    maxRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '②データ範囲を一括で転記
    Range("E3:G" & maxRow).Value = Range("A3:C" & maxRow).Value
End Sub

'別シートから転記の例
Sub Sample4()
    '①シートを変数にセット
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Set ws1 = Worksheets("Sheet1")
    Set ws2 = Worksheets("Sheet2")
    
    '②シートを指定してデータを転記
    ws2.Range("A1:C4").Value = ws1.Range("A1:C4").Value
End Sub
```

# Excelマクロ・VBA 2つの文字列を比較し、違う個所の文字色を赤に変更する方法
```vba
Option Explicit

Sub StrDifEmphasis()
    Const Str1StartSetCell As String = "A2"   ' 文字列1の開始セルの設定セルを指定
    Const TargetCountSetCell As String = "B2" ' 対象行数の設定セルを指定
    
    Dim Str1StartCell As String '文字列1の開始セル
    Dim targetCount As Integer  '対象行数
    
    Str1StartCell = ActiveSheet.Range(Str1StartSetCell).Value '文字列1の開始セルを取得
    targetCount = ActiveSheet.Range(TargetCountSetCell).Value '対象行数を取得
    
    Dim rowCount As Integer ' 行数のカウンター
    
    ' 対象行走査ループ。文字列1の開始セルから終了セル（対象行数分下）までループ
    For rowCount = 1 To targetCount
        
        ' 頻繁に使用する箇所を変数化(コードを短く且つ冗長性を排除するため)
        Dim str1cell As Range ' 文字列1セル
        Dim str2cell As Range ' 文字列2セル
        Dim resultCell As Range ' 結果セル

        Dim str1 As String ' 文字列1の値
        Dim str2 As String ' 文字列2の値
        
        'セルを取得
        Set str1cell = ActiveSheet.Range(Str1StartCell).Offset(rowCount - 1, 0)
        Set str2cell = ActiveSheet.Range(Str1StartCell).Offset(rowCount - 1, 1)
        Set resultCell = ActiveSheet.Range(Str1StartCell).Offset(rowCount - 1, 2)
        
        '文字列1と2の値を取得
        str1 = str1cell.Value
        str2 = str2cell.Value

        'セルの状態を初期化。文字列セルを黒文字に、結果を空白にする
        resultCell.Value = ""
        str1cell.Font.Color = vbBlack
        str2cell.Font.Color = vbBlack

        ' 2つの文字列が異なる場合にのみ処理を行う
        If str1 <> str2 Then
            
            ' 結果セルにメッセージを設定
            resultCell.Value = "異なる文字列です"
            
            Dim maxLen As Integer ' 2つの文字列の長い方の文字数
            
            ' 2つの文字列の長い方の文字数を設定
            If Len(str1) > Len(str2) Then
                ' 文字列1の方が長いため、文字列1の文字数を設定
                maxLen = Len(str1)
            Else
                ' 文字列2の方が長いため、文字列1の文字数を設定
                ' (文字数が同じ場合もこの処理。str1で行っても同じ)
                maxLen = Len(str2)
            End If
            
            Dim charCount As Integer ' 比較用文字数カウンター
            
            ' 文字比較ループ。大きいほうの文字列の文字数だけループ
            For charCount = 1 To maxLen
                Dim char1 As String ' 文字列1から抽出した1文字
                Dim char2 As String ' 文字列2から抽出した1文字
                
                Dim isChar1Under As Boolean ' 文字列1の文字数内か否か
                Dim isChar2Under As Boolean ' 文字列2の文字数内か否か
                
                '文字列1から1文字抽出
                If charCount <= Len(str1) Then
                    'charCountが文字数内に収まっているため1文字抽出
                    char1 = Mid(str1, charCount, 1)
                    isChar1Under = True
                Else
                    '文字数内に収まっていないため空白文字とする
                    char1 = ""
                    isChar1Under = False
                End If
                
                '文字列2から1文字抽出
                If charCount <= Len(str2) Then
                    'charCountが文字数内に収まっているため1文字抽出
                    char2 = Mid(str2, charCount, 1)
                    isChar2Under = True
                Else
                    '文字数内に収まっていないため空白文字とする
                    char2 = ""
                    isChar2Under = False
                End If
                
                ' 相違している文字を赤色に変更
                If char1 <> char2 Then
                    If isChar1Under Then
                        str1cell.Characters(Start:=charCount, Length:=1).Font.Color = vbRed
                    End If
                    
                    If isChar2Under Then
                        str2cell.Characters(Start:=charCount, Length:=1).Font.Color = vbRed
                    End If
                End If
                
            Next
            
        End If
        
    Next
    
    MsgBox ("終了")
    
End Sub

```
