# 【Excel VBA】すべてのシートに倍率100%、アクティブセルをA1に移動するマクロ

```vba
Sub Zoom100CursorA1()
   '#「s」という変数を「オブジェクト型」で定義
   Dim s As Object
   '# 現在開いているブックのすべてのシートに対して順番に処理
   For Each s In ActiveWorkbook.Sheets
       s.Activate
       '# A1にカーソルを合わせる
       ActiveSheet.Range("A1").Select
       '# 倍率を「100%」にする
       ActiveWindow.Zoom = 100
   '# 次のシートに対して処理する
   Next s
   '# 一番左のシートを選択する
   Sheets(1).Select
End Sub
```

#【Excel VBA】メール作成マクロ
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
