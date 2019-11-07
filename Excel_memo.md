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
