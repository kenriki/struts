# VBSからExcelマクロを実行させる

> これにはVBSを使用します。VBS (Visual Basic Scripting Edition)はWindowsに標準装備されているスクリプト言語で、VBAとほぼ同じ構文でプログラミングでき、各Officeアプリをオブジェクトで呼べます。

>「もうこれボタン押すだけだし、オレがやる必要なくね？」状態までVBAができるあなたであれば、簡単に使いこなせます。

```vbs

Const WB_PATH = "C:\Users\Ore\Desktop\test.xlsm"
Const PROC_NAME = "main"
Dim excelApp
Set excelApp = CreateObject("Excel.Application")

With excelApp
	.Visible = False
	Dim wb
	Set wb = .Workbooks.Open(WB_PATH)
	.Run wb.Name & "!" & PROC_NAME
	.DisplayAlerts = False
	wb.Save
	wb.Close
End With

excelApp.Quit

```

> このように見た感じはほぼVBAですが、注意点として変数宣言に型を指定してはいけません。エラーで止まります。

> WB_PATH　はExcelファイルのパスです。
> PROC_NAME　は実行するプロシージャの名前です。
> あなたの環境に合わせて変更してください。

> やっていることは「もうこれボタン押すだけだし～」のあなたなら説明の必要はないと思いますが、キモはExcel.ApplicationオブジェクトのRunメソッドによるプロシージャ実行です。これでイベントに頼らずに、このスクリプトが実行されると指定したファイルの指定した名前のプロシージャが実行されます。

> このテキストファイルを拡張子.vbsで保存します。

> では、本来とは逆の手順になりますが、テストのため上のvbsファイルで指定しているパスに実際にExcelファイルを置き、指定したプロシージャ名でマクロをこしらえてみます。

> 次のようなプロシージャを作りました。

```vba
Sub main()
    Range("a1") = "東京都"
End Sub
```

> vbsファイルをダブルクリックで実行します。
> 実行しても特に何もリアクションがありませんが、Excelファイルを開いてみると

> ちゃんとマクロが実行されています。

> あとはvbsファイルをタスクスケジューラへ登録すれば目的が達成されます。

> この要領で毎日、毎週、毎月やらなければならない作業を完全自動でコンピューターにやらせることができます。

