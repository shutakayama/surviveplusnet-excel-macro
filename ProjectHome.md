Excel Macro. for Japanese Only.

Excelマクロ集です。Microsoft Office Excel アドイン (.xla)形式で、Excel 2007 等で使用できます。

匿名メソッド for VBA
```
IF Func("cell", "Return cell.Text=""Hello""").Run( Range("A1") ) Then MsgBox "Hello"
If Func("a,b", "a=a*10:Return a+b").Run(1,1) = 11 Then MsgBox "OK"
```

匿名クラス for VBA
```
Dim o As Variant

Set o = NewWith("a", 1, "b", NewWith("b1", 100))
MsgBox o.b.b1

Set o = NewWith("c", Range("A1"))
o.c.Formula = 100
MsgBox o.c.Text
```

LINQ for VBA 実装中（※From, Where, ForEach を実装済み）
```
Call From(Range("A1", "A6")) _
    .Where(Func("cell", "Return cell.Text=""1""")) _
    .ForEach(Action("cell", "cell.Formula=30"))
```


This is open-source software licensed under the MIT License.

これはMITライセンスの オープンソース ソフトウェアです。 あなたはこれを Excel のアドイン用フォルダ※にコピーして使用することができます。 また、ソースコードをそれぞれブックにインポートして使用することが出来ます。

※Windows Vista の例
```
C:\Users\(Your Name)\AppData\Roaming\Microsoft\AddIns
```

Copyright© 2008~2010 SHIN-ICHI All Rights Reserverd. ( http://surviveplus.net )


---

一部のマクロを使用するには、他のライブラリがインストールされている必要があります。

Surviveplus.net Libraries for Macro
http://code.google.com/p/surviveplusnet-libraries-for-macro/