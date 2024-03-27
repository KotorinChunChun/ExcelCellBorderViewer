# ExcelCellBorderViewer
Excelのセル付き罫線がどっちなのかリアルタイムに確認するためのExcelアドインです。

![image](https://github.com/KotorinChunChun/ExcelCellBorderViewer/assets/55196383/ca004067-72ac-4907-86e8-3439f6c4b3b0)


![セル付き罫線ビューアー](https://github.com/KotorinChunChun/ExcelCellBorderViewer/assets/55196383/cb5fc158-08ff-4beb-9dbb-ca11924e4484)



## 注意

このリポジトリは、似たような機能を実装したい開発者への参考資料、VBAの学習用、そしてネタツールとしてお楽しみ頂くために公開しています。

完成されたツールとして使っていただくことは想定していません。

本プログラムを使用したことによる一切の不利益を作者は補償できません。



## ダウンロード

binの中に入っているxlamがソースコードを含むアドイン本体です。

[ここをクリック](https://github.com/KotorinChunChun/ExcelCellBorderViewer/raw/main/bin/%E3%82%BB%E3%83%AB%E4%BB%98%E3%81%8D%E7%BD%AB%E7%B7%9A%E3%83%93%E3%83%A5%E3%83%BC%E3%82%A2%E3%83%BC.xlam)するとダウンロードできます。



## 機能

- 選択したセルの周辺のセルに付与された罫線の確認
- クリックした罫線の追加（※VBAで変更を加えるため元に戻すが使えなくなります）



## 改善すべき点について
- セル範囲に対応していません。フォーカスのあたっている1セルのみです。セル結合も考慮していません。
- 斜め線に対応していません。
- 罫線の色に対応していません。なんでも黒です。
- 罫線の種類に対応していません。なんでも実線です。



## 工夫とギミックについて

### イベントの検知をグローバルに。フォームモジュール単独で実現

CellBorderViewFormクラスでは、 `Private WithEvents app As Excel.Application` を宣言し、コンストラクタ`UserForm_Initialize` で `Set app = Application` してエクセルアプリケーション全体のイベントをフックの対象としています。

WidthEventsはクラスモジュールで実施するのが一般的ですが、フォームモジュールもクラスモジュールの一つなので、このように実装できます。


### フォームのインスタンスをプロシージャに内包

WithEventsによるイベントフックを常駐させるには、クラス（フォーム）の変数をグローバル変数などで保持するのが一般的ですが、起動用プロシージャ内でStatic宣言することによりグローバルに保持しています。

また、フォームの多重起動防止とフォームを閉じられた場合の再生成も、1つのプロシージャのみで実現しています。

```
Sub Startセル付き罫線ビューアー()
    Static fm As CellBorderViewForm
    
    On Error Resume Next
        Debug.Print Now, "fm.Visible : " & fm.Visible
        If Err Then Set fm = Nothing
    On Error GoTo 0
    
    If fm Is Nothing Then
        Set fm = New CellBorderViewForm
        fm.Show False
    End If
End Sub
```

### コントロールの動的生成と配列化

本プログラムでは、フォームのデザイナを一切使用せず、全てのコントロール・プロパティの設定をソースコード上から行っています。

そのため、本リポジトリのモジュールをインポートするのは必須ではありません。

自分で空のフォームを作成し、ソースコードだけコピペすると動きます。

### セル情報のクラス化による効率化

今回は1つのセルに対する `VirtualCell` クラスを定義し、テキストボックスと上下左右の罫線を管理するコントロールをひとまとめに管理しています。

コントロールとセルを1:1と見なし、VirtualCellを配列として生成しています。

これにより、`CellBorderViewForm` のソースコードはかなりスッキリしたと思います。

上下左右を配列化して XlBordersIndex と対応付けると短くできるかもしれませんが、今回は可読性を優先しました。

### セルの書式設定をXMLで取得することによる安定化

一般的な罫線を特定する方法 `Range.Borders(・・・).LineStyle` では、上のセル付きの下線か、下のセル付きの上線か識別することができません。

当初、このプロジェクトでは、検査したいセルを1件づつ未使用のセルに `Range.Copy Destination:Range` してから確認を行っており、致命的な問題が含まれていました。

新たに `Range.Value(xlRangeValueXMLSpreadsheet)` を採用したことで書式設定をXML形式で取得し、クリップボードへ頼らずに安定してセルに付与された罫線が特定できるようになりました。



## 謝辞

- [筒井.xls様](https://twitter.com/Tsutsui0524)
    - 話題を提供していただきました。
- [furyutei様](https://twitter.com/furyutei)
    - VirtualCell.cls.vbaのGetBorderInfo関数の実装をコピーさせていただきました。
