
ppt2pdf
=======

ppt2pdfはPowerPointのプレゼンテーション資料をPDFに変換するコマンドライン・ツールです。

[PowerPoint プレゼンテーションを PDF ファイルで保存する](https://support.office.com/ja-jp/article/powerpoint-%E3%83%97%E3%83%AC%E3%82%BC%E3%83%B3%E3%83%86%E3%83%BC%E3%82%B7%E3%83%A7%E3%83%B3%E3%82%92-pdf-%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%81%A7%E4%BF%9D%E5%AD%98%E3%81%99%E3%82%8B-9b5c786b-9c6e-4fe6-81f6-9372f77c47c8)のはPowerPointのファイルが1つだけなら簡単ですが、ファイルの数が多いと大変です。

ppt2pdfはこの問題を解決します。ppt2pdfはPowerPointで指定されたファイルを開き、PDFとしてエクスポートします。PowerPointの機能を使っているため、出力されたPDFのレイアウトが崩れたりしません。

使い方
------

ppt2pdfコマンドに引数としてPowerPointファイルを指定します。

```console
C:\ppt2pdf>ppt2pdf sample.pptx sample2.pptx
ppt2pdf v0.1 by Shinichi Akiyama
PowoerPoint version 16.0
C:\ppt2pdf\sample.pptx was opened.
C:\ppt2pdf\sample.pdf was created.
C:\ppt2pdf\sample2.pptx was opened.
C:\ppt2pdf\sample2.pdf was created.
All files were converted.

C:\ppt2pdf>
```

ワイルドカードも指定可能です。

```console
C:\ppt2pdf>ppt2pdf *.pptx
ppt2pdf v0.1 by Shinichi Akiyama
PowoerPoint version 16.0
C:\ppt2pdf\sample.pptx was opened.
C:\ppt2pdf\sample.pdf was created.
C:\ppt2pdf\sample2.pptx was opened.
C:\ppt2pdf\sample2.pdf was created.
All files were converted.

C:\ppt2pdf>
```

Author
------

[Shinichi Akiyama](https://github.com/shakiyam)

License
-------

[MIT License](https://opensource.org/licenses/mit)
