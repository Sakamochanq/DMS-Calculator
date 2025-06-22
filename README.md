<div align="center">
    <a href="#">
        <img src="./assets/DMS-Calculator-Logo.png" width="500px">
    </a>
    <br>
    <hr>
</div>

<br>

日本語　｜　[English](./docs/README-en.md)

<br>

デフォルトの Office Excel では 度分秒 をフル文字列として扱ってしまい、四則演算を行うことが出来ません。複数のセルを使用し、複雑な変換式を経て計算を行うことは可能ですが、非常に面倒です。そこでセル上で度分秒計算を実行できるVBAScriptを開発しました。sok関数内に度分秒を代入することで簡単に計算を行うことが出来ます。Excelでのメリットは度分秒に対してオートフィルと行えるという点です。簡単な使い方は以下のGIF画像を参照してください。 🌵

<br>

<div align="center">
    <a href="#">
        <img src="./assets/DMS-Calculator-Demo.gif" width="450px">
    </a>
</div>

<br>
<br>

セル内で以下のPrefixを入力することで、様々なsok関数を自動補完できます
```
=sok_
```

<br>

### sok って何 ?

測量 = sokuryou = sok :)

<br>

### 機能
- [x] 指定範囲での加算、減算
- [ ] 度分秒の負の値計算の簡略化
- [x] 方位角の算出
- [x] cos、sinの計算
- [x] コンパス
- [x] 文字列からDecimal型に変換
- [ ] コードのリファクタリング

<br>

### 使用方法

1.  Excelより開発タブからVBAEditorを起動します。

2.  標準モジュールとして `main.vba` を追加します。

3.  保存し、セル上で関数を使用します。  

<br>

> [!Note]  
> このスクリプトはExcelVBAで動作するように設計されており、他の環境では正しく動作しない可能性があります。
<br>

### 使用例

```python
A1 = 179°50′0″
B1 = 0°10′0″

=sok_add(A1, B1) #180°0′0″
```

<br>

```python
A1 = 180°50′0″
B1 = 0°50′0″

=sok_sub(A1, B1) #180°0′0″
```

<br>

```python
A1 = 179°30′0″
A2 = 0°10′0″
A3 = 0°20′0″

=sok_sumAll(A1:A3) #180°0′0″
```

<br>

```python
A1 = 180°0′0″

=sok_compass(A1) #SE
```

<br>
<hr>

### 作成者 .他

- [Sakamochanq](https://github.com/Sakamochanq) による開発

- [Github Copilot](https://github.com/features/copilot) による開発支援

- [DeepL](https://www.deepl.com/) によるドキュメントの翻訳

<br>

### License

Release under the [MIT License](https://github.com/Sakamochanq/DMS-Calculator/blob/master/LICENSE)
