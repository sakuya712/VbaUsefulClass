
目次
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [Dic Class](#dic-class)
  - [Methods](#methods)
    - [Add](#add)
    - [AddIf](#addif)
    - [Exists](#exists)
    - [Items](#items)
    - [Key](#key)
    - [Remove](#remove)
    - [SetCompareMode](#setcomparemode)
  - [Properties](#properties)
    - [Count](#count)
    - [DuplicateCollection](#duplicatecollection)
    - [Items](#items-1)
    - [Keys](#keys)

<!-- /code_chunk_output -->

# Dic Class

VBAにはDictionaryオブジェクト(連想配列)がありますが参照しないとインテリセンスが使えないというの面倒    
またDictionaryは重複しているかの判断ができるので、これを利用することで重複しないリスト、重複したものを出すリストなどを作成できます。  

## Methods

### Add

\<引数> キー、要素  
\<戻値> なし  

- 通常のDictionaryと同じようにキーと連想する要素を追加します

### AddIf

\<引数> キー、要素、[DuplicateMode列挙型]  
\<戻値> なし  

- キーが重複していないかを確認して、追加します。  
- 重複していた場合はDuplicateMode列挙型で決めた処理を行います
- 第3引数を省略した場合はdmNothingの処理が行われます
- 重複リストはDuplicateCollectionプロパティから取得できます

**DuplicateMode**
|  要素名           |  説明  |
| ------------------| --------|
|  dmNothing          |  Addせずになにもしない|
|  dmOverwrite       |  要素を上書きする|
|  dmSetDuplicateCollection       |  Addせずに重複コレクションに追加する|
|  dmSetBothDuplicateCollection          |  Addせずに重複コレクションに追加する。既にセットした重複された側も追加する|

### Exists

\<引数> キー  
\<戻値> 真偽値 

- 既にキー使われている場合はTrueを返す


### Items

\<引数> キー  
\<戻値> 要素

- キーに連想された要素を返す
- 既定メンバ
- 元のdictionaryと異なり読み取り専用です

### Key
\<引数> 変更するキー、新しいキー   
\<戻値> なし

- 新しくキーを変更する(要素は変わらない)

### Remove
\<引数> 削除するキー  
\<戻値> なし

- 要素を削除します

### SetCompareMode

\<引数> VbCompareMethodの列挙型  
\<戻値> なし 

- 文字列キーを比較するときの比較モードを設定できます。
- Dictionaryとほぼ同じですが、vbUseCompareOptionの設定は無効にしています。
- コンストラクタ時点ではvbTextCompareの設定です。


## Properties

### Count

\<引数> なし  
\<戻値> 要素数 

- 要素数を返します

### DuplicateCollection

\<引数> なし  
\<戻値> 重複した要素のコレクション  

- AddIfで重複のしたものをコレクションする設定にした場合、コレクションが返されます。
- 重複したものだけ修正したり、利用者に教えるために使うと便利です。

### Items

\<引数> なし  
\<戻値> 要素の配列 

- 要素の内容を全て返す
- 最初に入れた要素がインデックス番号が一番最初

### Keys

\<引数> なし  
\<戻値> キーの配列 

- キーの文字列を全て返す
- 最初に入れた要素がインデックス番号が一番最初