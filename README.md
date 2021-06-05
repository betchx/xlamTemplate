# xlamTemplate
A template of Excel VBA addin file with automatic menu etc.

## Feature
Add easy Access to created macros.

### Automatic Dynamic Menu
It scan internal codes and add buttons into a dynamic menu in the "Home tab".
Special comments before subroutine definition provide follows:
*  Customized icon
*  Customized Caption
*  Tooltips
*  Visibility
*  Context based Visibility

#### Customized Icon
"' imageMso: idMSO" can change button icon to specified one by the isMSO.
The idMSO can be obtained from Microsoft's information.
(see https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/3f465fd5-af0e-4d3b-80eb-53f90f5c7a3a )

#### Customized Caption
Macro name is used for button caption as default.
"' label: TEXT" can change the caption to the TEXT.

#### Tooltips
Two type tooltips can be added.
"' screentip:" adds comment into upper side of tooltip. (default: name of Macro)
"' supertip:" adds comments into lower side od tooltip. (default: none)

#### Visibility
Non-private subroutines without any argument will be added into the dynamic menus.
"' hidden:" can suppress this behavior for next subroutine even if it is no-argument non-public subroutine.

#### Context based visibility
"' target: TYPE" can limit the visibility of button and menu by selection.
Following macro will be shown only if any of selection and its parent is TYPE.
The TYPE is usually Range, Chart or Shape.

### Exporting code

### Save as addin

### Automatic disabling installed addin


# xlamテンプレート
Microsoft ExcelのVBAアドイン開発者用テンプレート

## 特徴
内部のマクロを実行しやすくくします．

### 動的自動メニュー
アドイン内部のコードを走査して，マクロを呼び出すメニューをホームタブに追加します．

#### ボタンのカスタマイズ
マクロ定義の直前に記入したスペシャルコメントにより，次のようなカスタマイズが可能．
* アイコンの変更
* キャプションの変更
* ツールチップの追加
* 可視性
* 選択表示

ほかにもあるが，詳細は仕様書を参照．

#### アイコンの変更
"' imageMso: idMSO"
idMSOで指定されたものにアイコンを変更する．
idMSOはマイクロソフトからの情報を参照のこと．
(参考： https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/3f465fd5-af0e-4d3b-80eb-53f90f5c7a3a )

#### キャプションの変更
通常はマクロ名がキャプションとして使用される．
"' label: TEXT" を追加することで，TEXTをボタンのキャプションにできる．

#### ツールチップの追加
ツールチップの上下2か所の記入欄にコメントを記入可能．
上部： "' screentip:"  (デフォルトはマクロ名）

#### 可視性
Privateでなく，引数のないサブルーチンがマクロとしてメニューに登録される.
しかし，"' hidden:"を追加することで，登録されないようにできる．

#### 選択表示
現在選択されているもの応じてボタンを表示するかどうかを制御できる．
"' target: TYPE"を追あすると，Selectionもしくはその親(Parent)にあるものの型がTYPEの場合にのみ表示される様になる．
TYPEにはたいていの場合，"Range"か"Chart"で，たまに"Shape"も用いられる．

### コードの書き出し

### アドインとして保存

### インストール済みアドインの自動無効化
