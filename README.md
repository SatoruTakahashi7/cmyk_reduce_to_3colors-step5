スクリプト名：cmyk_reduce_to_3colors-step5  
動作環境：MacOS X14.7.6、Illustrator2024（v28.7.7）  
開発環境：Mac Studio (2022)  
開発言語：JavaScript(ChatGPTによる）  


## 何をするスクリプトか？  
### スクリプトの目的
- Adobe Illustrator上で、選択されたオブジェクトのCMYKカラー（4色）を、CMYのうち2色＋K（ブラック）の合計3色構成に変換し、印刷時の色味をできる限り維持するスクリプトです。  

### 処理対象と条件  
- 対象オブジェクト：Illustrator上で選択中のオブジェクト（パス、テキストなど）  
- 対象のカラー属性：塗り（Fill）および線（Stroke）の両方に対応  
- 変換対象の条件：CMYKすべてのチャンネルが 0 以上のカラー（すなわち、4色全てに値があるカラー）  

### 変換のルール  
- CMYのうち1チャンネルのみを削除（= 0に）し、残りの2チャンネル＋Kで構成された3色カラーに変換します。   
- 単純な「最小値を0にする」ルールではなく、色相に基づく判断や色差を考慮したロジックを用いて、変換後も視覚的な色味が近くなるようにします。   
- 必要に応じて、CMYおよびKの値を微調整して、人間の目で見たときの色の変化を最小限に抑えます。   

### 想定される処理例  
- 入力カラー例：C=11.2, M=16.4, Y=5.2, K=20.3  
- 変換処理：色差が小さくなるよう、たとえばY=0にしつつCとMを補正し、Kも微調整することで、視覚的に元の色に近い3色構成のカラーを再構築

### 主な仕様  
- 選択中のオブジェクト（PathItem/GroupItem/CompoundPathItem/TextFrame）の塗り・線がCMYKカラーの場合に変換。  
- CMYK→Lab 変換には app.convertSampleColor() を使用（Illustratorのカラー設定で"Japan Color 2001 Coated"が選ばれている前提）。  
- (C, M, Y) のうち1チャネルを0にして2チャネル + Kの計3チャネルで総当たりし、Lab色差(ΔE76)が最小となるものを求める。  
- 変換結果は最後に Math.round() で整数に丸める。  
- 計算負荷が高いため、STEP値が小さいと処理が非常に遅くなる可能性があります。  

### 使い方  
1) Illustratorで作業用CMYKプロファイルを"Japan Color 2001 Coated"に設定しておく。  
2) 変換したいオブジェクト（CMYK塗り/線を持つもの）を選択。  
3) このスクリプト(.jsx)を実行すると、選択オブジェクトのCMYKカラーが3色化される。  
4) しかし、浮動小数表現の誤差が完全にゼロにならないので、数値を手入力し直さないとならない（どうすればいいのかわからない）。  

## 使用条件  
このスクリプトが正常に動作する環境は以下の通りです。  
- MacOS X14.7.6  
- Illustrator2024（v28.7.7）  

## インストール  
スクリプト本体（cmyk_reduce_to_3colors-step5.jsx）を  
~/Applications/Adobe Illustrator 2024/Presets.localized/ja_JP/スクリプト  
にコピーしてください。エイリアスを入れておくだけでもかまいません。  

## 免責事項  
- 本アプリケーションはIllustratorにおける作業効率支援なのであって、処理結果を保証するものではありません。かならず確認をされることをおすすめします。  
- このツールを使用する上でデータの破損などのあらゆる不具合・不利益については一切の責任を負いかねますのでご了解ください。  
- このツールはすべてのMacintoshとMac OS上で動作をするという確認をとっていませんし、事実上出来ません。したがって、動作を保証するものではありません。  

## 履　歴  
2025-05-19	ver.0.1	とりあえず  

GYAHTEI Design Laboratory  
髙橋聡  
