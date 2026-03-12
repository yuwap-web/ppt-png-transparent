# PowerPoint PNG 透過コンバーター

PowerPointで書き出したPNG画像の背景を、簡単に透過PNGへ変換するWindows用ツールです。

## 主な機能

- PNGをEXEに直接ドロップして変換
- フォルダごとの一括変換
- 1920x1080固定出力
- PowerPointの緑背景を優先的に透過
- 緑背景でない場合も四隅の背景色を自動検出して透過
- PNGとWebPを同時出力

## 使い方

1. `PPTPNGTransparentConverter.exe` を用意
2. PNGファイルまたはPNGが入ったフォルダをEXEにドラッグ＆ドロップ
3. 同じフォルダに変換後ファイルが保存されます

## 出力ファイル名

- `元ファイル名_transparent_1920x1080.png`
- `元ファイル名_transparent_1920x1080.webp`

## 推奨PowerPoint設定

- 背景色: `RGB(0,255,0)`
- スライドサイズ: `16:9`

## 動作環境

- Windows 10
- Windows 11