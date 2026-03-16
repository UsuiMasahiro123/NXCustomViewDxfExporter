# 修正指示: ビュー選択ダイアログのフォント修正（v4）

## 問題
チェックリストボックス内の文字だけ滲んでいる。
ダイアログのタイトルやボタンの文字は綺麗に表示されている。

## 原因
CheckedListBoxのフォントレンダリングにClearTypeが適用されていない。

## 修正内容
CheckedListBoxに以下を適用する：
- Graphics.TextRenderingHint を ClearTypeGridFit に設定
- フォントを new Font("Meiryo UI", 9f, GraphicsUnit.Point) に変更
- CheckedListBoxをオーナードローモード（OwnerDrawFixed）に変更して
  DrawItemイベントでTextRendererHint.ClearTypeGridFitを使って描画する

## ビルドまでお願いします