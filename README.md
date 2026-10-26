# okyu-slides

Central AX がお客様向けに用意した、Claude Code 用のスライド作成 skill です。デザインは2種類あります。

| デザイン | 見た目 | 頼み方の例 |
|---|---|---|
| A | 紺の枠×大きい数字 | 「okyu-slides の A で〇〇のスライドを作って」 |
| B | グレーの紙面×紺のコーナー×薄い特大番号 | 「okyu-slides の B で〇〇のスライドを作って」 |

## 入れ方

Claude Code で次の2行を1行ずつ入力します。

```
/plugin marketplace add https://github.com/Central-AX/okyu-slides.git
/plugin install okyu-slides@okyu-slides
```

入れたあと Claude Code を再起動してください。

コマンドが使えない画面のときは、Claude Code にこう頼めば入れてくれます。

> https://github.com/Central-AX/okyu-slides のプラグインを入れて

初回にスライドを作るとき、必要な Python のライブラリ（playwright・python-pptx・lxml・pillow）と Chromium を Claude Code が入れます。

## 注意

- 表紙・章扉の写真とロゴは仮画像です。お使いの会社の写真・ロゴに差し替えてお使いください
- 見本画像（`見本/`）の文言・数字はデザインの参考用です
- B の書体は LINE Seed JP（同梱・SIL Open Font License）です。PowerPoint で同じ見た目にするには、PC に LINE Seed JP（OTF版・無料）を入れてください

## 更新

```
/plugin marketplace update okyu-slides
```
