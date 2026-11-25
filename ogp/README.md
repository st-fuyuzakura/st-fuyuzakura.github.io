# 共有画像（ogp.png）の作り方

`ogp.png`（1200×630）は `ogp/source.html` をヘッドレスChromeで撮影したもの。料金・納期・音源を変えたら作り直す。

1. 本物の音源に替えたときだけ：波形の値 `P` を作り直す（`afconvert` で wav にしてから 258 区間の音量を計算し、MIX前とマスタリング後を同じ物差しで 0〜1 にそろえる）。
2. 撮影：

```bash
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --disable-gpu --hide-scrollbars --window-size=1200,630 --virtual-time-budget=15000 --screenshot=ogp.png "file://$PWD/ogp/source.html"
```

SNS側は画像をキャッシュするので、差し替え後に反映されないときはファイル名を変えて `index.html` の `og:image` と `twitter:image` も直す。
