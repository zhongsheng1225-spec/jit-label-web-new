# JIT 面单商品名称注释工具 Streamlit 版

这个版本适合放到 Streamlit Community Cloud。仓库里已经带有小型对应表：

```text
data/mapping.csv
```

同事打开网页后只需要上传面单 PDF，不用每次上传原始 Excel。

## 功能

- 读取面单 PDF
- 按两页一组处理：揽收面单 + 货品标签
- 从货品标签识别商品编码
- 用 `data/mapping.csv` 匹配商品名称
- 每张揽收面单前插入同尺寸空白页
- 在空白页写商品编码和商品名称
- 下载新 PDF
- 下载复核清单 CSV
- 拦截页数异常、未识别编码、未匹配商品名称等问题

## 上传到 GitHub

1. 新建或清空一个 GitHub 仓库。
2. 上传本文件夹里的所有文件。
3. 仓库根目录需要看到：

```text
streamlit_app.py
jit_processor.py
requirements.txt
data/mapping.csv
```

## 部署到 Streamlit

1. 打开 Streamlit Community Cloud。
2. 选择 New app。
3. 连接你的 GitHub 仓库。
4. Main file path 填：

```text
streamlit_app.py
```

5. 点 Deploy。
6. 部署完成后，Streamlit 会给你一个网址。

## 日常使用

1. 打开 Streamlit 网址。
2. 上传面单 PDF。
3. 先看复核清单。
4. 复核通过后下载新 PDF。

## 更新对应表

网页左侧有“更新对应表文件”。

1. 上传原始 Excel 对应表。
2. 下载生成的 `mapping.csv`。
3. 回到 GitHub，进入 `data/mapping.csv`。
4. 用新下载的 `mapping.csv` 替换旧文件。
5. Streamlit 会自动重新部署或刷新后使用新版对应表。

这样做的原因：原始 Excel 很大，直接在线上传处理会慢；`mapping.csv` 只有编码和商品名称，体积很小，读取速度快很多。
