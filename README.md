# JIT 面单商品名称注释工具 Streamlit 版

这个版本适合放到 Streamlit Community Cloud。商品对应表可以直接放 Excel：

```text
data/mapping.xlsx
```

同事打开网页后只需要上传面单 PDF，不用每次上传 Excel。

对应表格式：

- A 列：商品编码
- B 列：商品名称
- 可以有表头，也可以没有表头
- 优先读取 `Sheet2`；如果没有 `Sheet2`，读取第一个工作表

## 功能

- 读取面单 PDF
- 按两页一组处理：揽收面单 + 货品标签
- 从货品标签识别商品编码
- 用 `data/mapping.xlsx` 匹配商品名称
- 每张揽收面单前插入同尺寸空白页
- 在空白页写该揽收面单下全部商品的编码、名称和件数
- 支持一张揽收面单对应多个产品标签；相同编码会合并显示数量
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
data/mapping.xlsx
```

如果暂时还没有 `data/mapping.xlsx`，程序也兼容旧的 `data/mapping.csv`。

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

推荐方式：直接在 GitHub 更新 Excel。

1. 在 GitHub 打开 `data` 文件夹。
2. 上传新的 Excel 对应表。
3. 文件名保持为：

```text
mapping.xlsx
```

4. 提交后 Streamlit 会自动重新部署或刷新后使用新版对应表。

网页左侧也有“更新对应表文件”，那里上传 Excel 只是临时测试，本次页面会话生效；要长期生效，还是上传到 GitHub 的 `data/mapping.xlsx`。
