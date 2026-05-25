# 本地离线运行说明

这个报价网页只在本机运行，不需要部署到 Supabase、Netlify 或公网。

## 启动

在当前目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\Start-QuoteSite.ps1
```

如果 8787 端口被占用，可以换一个端口：

```powershell
powershell -ExecutionPolicy Bypass -File .\Start-QuoteSite.ps1 -Port 8789
```

然后打开：

```text
http://127.0.0.1:8789/
```

## 使用

1. 默认读取 `assets/sales-price-template.xls`。
2. 也可以在页面里上传新的 `.xls`、`.xlsx` 或 `.xlsm` 价格表。
3. 在文本框输入产品，每行一个，例如：

```text
RT30 | 30mg
MS40 | 40mg
WA10 x 3
```

4. 点击“预览匹配”查看已匹配和未匹配产品。
5. 点击“生成并下载 Excel”导出 `报价单_日期时间.xlsx`。

## 输入格式

- `RT30`：按缩写或产品名匹配。
- `RT30 | 30mg`：产品有多个规格时，用 `|` 后面的规格帮助匹配。
- `MS40*2` 或 `MS40 x 2`：数量为 2。

页面依赖的 SheetJS 已保存为本地文件 `assets/xlsx.full.min.js`，运行时不会访问公网。
