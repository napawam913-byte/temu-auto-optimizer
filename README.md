# Temu Auto Optimizer

Temu 半托管自动上架助手，Tkinter 桌面版。支持 Excel 拆分、批量标题/描述优化、图片下载压缩、去重、类目过滤、网页预览和图片微调，并输出店小秘 TEMU 半托管导入模板。

## 运行

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

首次使用可以复制 `config.example.json` 为 `config.json`，再在软件设置页填写自己的 API Key。`config.json` 会保存本机密钥，已经被 `.gitignore` 忽略，不要上传到 GitHub。

## 打包

```bash
pyinstaller --noconfirm --onefile --windowed --name TemuAutoOptimizer main.py
```

生成文件在 `dist/TemuAutoOptimizer.exe`。

## 关键说明

- 支持云启数据导出的特殊文件：即使扩展名是 `.csv`，只要内容实际是 Excel，也会自动识别读取。
- 自动上架页可以选择最终上架模板，例如 `D:\Downloads\import_created_product_popTemu.xlsx`，程序会复制模板并填充 `导入模板` 工作表。
- 自动上架页可以指定输出目录，并自定义结果 Excel 文件名；程序会在输出目录下创建 `temu_output_时间戳` 子文件夹，避免覆盖旧批次。
- 拆分工具页可以指定拆分输出目录；程序会在该目录下创建 `split_时间戳` 子文件夹，并生成拆分 Excel 和 `split_files.zip`。
- 设置页已将模型配置拆成两块：`文字模型（标题优化 / 描述生成）` 和 `图片模型（图片质检 / 背景处理 / 修图）`。
- 标题优化会使用原中文标题和已有英文标题作为参考，按“标题目标字符数 ±5”生成英文标题；如果命中 `config/sensitive_words.json` 中的敏感词或长度不合格，会自动带着失败原因重试。
- DeepSeek 文字模型可设置为：Provider `deepseek`，Base URL `https://api.deepseek.com`，Model `deepseek-chat`。
- OpenAI 图片模型可设置为：Provider `openai`，Base URL `https://api.openai.com/v1`，Model `gpt-image-1`。
- 图片微调提示词和微调图片数量在设置页维护，默认每个商品微调 2 张。
- 自动上架页的 `网页预览` 按钮可选择处理后的 Excel 文件，并打开本地网页预览；网页中可查看标题、英文标题、申报价格、预览图和轮播图。
- 网页预览支持响应式布局、分页、每页数量自定义、轮播图横向滚动、一键微调未修改商品，并会跳过已经微调过的商品。
- 店小秘多图字段使用换行符分隔，避免“多条数据分隔符不正确”的导入错误。
- 若不填写 API Key，程序会使用降级逻辑：保留标题并生成基础英文描述，方便先测试 Excel 流程。
