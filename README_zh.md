# 智填 FormFiller

一个本地运行的智能填表工具：从资料文件（简历、信息表等）提取信息 → 智能匹配 → 填充 Word/Excel 模板 → 生成与模板格式完全一致的文件。

**核心特点**：纯规则实现（别名表 + 相似度匹配），不调用任何大模型 API，不联网，可离线运行，数据全部留在本机 `data/` 目录。

## 支持两类模板

| 模板类型 | 示例 | 识别方式 |
|------|------|------|
| `{{字段名}}` 占位符模板 | 入职登记表、报价单 | 自动识别占位符 |
| 「标签格 + 空格」传统表单 | 高校申报书、应聘人员信息表等机关/高校正式表格 | 解析到 0 个占位符时自动切换表单模式 |

## 主要功能

- **资料库**：上传文档（docx/docm/xlsx/pdf/txt/md/csv/json 及图片）自动提取键值信息；正文写在**文本框（textbox）**里的简历也能提取；持久化在 `data/db.json`，建一次反复用
- **智能匹配**：别名表（`联系电话`→`电话`、`毕业院校`→`学校` 等）+ 后缀匹配 + 字符重叠相似度，资料字段与模板字段**不需要名字完全一致**
- **派生字段**：按月份精算年龄；复合标签（如「最后学历毕业院校及学位」）自动从多个源字段拼接
- **多行列表区块**：学习经历 / 工作经历 / 已发表论文 / 家庭成员等「表头 + 数据行」表格，按列名相似度自动映射，数据行数不够自动加行
- **论文引用解析**：从简历文本（含文本框）解析 APA 引用，拆出作者排序、标题、期刊、ISSN、影响因子、收录情况、年卷期页；内置常见期刊 ISSN 映射表；用 `--set 本人=<英文姓氏>` 自动推导作者排序
- **图片占位符**：`{{img:头像}}` 自动插入资料库图片，`头像_1.png`、`头像(2).png` 等序号变体均可匹配
- **两种使用方式**：CLI 一条命令出结果；Web 界面（资料库 + 模板管理 + 生成记录）适合反复使用

## 快速开始

### 安装依赖

```bash
pip install -r backend/requirements.txt
```

### CLI 用法

```bash
# 1. 查看模板结构（有哪些占位符 / 表单字段）
python cli.py scan --template 模板.docx

# 2. 填充（多份资料可同时传入）
python cli.py fill --template 模板.docx --source 简历.docx 岗位信息.txt --out 结果.docx

# 缺的字段用 --set 补（可重复传）
python cli.py fill --template 模板.docx --source 简历.docx --set 签名=张三 --out 结果.docx

# 表单模式 + 论文表：指定本人英文姓氏以自动推导作者排序
python cli.py fill --template 申报书.docx --source 简历.docx 信息表.docm --set 本人=Wang --out 申报书_已填.docx
```

有缺失字段时默认**拒绝生成并列出缺什么**；确认可以留空再加 `--force`。

### 启动 Web 界面

```bash
cd backend
python -m uvicorn main:app --host 127.0.0.1 --port 8765
```

打开浏览器访问 http://127.0.0.1:8765

> 注意：**必须在 `backend/` 目录下执行**。后端用扁平导入（`import extractor`），从别处启动会报 ModuleNotFoundError。

## 运行测试

```bash
# 模块级测试（50 项，覆盖提取、匹配、填充、论文解析）
cd backend
python test_modules.py

# API 端到端测试（需先启动服务）
python test_api.py
```

## 项目结构

```
├── cli.py             命令行入口（scan / fill / serve）
├── backend/           FastAPI 后端
│   ├── main.py        API 服务入口
│   ├── extractor.py   文件信息提取（docx/docm/xlsx/pdf/txt/…，含文本框）
│   ├── filler.py      占位符模板解析与填充
│   ├── form_analyzer.py  表格型表单分析、列表区块、论文引用解析
│   ├── storage.py     本地数据与文件存储
│   ├── make_samples.py   生成内置示例
│   ├── test_modules.py   模块级测试
│   └── test_api.py       API 端到端测试
├── frontend/          单页 Web 前端
├── sample/            内置示例（员工入职登记表、报价单、虚构简历等，可直接跑通流程）
└── data/              运行时数据（资料库、模板、生成文件；首次运行自动创建，不入库）
```

## 支持格式

- **模板**：`.docx`、`.xlsx`（表格型表单模式目前仅支持 .docx）
- **资料文件**：docx、docm（含宏 Word，读取时在内存中转正 content type）、xlsx、pdf、txt、md、csv、json 及常见图片
- **`.doc` 旧格式**：提取器不直接读 OLE2 二进制。Windows 上有 Word 时，可先用 Word 另存为 `.docx` 再喂给工具

## 注意事项

- 生成后的 Word/Excel 会尽量保持原模板格式；复杂图表、宏、特殊控件可能无法保留（受 python-docx / openpyxl 限制）
- 公司名等字段可能填在**页眉/页脚**里，校验时要连页眉一起扫
- 生成结果建议**人工复核**，尤其是签名、日期、金额这类字段
- `data/` 目录存放你的真实资料，已在 `.gitignore` 中排除，不会随仓库提交

> English version: [README.md](README.md)
