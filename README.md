# Freight Invoice Merger Pro

一款现代化的 Windows 桌面应用程序，用于自动整合 Freight Invoice（运费发票）数据。

## 功能特点

- 📂 **模板选择** — 选择 Excel 模板文件
- 📦 **ZIP 批量处理** — 上传包含 Invoice Excel 文件的 ZIP 压缩包
- ⚡ **一键整合** — 自动解压、提取、填充、排序
- 🧩 **插件化规则** — 支持 Rule1、Rule2... 无需修改核心代码
- 🖥️ **现代 UI** — 采用 Windows 11 / Office 365 风格的界面设计
- 📊 **实时日志** — 处理过程实时显示，支持滚动和复制

## 快速开始

### 方式一：直接运行 EXE（无需安装 Python）

1. 下载 `FreightInvoiceMerger.exe`
2. 双击运行
3. 选择模板文件（`Freight Invoice list 2026Fareast.xlsx`）
4. 选择包含 Invoice 的 ZIP 压缩包
5. 点击「开始整合」
6. 等待处理完成，自动打开结果文件

### 方式二：从源码运行

#### 安装依赖

```bash
pip install -r requirements.txt
```

#### 运行

```bash
python main.py
```

## 项目结构

```
FreightInvoiceMerger/
├── app.py              # 应用入口
├── main.py             # 主程序
├── config.json         # 配置文件
├── requirements.txt    # 依赖列表
├── generate_samples.py # 示例数据生成器
├── build.py            # 打包脚本
├── core/
│   ├── processor.py    # 主处理流程
│   ├── extractor.py    # 数据提取
│   ├── parser.py       # 发票解析
│   ├── validator.py    # 输入验证
│   ├── sorter.py       # 数据排序
│   └── writer.py       # 模板写入
├── rules/
│   ├── __init__.py     # 规则注册器
│   └── rule_v1.py      # 规则1（默认）
├── ui/
│   ├── main_window.py  # 主窗口
│   ├── widgets.py      # 自定义控件
│   └── styles.qss      # 样式表
├── utils/
│   ├── logger.py       # 日志工具
│   ├── excel.py        # Excel 工具
│   └── zip_helper.py   # ZIP 工具
└── resources/          # 资源文件
```

## 打包为 EXE

```bash
python build.py
```

打包后生成：`dist/FreightInvoiceMerger.exe`

### 打包要求

- Python 3.13+
- PyInstaller 6.0+

## 配置说明

`config.json`:

```json
{
    "output_name": "Freight Invoice list 2026Fareast_Output.xlsx",
    "open_after_finish": true,
    "auto_sort": true,
    "remember_last_path": true
}
```

- `output_name`: 输出文件名
- `open_after_finish`: 完成后自动打开
- `auto_sort`: 自动按 U 列排序
- `remember_last_path`: 记住上次路径

## 添加新规则

1. 在 `rules/` 目录下创建新文件，如 `rule_v2.py`
2. 定义 `RULE_META` 和 `FIELD_MAPPINGS`
3. 自动注册，无需修改任何 UI 或核心代码

## 许可证

MIT License
