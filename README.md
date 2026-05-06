# 统班应用 (TonBanAPP)

## 简介
统班应用是一个用于资助服务中心排班统计的桌面应用程序。它可以帮助管理员快速解析排班名单，生成周统计和月统计报表，并支持加班补录和放假核减等功能。

## 功能特点
- **总名单管理**：支持从Word文档导入和更新总名单。
- **周统计生成**：
  - 解析排班Word文档，自动计算应值班次和实际班次。
  - 支持放假核减（自动扣减放假当天的应值班次）。
  - 支持加班补录（手动添加额外班次）。
  - 生成包含“姓名”、“应值班次”、“实际班次”、“缺班”和“备注”的Excel报表。
- **月统计生成**：
  - 汇总多个周统计Excel文件。
  - 生成月度统计Word文档，包含多值班次和缺班数量。

## 目录结构
- `main.py`: 应用程序入口。
- `src/`: 源代码目录。
  - `ui/`: 用户界面模块。
    - `main_ui.py`: 主界面实现。
  - `modules/`: 核心业务逻辑模块。
    - `data_manager.py`: 数据管理（总名单等）。
    - `excel_generator.py`: Excel报表生成。
    - `word_generator.py`: Word报表生成。
    - `word_parser.py`: Word文档解析。
  - `utils/`: 工具函数模块。
    - `helpers.py`: 辅助函数（如拼音排序等）。
- `data/`: 数据存储目录（如 `config.json`）。
- `output/`: 生成的报表输出目录。
- `tests/`: 单元测试和集成测试目录。

## 依赖环境
- Python 3.8+
- 依赖包请参考 `requirements.txt`：
  - `python-docx`
  - `pandas`
  - `openpyxl`
  - `pypinyin`
  - `customtkinter`
  - `pyinstaller`
  - `pytest`
  - `pytest-cov`

## 安装与运行
1. 克隆或下载本项目代码。
2. 创建并激活虚拟环境（推荐）：
   ```bash
   python -m venv .venv
   # Windows
   .venv\Scripts\activate
   # Linux/macOS
   source .venv/bin/activate
   ```
3. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
4. 运行应用：
   ```bash
   python main.py
   ```

## 使用说明
1. **基础设置**：在“基础设置”标签页中，点击“导入 / 更新总名单 (Word)”按钮，选择包含所有人员名单的Word文档。
2. **周统计生成**：在“周统计生成”标签页中，选择排班Word文档，设置放假日期和加班人员，点击生成按钮即可在 `output` 目录下生成周统计Excel文件。
3. **月统计生成**：在“月统计Word生成”标签页中，选择需要汇总的周统计Excel文件，点击生成按钮即可在 `output` 目录下生成月统计Word文件。

## 测试
运行所有测试用例：
```bash
pytest
```
生成测试覆盖率报告：
```bash
pytest --cov=src --cov-report=html
```
覆盖率报告将生成在 `htmlcov` 目录下，打开 `index.html` 即可查看。