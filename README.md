# TonBanAPP

面向 HUST 资助中心助理值班统计的 Windows 桌面应用。TonBanAPP 将总名单维护、特殊规则配置、周统计、期末周统计和月统计整合到同一套流程中，可从排班 PDF/Word、实际值班 Word 和历史统计文件生成规范的 Excel、Word 报表。

- 当前版本：`v1.8`
- 运行平台：Windows
- 技术栈：Python、CustomTkinter、pandas、openpyxl、python-docx、Tesseract OCR

[下载 TonBanAPP v1.8](https://github.com/JunDengH/TonBanApp/releases/tag/v1.8)

## 功能概览

- 从 Excel 通讯录或旧版 Word 名单建立助理总名单。
- 维护姓名、部门、专业、年级、毕业季和长期请假标记。
- 配置毕业季三档值班规则和长期请假规则。
- 读取文本型或扫描型排班 PDF，也兼容排班 Word。
- 根据排班文件、实际值班 Word 和放假日生成周统计 Excel。
- 将两个期末周合并为一个统计周期，生成期末周 Excel。
- 汇总 1～4 份周统计 Excel，结转上月数据并生成月统计 Word。
- 手动维护或通过 Word 导入、导出加班补录与罚班记录。
- 在生成文件前预览、校验并修正统计数据。
- 支持深色/浅色主题以及自定义输出目录。

## 业务流程

```text
导入或维护总名单
        ↓
配置毕业季 / 长期请假规则
        ↓
生成普通周统计或期末周统计 Excel
        ↓
汇总 1～4 份周统计 Excel
        ↓
结转上月数据 + 加班补录 + 罚班记录
        ↓
生成月统计 Word
```

## 快速开始

### 直接运行发布版

从 [Releases](https://github.com/JunDengH/TonBanApp/releases) 下载 `TonBanAPP.exe` 后直接运行。发布版已经包含离线 Tesseract OCR，无需单独安装 OCR 软件。

当前发布版下载地址：

[TonBanAPP.exe（v1.8）](https://github.com/JunDengH/TonBanApp/releases/download/v1.8/TonBanAPP.exe)

### 从源码运行

要求：

- Windows 10/11
- Python 3.10 或更高版本
- 可用的 Tcl/Tk 图形环境

在 PowerShell 中执行：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

也可以双击根目录的 `start_tonban.bat`。该脚本会优先使用 `.venv\Scripts\python.exe`，否则尝试使用系统的 `py` 启动器。

## 页面与输入输出

| 页面 | 主要用途 | 输入 | 输出 |
| --- | --- | --- | --- |
| 总名单 | 建立和维护人员基准数据 | Excel 通讯录或 Word 名单 | 保存到本地配置 |
| 规则配置 | 维护毕业季和长期请假规则 | 当前总名单 | 保存到本地配置 |
| 周统计 | 计算单周应值、实值和缺班 | 排班 PDF/Word、实际 Word、放假日 | 周统计 Excel |
| 期末周统计 | 合并两个期末周进行统计 | 双周期末排班 PDF、两份实际 Word | 期末周统计 Excel |
| 月统计 | 汇总周表并结转历史数据 | 1～4 份周统计 Excel、可选上月 Word | 月统计 Word |
| 概览 | 查看名单、规则和流程状态 | 本地配置 | 无 |

## 统计规则

### 普通周统计

周统计输出列为：

```text
姓名 / 应值班次 / 实际班次 / 缺班 / 备注
```

基础口径：

```text
应值班次 = 排班文件中的姓名出现次数，经放假和特殊规则调整
实际班次 = 实际值班 Word 中的姓名出现次数
缺班 = max(0, 应值班次 - 实际班次)
```

规则优先级与行为：

- 放假核减仅支持周一至周五，只核减排班文件中对应日期的班次。
- 长期请假助理的应值班次固定为 2。
- 毕业季规则优先于长期请假规则。
- 毕业季“正常值班”保留排班结果。
- 毕业季“少值班”将应值班次设为 1，再应用放假核减，最低为 0。
- 毕业季“无需值班”将应值班次设为 0。
- 生成前可修改应值班次、实际班次和备注，缺班会自动重算。

程序会对以下情况给出生成前提醒：

- 单人本周应值班次大于 2。
- 文件中出现不在总名单内的姓名。
- 选中的 Word 未识别到任何总名单内姓名。
- 未知姓名与总名单姓名存在字符或拼音相似关系，可能是错别字。

### 期末周统计

当前期末周流程会把两个周次合并成一个统计周期：

- 排班 PDF 必须包含两个带周次标题的区块，两个周次不要求连续。
- 两个区块分别对应两份实际值班 Word，顺序可在界面中核对。
- 排班 PDF 中出现的唯一姓名在整个期末周期应值 2 班；重复排班会提示，但应值仍为 2。
- 两份实际 Word 合并统计，每次姓名出现按 2 班计算。
- 期末周不执行放假核减。
- 毕业季“少值班”与期末周双倍计数冲突，存在生效人员时会阻止生成。
- 默认只生成一份合并 Excel。

默认文件名：

```text
{年份}{春季/秋季}学期期末周助理值班统计.xlsx
```

### 月统计

月统计输出列为：

```text
序号 / 姓名 / 总计班次 / 应值班次 / 多值班次 / 缺班数量
```

计算口径：

```text
总计班次 = 本月周表实际班次 + 上月多值班次 + 本月加班补录
应值班次 = 本月周表应值班次 + 上月缺班数量 + 本月罚班记录
多值班次 = 总计班次 - 应值班次
缺班数量 = 应值班次 - 总计班次
```

- 至少选择 1 份、最多选择 4 份周统计 Excel。
- 上月 Word 可留空；留空时历史结转按 0 处理。
- 只统计当前总名单内人员。
- 加班补录计入总计班次，罚班记录计入应值班次。
- 生成前可修改总计班次和应值班次，多值、缺班会自动重算。
- 最终 Word 中非正数的多值/缺班单元格会留空并绘制斜线。

默认文件名：

```text
{年份}{春季/秋季}学期{起始周}-{结束周}周助理值班统计.docx
```

## 文件格式约定

### 总名单

总名单页面支持：

- `.xlsx` 通讯录：需要能识别到“姓名”和“专业年级”，或“姓名”“专业”“年级”表头。
- `.docx` 名单：兼容旧流程，只导入正文段落和表格中的姓名。

Excel 通讯录解析具有以下规则：

- 不依赖固定列顺序，在前 20 行内定位表头。
- 部门列为空时沿用上一条非空部门。
- 支持 `新闻22级`、`24级法学`、`材料 24级` 等专业年级格式。
- 根据可解析届号中最靠前的一组人员自动标记毕业季助理。
- 姓名按首次出现去重，并归一化汉字之间的全角空格。

### 排班文件和实际值班文件

- 普通周排班支持 `.pdf` 和 `.docx`，优先推荐 PDF。
- 实际值班名单使用 `.docx`。
- 文本型 PDF 直接解析表格中的星期列和姓名。
- 扫描型 PDF 会渲染页面、检测表格网格，再调用内置 Tesseract OCR。
- 多页 PDF 会逐页合并，姓名重复出现会保留为多次排班。
- 扫描模糊、表格倾斜或边框缺失可能降低识别率，生成前应检查预览和未知姓名提醒。

Word 排班中的星期标题应单独成行。支持：

```text
周一 / 星期一 / 礼拜一 / 周1
周二 / 星期二 / 礼拜二 / 周2
周三 / 星期三 / 礼拜三 / 周3
周四 / 星期四 / 礼拜四 / 周4
周五 / 星期五 / 礼拜五 / 周5
周六 / 星期六 / 礼拜六 / 周6
周日 / 星期日 / 星期天 / 礼拜日 / 礼拜天 / 周7
```

### 周统计 Excel

供月统计读取的周表至少需要以下列：

```text
姓名 / 应值班次 / 实际班次
```

计数必须为空或非负整数。月统计会忽略不在当前总名单中的行。

### 上月月统计 Word

上月 Word 至少需要一个包含以下列的表格：

```text
姓名 / 多值班次 / 缺班数量
```

空单元格、非数字文本、负数以及斜线占位单元格均按 0 处理。

### 加班补录和罚班记录

月统计页面支持手动维护、导入和导出：

- 加班补录 Word：`姓名 / 加班班次`
- 罚班记录 Word：`姓名 / 罚班次数`

姓名必须存在于总名单中，次数必须为正整数；同一人多行会累加。导入会覆盖界面中对应类型的现有记录。

根目录提供：

- `加班补录空白模板.docx`
- `罚班记录空白模板.docx`

可运行 `python create_overtime_template.py` 重新生成模板。

## 本地数据与输出目录

源码运行时，应用会在根目录创建并维护：

```text
data/config.json
```

其中包含：

- 总名单与人数
- 部门、专业和年级
- 毕业季、长期请假人员及规则模式
- 深色/浅色主题
- 自定义输出目录

`data/config.json` 已被 Git 忽略，不应作为公共数据提交。默认输出目录为根目录下的 `output/`；可在应用设置中改为其他文件夹。同名报表不会直接覆盖，程序会自动生成唯一文件名。

## 项目结构

```text
TonBanAPP/
├── main.py                         # 应用入口
├── start_tonban.bat                # Windows 快速启动脚本
├── requirements.txt                # 运行与打包依赖
├── sitecustomize.py                # Windows Python 3.14 Tcl/Tk 兼容处理
├── create_overtime_template.py     # 生成补录与罚班 Word 模板
├── README.md
├── 加班补录空白模板.docx
├── 罚班记录空白模板.docx
├── src/
│   ├── ui/                         # CustomTkinter 页面、弹窗和通用组件
│   ├── modules/                    # 解析、统计、校验与报表生成
│   └── utils/                      # 输出路径、标题和通用工具
├── tests/                          # pytest 自动化测试
├── vendor/tesseract/               # 离线 OCR 运行库、模型和许可证
├── build/pyinstaller/
│   └── TonBanAPP.spec              # PyInstaller 单文件配置
├── data/                           # 运行时本地配置，Git 忽略
├── output/                         # 默认报表输出目录，Git 忽略
└── tonban_exe/                     # 打包结果目录，Git 忽略
```

## 主要模块

| 模块 | 职责 |
| --- | --- |
| `src/ui/main_ui.py` | 主窗口、导航、主题、设置和页面切换 |
| `src/ui/name_list_tab.py` | 总名单导入、搜索和编辑 |
| `src/ui/rules_tab.py` | 毕业季与长期请假规则配置 |
| `src/ui/weekly_tab.py` | 普通周输入、校验、预览和生成 |
| `src/ui/final_weeks_tab.py` | 双周期末排班、预览和合并输出 |
| `src/ui/monthly_tab.py` | 周表汇总、历史结转、调整记录和月报生成 |
| `src/modules/data_manager.py` | `data/config.json` 的加载、迁移和保存 |
| `src/modules/contact_parser.py` | Excel 通讯录解析与毕业季识别 |
| `src/modules/schedule_parser.py` | PDF/Word 排班解析统一入口 |
| `src/modules/pdf_schedule_parser.py` | 文本 PDF、扫描 PDF、网格与 OCR 解析 |
| `src/modules/word_parser.py` | Word 名单、排班和上月月报解析 |
| `src/modules/excel_generator.py` | 周统计计算、校验、Excel 生成和周表聚合 |
| `src/modules/final_weeks_parser.py` | 双周期末 PDF 区块和周次识别 |
| `src/modules/final_weeks_service.py` | 期末周合并统计与原子化输出 |
| `src/modules/word_generator.py` | 月统计计算、排序和 Word 生成 |
| `src/modules/overtime_word.py` | 加班补录、罚班记录 Word 导入导出 |
| `src/modules/ocr_runtime.py` | 定位并调用开发版或打包版 Tesseract |

## 开发与验证

### 语法检查

```powershell
.\.venv\Scripts\python.exe -m compileall -q main.py src create_overtime_template.py sitecustomize.py
```

### 自动化测试

```powershell
.\.venv\Scripts\python.exe -m pytest -q --basetemp .pytest_tmp
```

测试覆盖名单与统计改进、期末周解析/服务/UI 契约、文本型 PDF、扫描型 PDF 和内置 OCR 路径解析。

## 打包

项目使用 `build/pyinstaller/TonBanAPP.spec` 生成 Windows 单文件应用，并将 `vendor/tesseract` 中的 OCR 运行库、简体中文/英文模型及许可证一并打包。

```powershell
cd build\pyinstaller
..\..\.venv\Scripts\python.exe -m PyInstaller TonBanAPP.spec `
  --distpath ..\..\tonban_exe `
  --workpath . `
  --clean `
  --noconfirm
cd ..\..
```

生成文件：

```text
tonban_exe/TonBanAPP.exe
```

注意：

- `data/config.json` 不会被打进 EXE。
- `tonban_exe/`、`output/` 和 PyInstaller 中间文件已被 Git 忽略。
- 内置 Tesseract 及模型的来源、校验值和许可证见 `vendor/tesseract/VENDOR.md`。

## 常见问题

### 导入通讯录失败

确认文件为 `.xlsx`，并包含“姓名”和“专业年级”，或“姓名”“专业”“年级”表头。表头应位于工作表前 20 行。

### 排班 PDF 没有识别到姓名

文本 PDF 需要包含周一至周五列；扫描 PDF 需要清晰、方向正确且表格边框完整。若仍无法识别，可改用兼容的 Word 排班文件，或在生成前预览中人工修正班次。

### Word 排班没有识别某一天

星期标题必须单独成行。`周一` 可以识别，`周一 张三 李四` 不会被当作有效星期标题。

### 无法写入 Excel 或 Word

目标文件可能正被 Excel、Word 或其他程序占用。关闭文件后重新生成；若已存在同名文件，应用会自动选择新的文件名。

### 月统计结果不完整

检查总名单是否为当前名单、周统计 Excel 是否包含必要列、上月 Word 表头是否正确，以及人员姓名是否完全一致。

## 第三方组件

仓库内置的 Tesseract 运行库和训练数据按 Apache-2.0 分发，相关声明保存在：

- `vendor/tesseract/LICENSE-TESSERACT.txt`
- `vendor/tesseract/LICENSE-TESSDATA.txt`
- `vendor/tesseract/README-UPSTREAM.md`
- `vendor/tesseract/VENDOR.md`
