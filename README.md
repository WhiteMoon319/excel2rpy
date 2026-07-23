# Ren'Py Excel 转换工具

双向转换 Ren'Py 脚本 (`.rpy`) 和 Excel 表格 (`.xlsx`)，方便用表格批量编辑对话、场景、分支等。

编剧只需填表 + 下拉选择，无需写任何 Python 语法即可完成变量管理和图片引用。

## 文件

脚本可放在任意子目录（如 `scripts/`），会自动向上查找项目根目录（含 `game/` 的目录）。

```
excel_to_rpy.py      主入口：转换、校验、模板、定义管理和回归测试
rpy_to_excel.py      独立的 .rpy -> Excel 命令行/交互入口
test_roundtrip.py    独立的 roundtrip 回归测试入口
```

日常使用只需要双击 `excel_to_rpy.py`。另外两个脚本保留给命令行调用、单独调试和自动化测试。

### 推荐目录结构

```
project_root/
├── game/               ← Ren'Py 脚本和资源
├── scripts/            ← 本工具脚本（可选，放哪都行）
│   ├── excel_to_rpy.py
│   ├── rpy_to_excel.py
│   └── test_roundtrip.py
└── excel/              ← 表格文件（建议）
```

## 安装

需要 Python 3.10 或更高版本，以及 `openpyxl`：

```bash
pip install openpyxl
```

## 快速开始

**双击 `excel_to_rpy.py`** 即可打开统一交互菜单，不需要记命令。

主菜单包含：

1. Excel 转 `.rpy`
2. 校验 Excel 表格
3. 生成教学模板
4. 生成空白模板
5. 定义管理
6. `.rpy` 转 Excel
7. 运行自带回归测试
8. 退出

`rpy_to_excel.py` 和 `test_roundtrip.py` 仍可单独双击运行。

命令行也支持：

```bash
# 生成教学模板（含示例和说明）
python excel_to_rpy.py --template

# 生成空白模板（仅表头+下拉）
python excel_to_rpy.py --blank

# 自定义模板路径（对 --template 和 --blank 都适用）
python excel_to_rpy.py --template --template-path my_template.xlsx

# Excel -> .rpy（自动先校验，有错误也会继续转换）
python excel_to_rpy.py script.xlsx -o output.rpy

# 仅校验不转换
python excel_to_rpy.py script.xlsx --check

# 定义管理 — 重建 defines.rpy（全量扫描源文件）
python excel_to_rpy.py --rebuild-defines

# 重建前备份旧文件
python excel_to_rpy.py --rebuild-defines --backup

# 重建但不清理源文件中的定义行
python excel_to_rpy.py --rebuild-defines --no-cleanup

# 仅预览重建结果，不实际写入
python excel_to_rpy.py --rebuild-defines --dry-run

# .rpy -> Excel
python rpy_to_excel.py script.rpy -o output.xlsx

# .rpy -> Excel（不拆分 Sheet，全部放到一个表）
python rpy_to_excel.py script.rpy -o output.xlsx --no-split

# 测试指定脚本的双向转换
python test_roundtrip.py test.rpy

# 测试仓库自带样本
python test_roundtrip.py test_roundtrip.rpy
```

## Excel 表格格式（12 列）

| 列 | 说明 | 示例 |
|----|------|------|
| A 场景标签 | `label` 的名称 | `start` |
| B 指令类型 | 下拉选择（中文标注）/ 留空自动推断 | `dialogue（角色对话）` |
| C 图片/背景 | 图片路径或函数调用 | `images/bg_room.jpg` |
| D 角色名 | 说话的角色（dialogue/show/hide/角色定义） | `主角` |
| E 变量名 | 变量名（变量定义/赋值/增减/开关/条件比较） | `score` |
| F 逻辑连接 | 复合条件的连接符，下拉 `and` / `or` / 空 | `and` |
| G 对话文本 | 对话内容或数值 | `你好！` |
| H 选项文本 | 菜单选项文字 | `去图书馆` |
| I 跳转目标 | `jump`/`call` 的目标 / `return`、`$` | `library` |
| J 音频路径 | BGM、音效、配音文件路径 | `audio/bgm.ogg` |
| K 位置/特效/属性 | 中文位置和转场 | `溶解`、`左边` |
| L 备注 | 额外说明 | `centered` |

## 结构化指令类型

### 定义类

| 指令类型 | 填写方式 | 生成代码 |
|----------|----------|----------|
| `角色定义（定义角色）` | D列=角色名, G列=Character(...) | `define name = Character(...)` |
| `变量定义（定义变量）` | E列=变量名, G列=初始值 | `default name = value` |
| `图片定义（定义图片）` | C列=路径, E列=标识名 | `image name = "path"` |

### 操作类

| 指令类型 | 填写方式 | 生成代码 |
|----------|----------|----------|
| `变量赋值（变量赋值）` | E列=变量名, G列=新值 | `$ name = value` |
| `变量增减（变量增减）` | E列=变量名, G列=增量（如1或-3） | `$ name += 1` / `$ name -= 3` |
| `变量开关（变量开关）` | E列=变量名, G列=下拉选true/false | `$ name = True` / `$ name = False` |

### 条件类（结构化）

| 指令类型 | 填写方式 | 生成代码 |
|----------|----------|----------|
| `变量=（等于）` | E列=变量名, F列=and/or/空, G列=比较值 | `if name == value:` |
| `变量≠（不等于）` | E列=变量名, F列=and/or/空, G列=比较值 | `if name != value:` |
| `变量>（大于）` | E列=变量名, F列=and/or/空, G列=比较值 | `if name > value:` |
| `变量≥（大于等于）` | E列=变量名, F列=and/or/空, G列=比较值 | `if name >= value:` |
| `变量<（小于）` | E列=变量名, F列=and/or/空, G列=比较值 | `if name < value:` |
| `变量≤（小于等于）` | E列=变量名, F列=and/or/空, G列=比较值 | `if name <= value:` |
| `else（否则）` | 无需额外填写 | `else:` |

**复合条件**：F 列（逻辑连接）用于串联多个结构化条件。

```
第1行: 变量≥ | score | and | 10      → if score >= 10
第2行: 变量< | score |     | 50      → and score < 50:
```
→ 生成 `if score >= 10 and score < 50:`

- **F 列非空** → 该行是复合条件的一部分，继续拼接
- **F 列为空** → 结束条件链，加冒号换行
- **F 列始终为空** → 独立单条件

> 复杂条件仍可使用传统的 `if（条件判断）` / `elif（否则如果）`，在 G 列写完整 Python 表达式。

### 保留的传统指令

`dialogue`、`scene`、`show`、`hide`、`menu`、`menu_option`、`jump`、`call`、`return`、
`play_music`、`stop_music`、`queue_music`、`play_sound`、`stop_sound`、
`voice`、`pause`、`player_input`、`window`、`$（设置变量）`、`if（条件判断）`、`elif（否则如果）`

使用传统指令时，`$` / `if` / `elif` 的代码填在 G 列（对话文本）。

## 下拉菜单

| 列 | 下拉数据来源 |
|----|-------------|
| B 指令类型 | 固定下拉列表（所有支持的指令） |
| C 图片/背景 | 扫描 `game/*.rpy` 中的 `image` 定义名 + `game/` 下所有图片文件 |
| D 角色名 | 扫描 `game/*.rpy` 中的 `define` 角色 + 表中定义的角色 |
| E 变量名 | 扫描 `game/*.rpy` 中的 `default` 变量 + 表中定义的变量 |
| F 逻辑连接 | 固定 `and` / `or` 下拉（用于复合条件串联） |
| G 对话文本 | 当指令=变量开关时，固定 `true` / `false` 下拉 |

> 每次生成模板时，工具会从脚本所在目录向上查找含 `game/` 的项目根目录，自动扫描并生成对应的下拉列表。

## 特色功能

### 智能填充

- **角色名前向填充**：角色名列留空自动沿用上一行
- **变量名前向填充**：变量名列留空自动沿用上一行
- **指令自动推断**：有角色名 + 空指令 → 自动当对话；有文本 + 无角色 → 自动当旁白
- **旁白快捷**：角色名写 `旁白` 自动切换 narrator

### 中文术语映射

位置和转场支持中文输入，自动翻译为 Ren'Py 英文：

```
溶解→dissolve  褪色→fade  闪白→Fade(0.1,0,0.5)
左边→left     右边→right  中间→center
像素化→pixellate  横向振动→hpunch  纵向振动→vpunch
```

### 表格校验

转换前自动扫描并报告问题：

| 严重度 | 检查项 |
|--------|--------|
| ERROR | label 无名称、menu 无选项、menu_option 位置错误、重复标签 |
| WARN | 跳转目标未定义、缺角色名/对话文本、show 无图片、变量开关值非 true/false、变量定义/赋值缺少值 |

### 角色自动注册

无需手动写 `define character`——脚本会扫描全表并收集出现的角色名，自动生成 `define` 语句。

- 显式 `define_character` 行：D 列填变量名，G 列可填完整 `Character("显示名", ...)`；G 列为空则自动生成 `Character("变量名")`
- `show` 指令中的角色名（非 bg 开头）也会自动注册
- 自动注册的角色默认 `define 角色名 = Character("角色名")`，需要中文显示名请用显式定义
- 检测到 Ren'Py 工程时，定义追加到 `game/defines.rpy`；独立使用时，定义写在生成的 `.rpy` 顶部

### 旁白扩展

L 列（备注）可用于控制旁白格式：

| 备注值 | 效果 |
|--------|------|
| `centered` | 生成 `centered "文本"`（居中显示） |
| `explicit` | 生成 `narrator "文本"`（显式 narrator） |
| 留空 | 生成普通 `"文本"` |

### 玩家输入

`player_input` 指令的 K 列（位置/特效）可填写**默认值**，生成 `$ var = renpy.input("提示").strip() or "默认值"`。

### 保存容错

当保存目标目录被安全软件（如 Windows Defender 的勒索软件防护）阻止时，脚本会自动尝试备选路径：
目标目录 → 桌面 → 文档 → 系统临时目录。Excel 生成重试 3 次后才走备选链。
如果 `game/defines.rpy` 无法写入，定义会回退到本次生成的 `.rpy`，不会静默丢失。

### 变量增减

G 列填增量时支持显式正负号：`+1` → `$ var += 1`，`-3` → `$ var -= 3`；不带符号的正数也视为 `+=`。

### 多 Sheet 支持

每个 Sheet 视为一个独立场景。首个 Sheet 定义为 `start`，其他 Sheet 以标签名或 Sheet 名为入口。

## 双向转换（Roundtrip）

```
.rpy  ──rpy_to_excel──>  .xlsx  ──excel_to_rpy──>  .rpy
```

反向转换自动识别结构化指令（`variable_set`、`variable_change`、`variable_toggle`、`variable_eq` 等），在工具支持的指令子集内保持双向转换一致。

- 复合条件 `if a > 1 and b < 5:` 反向解析为**多行结构化条件**，自动填入 F 列逻辑连接符
- 转场和位置名称自动**英译中**（`dissolve`→`溶解`、`left`→`左边`）
- 反向生成的 Excel 同样自动填充角色、变量、图片下拉列表

主菜单的 **7. 运行回归测试** 会测试仓库自带样本。测试在系统临时目录中完成，不会修改真实工程的 `defines.rpy`，临时文件会在结束后自动清理。

当前 roundtrip 面向本工具列出的结构化指令，不是完整的 Ren'Py 语法解析器。ATL、screen、translate、复杂 Python 块、深层嵌套和含复杂转义的文本，建议保留在人工维护的 `.rpy` 中。

## 定义管理（Define Manager）

Excel 可以作为策划侧的主数据源；检测到 Ren'Py 工程时，工具可以把定义集中维护在 `game/defines.rpy`，其他 `.rpy` 只保留剧情逻辑。

### 自动追加

在检测到 Ren'Py 工程（包含 `game/` 目录）时，每次转换会从 Excel 提取 `define_character` / `define_variable` / `define_image`，并把尚不存在的定义追加到 `game/defines.rpy`。

- 普通转换只追加新名称，同名定义会跳过
- 普通转换保留 `defines.rpy` 中已有的注释和自定义内容
- 没有 `game/` 目录时，定义保留在生成的 `.rpy` 顶部
- `defines.rpy` 写入失败时，也会把定义回退到生成的 `.rpy`

### 交互菜单

主菜单选 **5. 定义管理** → 子菜单：

- **1. 追加新定义**：指定 Excel，提取其中的新定义追加到 `defines.rpy`
- **2. 重建全部定义**：扫描项目所有 Excel 和 `.rpy`（排除 `defines.rpy`），全量重建 `defines.rpy`

### 重建选项

选择重建时会依次询问：

| 选项 | 默认 | 说明 |
|------|------|------|
| 是否备份旧 defines.rpy？ | **N** | 备份为 `defines.rpy.bak` |
| 是否清理源文件中的定义行？ | **N** | 选择 Y 后删除 Excel 和 `.rpy` 中的定义行 |
| 是否预览变更（dry-run）？ | **N** | 仅打印输出，不写入、不清理 |

命令行执行 `--rebuild-defines` 时默认会清理源定义；需要保留源文件时必须添加 `--no-cleanup`。第一次重建建议先使用 `--dry-run`，确认结果后再执行，并配合 `--backup`。

### 定义整理规则

写入 `defines.rpy` 时，定义按**角色归属**分组：

```renpy
# ======================================
# 庄桂清（zhuang）
# ======================================

define zhuang = Character("庄桂清")
default love_zhuang = 0
default trust_zhuang = 0
image zhuang happy = "zhuang_happy.png"

# ======================================
# 玩家
# ======================================

default player_name = ""
default score_player = 0
default pc_health = 100

# ======================================
# 背景
# ======================================

image bg classroom = "bg_classroom.jpg"
image bg room_zhuang = "bg_room_zhuang.jpg"

# ======================================
# 系统/全局定义
# ======================================

default has_key = False
default money = 100
```

**分组优先级**：
1. **分类前缀**（最高）：首词为 `bg` 的定义强制归入"背景"组，不参与角色匹配
2. **角色匹配**：定义名中含已定义角色名（如 `love_zhuang` → 庄桂清）
3. **虚拟前缀**：`player_` / `pc_` / `mc_` / `主角` 前缀 → "玩家"组（前后缀均支持，如 `score_player`）
4. **系统/全局**：其余无法归类的定义

> Ren'Py 框架文件（`screens.rpy`、`gui.rpy`、`options.rpy`）中的 `default` 定义会被自动排除，不会吸入 `defines.rpy`。

### 去重规则

- **普通转换/追加模式**：同名定义已存在则跳过，保留原文件其他内容
- **重建模式**：同名定义以 **Excel 版本为准**，`defines.rpy` 会被重新生成；工具会输出冲突报告

## 注意事项

- 图片路径可用双引号包裹（`"images/bg.jpg"`），函数调用（`Transform(...)`）直接写，脚本自动识别括号区分
- 空行用于分隔场景、结束 `if` / `menu` 块
- `python:` 多行代码块会拆分为独立语句和 `if` 块，功能等价
- 运行时脚本会自动向上查找含 `game/` 的项目根目录，无需指定路径
- 双击运行时若 CWD 异常（如 `C:\Windows\System32`），脚本自动切换到自身目录
- 结构化条件 `变量=`、`变量≥` 等只在条件为简单 `if var OP val` 时有效；复杂条件请用传统 `if`
- 普通转换不会覆盖同名定义；需要用 Excel 版本重建全部定义时，请明确使用“定义管理 → 重建全部定义”
