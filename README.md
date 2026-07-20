# Ren'Py Excel 转换工具

双向转换 Ren'Py 脚本 (`.rpy`) 和 Excel 表格 (`.xlsx`)，方便用表格批量编辑对话、场景、分支等。

## 文件

```
excel_to_rpy.py    Excel -> .rpy （正向转换，表格生成脚本）
rpy_to_excel.py    .rpy -> Excel （反向转换，脚本还原表格）
```

## 安装

```bash
pip install openpyxl
```

## 快速开始

**双击 `excel_to_rpy.py`** → 交互菜单，可选：

- 生成教学模板（含示例数据和中文教程）
- 生成空白模板（仅表头）
- 转换写好的 Excel 为 `.rpy`

**双击 `rpy_to_excel.py`** → 把已有的 `.rpy` 脚本反向解析成 Excel。

命令行也支持：

```bash
# 生成模板
python excel_to_rpy.py --template

# Excel -> .rpy
python excel_to_rpy.py script.xlsx -o output.rpy

# .rpy -> Excel
python rpy_to_excel.py script.rpy -o output.xlsx
```

## Excel 表格格式

| 列 | 说明 | 示例 |
|----|------|------|
| 场景标签 | `label` 的名称 | `start` |
| 指令类型 | 下拉选择（中文标注） | `dialogue（角色对话）` |
| 图片/背景 | 图片路径或函数调用 | `images/bg.jpg` |
| 角色名 | 说话的角色或变量名 | `eileen` |
| 对话文本 | 对话内容或 Python 代码 | `你好！` |
| 选项文本 | 菜单选项文字 | `去图书馆` |
| 跳转目标 | `jump`/`call` 的目标标签 | `library` |
| 音频路径 | BGM、音效、配音文件路径 | `audio/bgm.ogg` |
| 位置/特效/属性 | `at left`、`with dissolve` 等 | `at center with fade` |
| 备注 | 额外说明或 `centered` 标记 | `centered` |

如果对应单元格不需要填，留空即可。

## 双向转换（Roundtrip）

```
.rpy  ──rpy_to_excel──>  .xlsx  ──excel_to_rpy──>  .rpy
```

经过反复验证，`.rpy` → Excel → `.rpy` 的输出与原始脚本功能完全一致。

## 注意事项

- 表格中的图片和音频路径需以双引号包裹，函数调用（如 `Transform(...)`）则直接写
- 空行用于分隔场景和结束 `if` / `menu` 块
- `python:` 多行代码块会拆分为独立 `$` 语句和 `if` 块，功能等价
- 双击运行时若提示权限错误，脚本会自动将工作目录从系统目录切换到脚本所在目录
