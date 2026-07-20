# Ren'Py Excel 转换工具

双向转换 Ren'Py 脚本 (`.rpy`) 和 Excel 表格 (`.xlsx`)，方便用表格批量编辑对话、场景、分支等。

## 文件

```
excel_to_rpy.py    Excel -> .rpy （正向转换）
rpy_to_excel.py    .rpy -> Excel （反向转换）
test_roundtrip.py  回归测试（改代码后跑一下）
```

## 安装

```bash
pip install openpyxl
```

## 快速开始

**双击任一 `.py` 文件** → 交互菜单。

| 脚本 | 做什么 |
|------|--------|
| `excel_to_rpy.py` | 1. 转换 / 2. 校验 / 3. 教学模板 / 4. 空白模板 |
| `rpy_to_excel.py` | 1. 转换 .rpy 为 Excel |
| `test_roundtrip.py` | 1. 测试自带样本 / 2. 测试游戏脚本 / 3. 测试其他 |

命令行也支持：

```bash
# 生成模板
python excel_to_rpy.py --template

# Excel -> .rpy（自动先校验）
python excel_to_rpy.py script.xlsx -o output.rpy

# 仅校验不转换
python excel_to_rpy.py script.xlsx --check

# .rpy -> Excel
python rpy_to_excel.py script.rpy -o output.xlsx

# 回归测试
python test_roundtrip.py test.rpy
```

## Excel 表格格式

| 列 | 说明 | 示例 |
|----|------|------|
| 场景标签 | `label` 的名称 | `start` |
| 指令类型 | 下拉选择（中文标注）/ 留空自动推断 | `dialogue（角色对话）` |
| 图片/背景 | 图片路径或函数调用 | `images/bg.jpg` |
| 角色名 | 说话的角色 / 留空沿用上一行 | `eileen` |
| 对话文本 | 对话内容或 Python 代码 | `你好！` |
| 选项文本 | 菜单选项文字 | `去图书馆` |
| 跳转目标 | `jump`/`call` 的目标 / 也支持 `return`、`$` | `library` |
| 音频路径 | BGM、音效、配音文件路径 | `audio/bgm.ogg` |
| 位置/特效/属性 | 支持中文 `溶解`→`dissolve`，`左边`→`left` | `at center with fade` |
| 备注 | 额外说明 | `centered` |

## 特色功能

### 智能填充

- **角色名前向填充**：角色名列留空自动沿用上一行，不用每行重复填
- **指令自动推断**：有角色名 + 空指令 → 自动当对话；有文本 + 无角色 → 自动当旁白
- **旁白快捷**：角色名写 `旁白` 自动切换 narrator

### 中文术语映射

位置和转场支持中文输入，自动翻译为 Ren'Py 英文：

```
溶解→dissolve  褪色→fade  闪白→Fade(0.1,0,0.5)
左边→left     右边→right  中间→center
像素化→pixellate  振屏→hpunch
```

### 表格校验

转换前自动扫描并报告问题：

| 严重度 | 检查项 |
|--------|--------|
| ERROR | label 无名称、menu 无选项、menu_option 位置错误、重复标签 |
| WARN | 跳转目标未定义、缺角色名/对话文本、show 无图片、音频路径缺失 |

校验通过命令行 `--check` 可独立运行，转换时也会自动执行，有错误会询问是否继续。

### 角色自动注册

无需手动写 `define character`——脚本自动扫描全表，收集所有出现的角色名，在输出的 `.rpy` 顶部生成 `define` 语句。

### 多 Sheet 支持

每个 Sheet 视为一个独立场景。首个 Sheet 为 `start`，其他 Sheet 以标签名或 Sheet 名为入口。

## 双向转换（Roundtrip）

```
.rpy  ──rpy_to_excel──>  .xlsx  ──excel_to_rpy──>  .rpy
```

经过回归测试验证，`.rpy` → Excel → `.rpy` 的输出与原始脚本内容完全一致。

## 注意事项

- 图片路径可用双引号包裹（`"images/bg.jpg"`），函数调用（`Transform(...)`）直接写
- 空行用于分隔场景、结束 `if` / `menu` 块
- `python:` 多行代码块会拆分为独立 `$` 语句和 `if` 块，功能等价
- 双击运行时若 CWD 异常（如 `C:\Windows\System32`），脚本自动切换到自身目录
- 拖错文件类型（.xlsx 拖给 rpy_to_excel）会友好提示
