#!/usr/bin/env python3
"""
Ren'Py Excel 转 .rpy 脚本 — 新手友好版

用法:
  python excel_to_rpy.py --template         生成 Excel 模板文件
  python excel_to_rpy.py input.xlsx         转换 Excel 为 .rpy
  python excel_to_rpy.py input.xlsx -o out  指定输出路径
  双击打开 → 交互菜单模式
"""

import argparse
import os
import sys
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("请先安装 openpyxl：pip install openpyxl")
    sys.exit(1)

# ── 模板配置 ──────────────────────────────────────────────

HEADERS = [
    ("场景标签", 16),
    ("指令类型", 18),
    ("图片/背景", 24),
    ("角色名", 14),
    ("对话文本", 40),
    ("选项文本", 20),
    ("跳转目标", 16),
    ("音频路径", 24),
    ("位置/特效/属性", 22),
    ("备注", 20),
]

# 指令类型下拉选项（中文翻译方便新手，解析时取"（"前的英文部分）
COMMAND_TYPES = [
    "label（场景标签）",
    "scene（背景图）",
    "show（显示角色）",
    "hide（隐藏角色）",
    "dialogue（角色对话）",
    "narrator（旁白）",
    "menu（选择菜单）",
    "menu_option（菜单选项）",
    "jump（跳转）",
    "call（调用子场景）",
    "return（返回）",
    "$（设置变量）",
    "if（条件判断）",
    "elif（否则如果）",
    "else（否则）",
    "play_music（播放BGM）",
    "stop_music（停止BGM）",
    "queue_music（排队播BGM）",
    "play_sound（播放音效）",
    "stop_sound（停止音效）",
    "voice（配音）",
    "pause（暂停等待）",
    "player_input（玩家输入）",
    "window（对话框开关）",
    "define_character（定义角色）",
    "default（默认变量）",
    "image（定义图片）",
]

# 样式
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
CELL_FONT = Font(name="微软雅黑", size=10)
CELL_ALIGN = Alignment(vertical="top", wrap_text=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="B4C6E7"),
    right=Side(style="thin", color="B4C6E7"),
    top=Side(style="thin", color="B4C6E7"),
    bottom=Side(style="thin", color="B4C6E7"),
)
ALT_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")


def generate_template(output_path: str):
    """生成 Excel 模板文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "script"

    # 写表头
    for col_idx, (header, width) in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 24

    # 指令类型下拉
    from openpyxl.worksheet.datavalidation import DataValidation

    cmd_formula = '"' + ",".join(COMMAND_TYPES) + '"'
    dv = DataValidation(type="list", formula1=cmd_formula, allow_blank=True)
    dv.error = "请从下拉列表选择指令类型"
    dv.errorTitle = "无效指令"
    dv.prompt = "请选择指令类型（括号内是中文说明）"
    dv.promptTitle = "指令类型"
    dv.add("B2:B10000")
    ws.add_data_validation(dv)

    # 预填示例数据（带变量和条件分支）
    sample_data = [
        ["start", "label（场景标签）", "", "", "", "", "", "", "", "游戏开始"],
        ["", "scene（背景图）", "images/bg_classroom.jpg", "", "", "", "", "", "with dissolve", "教室背景"],
        ["", "play_music（播放BGM）", "", "", "", "", "", "audio/bgm_happy.ogg", "", ""],
        ["", "narrator（旁白）", "", "", "你醒来发现自己在一间教室里……", "", "", "", "", ""],
        ["", "player_input（玩家输入）", "", "player_name", "请输入你的名字：", "", "", "", "无名", "让玩家输入主角名"],
        ["", "show（显示角色）", "", "eileen", "", "", "", "", "at left", "角色出现在左边"],
        ["", "dialogue（角色对话）", "", "eileen", "你好，[player_name]！欢迎来到这个世界！", "", "", "", "", "用 [变量名] 引用输入"],
        ["", "dialogue（角色对话）", "", "eileen", "你今天想做什么呢？", "", "", "", "", ""],
        ["", "menu（选择菜单）", "", "", "", "", "", "", "", "玩家做选择"],
        ["", "menu_option（菜单选项）", "", "", "", "去图书馆 📚", "library", "", "", ""],
        ["", "menu_option（菜单选项）", "", "", "", "去操场 🏃", "playground", "", "", ""],
        ["", "menu_option（菜单选项）", "", "", "", "回家 🏠", "go_home", "", "", ""],
        ["", "", "", "", "", "", "", "", "", "", ""],
        # 图书馆分支：展示变量和条件
        ["library", "label（场景标签）", "", "", "", "", "", "", "", "图书馆场景"],
        ["", "scene（背景图）", "images/bg_library.jpg", "", "", "", "", "", "", ""],
        ["", "dialogue（角色对话）", "", "eileen", "这里好安静啊。", "", "", "", "", ""],
        ["", "dialogue（角色对话）", "", "eileen", "你找到了一本好书，知识增加了！", "", "", "", "", ""],
        ["", "$（设置变量）", "", "", "knowledge += 1", "", "", "", "", "变量 knowledge 加 1"],
        ["", "dialogue（角色对话）", "", "eileen", "你还想继续探索吗？", "", "", "", "", ""],
        ["", "if（条件判断）", "", "", "knowledge >= 2", "", "", "", "", "如果知识≥2 → 分支"],
        ["", "dialogue（角色对话）", "", "eileen", "你已经读了不少书了！真棒！", "", "", "", "", "条件成立时执行"],
        ["", "jump（跳转）", "", "", "", "", "start", "", "", ""],
        ["", "else（否则）", "", "", "", "", "", "", "", "条件不成立时"],
        ["", "dialogue（角色对话）", "", "eileen", "再多读一本书吧~", "", "", "", "", "条件不成立时执行"],
        ["", "jump（跳转）", "", "", "", "", "library", "", "", ""],
        ["", "", "", "", "", "", "", "", "", "", ""],
        # 操场分支
        ["playground", "label（场景标签）", "", "", "", "", "", "", "", "操场场景"],
        ["", "scene（背景图）", "images/bg_playground.jpg", "", "", "", "", "", "", ""],
        ["", "dialogue（角色对话）", "", "eileen", "天气真好！适合运动！", "", "", "", "", ""],
        ["", "show（显示角色）", "", "eileen", "", "", "", "", "happy smile", "切换为高兴表情"],
        ["", "dialogue（角色对话）", "", "eileen", "跑起来好舒服！", "", "", "", "", ""],
        ["", "$（设置变量）", "", "", "health += 1", "", "", "", "", "设置变量"],
        ["", "return（返回）", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", "", ""],
        # 黑屏大字示例：常用于章节标题、日期地点提示
        ["chapter_title", "label（场景标签）", "", "", "", "", "", "", "", "黑屏大字演示"],
        ["", "window（对话框开关）", "", "", "hide", "", "", "", "", "隐藏对话框，纯黑屏"],
        ["", "scene（背景图）", "black", "", "", "", "", "", "", "black=纯黑背景"],
        ["", "narrator（旁白）", "", "", "第一章", "", "", "", "", "旁白=居中大字"],
        ["", "pause（暂停等待）", "", "", "2.0", "", "", "", "", "停留2秒"],
        ["", "narrator（旁白）", "", "", "启程之日", "", "", "", "", ""],
        ["", "pause（暂停等待）", "", "", "2.0", "", "", "", "", ""],
        ["", "window（对话框开关）", "", "", "show", "", "", "", "", "恢复对话框显示"],
        ["", "jump（跳转）", "", "", "", "", "start", "", "", ""],
        ["", "", "", "", "", "", "", "", "", "", ""],
        # 回家分支
        ["go_home", "label（场景标签）", "", "", "", "", "", "", "", "回家场景"],
        ["", "scene（背景图）", "images/bg_home.jpg", "", "", "", "", "", "", ""],
        ["", "voice（配音）", "", "", "", "", "", "audio/voice/eileen_home.ogg", "", "配音紧跟的对话"],
        ["", "dialogue（角色对话）", "", "eileen", "回家休息啦~", "", "", "", "", ""],
        ["", "return（返回）", "", "", "", "", "", "", "", ""],
    ]

    for row_idx, row_data in enumerate(sample_data, start=2):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = CELL_FONT
            cell.alignment = CELL_ALIGN
            cell.border = THIN_BORDER
            if row_idx % 2 == 0:
                cell.fill = ALT_FILL
        ws.row_dimensions[row_idx].height = 20

    # 冻结首行
    ws.freeze_panes = "A2"

    # ── 使用说明 Sheet ──
    ws_help = wb.create_sheet("使用说明")
    help_lines = [
        "═══════════════════════════════════════════",
        "  Ren'Py Excel → .rpy 转换工具 · 新手使用指南",
        "═══════════════════════════════════════════",
        "",
        "【一】指令类型速查表",
        "───────────────────────────────────────────",
        "  label（场景标签）     → 开始一个新场景，在「场景标签」列写标签名（如 start）",
        "  scene（背景图）       → 切换背景图，在「图片/背景」列填图片路径",
        "  show（显示角色）      → 显示角色立绘，填「角色名」+「图片/背景」+「位置/特效」",
        "  hide（隐藏角色）      → 隐藏某个角色立绘，填「角色名」",
        "  dialogue（角色对话）  → 角色说话，填「角色名」+「对话文本」",
        "  narrator（旁白）      → 画外音/旁白，只填「对话文本」",
        "  menu（选择菜单）      → 弹出选项给玩家选，后面必须跟 menu_option 行",
        "  menu_option（菜单选项）→ 一个选项，填「选项文本」+「跳转目标」（跳到哪个场景）",
        "  jump（跳转）          → 跳到指定场景，填「跳转目标」",
        "  call（调用子场景）    → 调用子场景（可 return 返回），填「跳转目标」",
        "  return（返回）        → 返回/结束当前场景",
        "  $（设置变量）         → 修改变量值，在「对话文本」列写代码（如 score += 10）",
        "  if（条件判断）        → 条件分支开始，在「对话文本」列写条件（如 score >= 50）",
        "  elif（否则如果）      → 上一个 if 不成立时的备选条件",
        "  else（否则）          → 所有条件都不成立时",
        "  play_music（播放BGM） → 播放背景音乐，「音频路径」",
        "  stop_music（停止BGM） → 停止背景音乐",
        "  queue_music（排队播BGM）→ 当前BGM播完后自动切换，「音频路径」",
        "  play_sound（播放音效）→ 播放音效，「音频路径」",
        "  stop_sound（停止音效）→ 停止正在播放的音效",
        "  voice（配音）         → 播放语音（对话前使用），「音频路径」",
        "  pause（暂停等待）     → 等待若干秒，「对话文本」列填秒数（如 2.5）",
        "  player_input（玩家输入）→ 弹出输入框让玩家打字，填「角色名」（变量名）+「对话文本」（提示）",
        "  window（对话框开关）  → 控制对话框显示/隐藏，「对话文本」填 show/hide/auto",
        "",
        "【二】位置/特效列 · 常用值",
        "───────────────────────────────────────────",
        "  角色的位置（show/hide 时使用）：",
        "    at left             → 屏幕左侧",
        "    at right            → 屏幕右侧",
        "    at center           → 屏幕正中",
        "    at offscreenleft    → 屏幕左侧外面（配合 move 移入）",
        "    at offscreenright   → 屏幕右侧外面",
        "    at truecenter       → 绝对居中",
        "",
        "  转场特效（scene/show 切换画面时使用）：",
        "    with dissolve       → 淡入淡出（最常用）",
        "    with fade           → 黑屏渐入",
        "    with pixellate      → 像素化过渡",
        "    with moveinleft     → 从左边滑入",
        "    with moveinright    → 从右边滑入",
        "    with hpunch         → 画面震动（打击感）",
        "    with vpunch         → 画面上下震动",
        "",
        "【三】变量和条件 · 怎么写",
        "───────────────────────────────────────────",
        "  1. 先定义变量（在 Ren'Py 的 script.rpy 或任意 .rpy 开头）：",
        "       default knowledge = 0",
        "       default health = 0",
        "",
        "  2. 用 $（设置变量）修改：",
        "       指令：$（设置变量）",
        "       对话文本：knowledge += 1          ← 知识+1",
        "       对话文本：score = score + 10      ← 分数+10",
        "       对话文本：has_key = True          ← 获得钥匙",
        "",
        "  3. 用 if（条件判断）做分支：",
        "       指令：if（条件判断）",
        "       对话文本：knowledge >= 2           ← 如果知识≥2",
        "       （下面紧跟条件成立时要执行的指令，会自动缩进）",
        "",
        "       指令：else（否则）                  ← 条件不成立时",
        "       （下面跟备选指令）",
        "",
        "  常用条件写法：",
        "    score >= 60           → 分数大于等于60（及格线）",
        "    has_key == True       → 拥有钥匙",
        "    love >= 80            → 好感度≥80",
        "    item_count > 0        → 拥有至少1个物品",
        "",
        "【四】选项后不同对话 · 两种写法",
        "───────────────────────────────────────────",
        "  写法A（推荐，结构清晰）：选项→跳转独立场景→在场景里写对话",
        "    见模板示例：menu_option「去图书馆」→ jump 到 library 场景",
        "    library 场景下有独立对话",
        "",
        "  写法B（简单，适合短分支）：用变量+if 判断",
        "    菜单选项设置变量 → 同一个场景内用 if 判断显示不同对话",
        "    示例：",
        "      label choice_result:",
        "          if choice == \"library\":",
        "              e \"你去图书馆了。\"",
        "          elif choice == \"playground\":",
        "              e \"你去操场了。\"",
        "",
        "【五】填写注意事项",
        "───────────────────────────────────────────",
        "  1. 每行一个指令，按剧情顺序从上到下填写",
        "  2. 「场景标签」列只在 label 行填写，其他行留空即可",
        "  3. menu 后紧跟 menu_option，直到遇到其他指令或空行",
        "  4. if 后紧跟条件成立的指令，直到 elif/else 或空行",
        "  5. 空行用来分隔不同场景区块，会被跳过",
        "  6. 图片/音频路径是相对于 Ren'Py 项目的 game/ 目录",
        "  7. 对话文本中不要手动输入双引号，脚本会自动加",
        "  8. 选项文本、角色名、对话支持 emoji（如 📚🏃🏠）",
        "",
        "【六】黑屏 + 居中大字 · 怎么写（章节标题 / 日期地点）",
        "───────────────────────────────────────────",
        "  在 Ren'Py 中显示「黑底白字居中大字」非常简单：",
        "",
        "  scene（背景图） ：「图片/背景」填 black  →  画面变纯黑",
        "  narrator（旁白）：「对话文本」填你的大字内容  →  自动居中大字显示",
        "  pause（暂停等待）：「对话文本」填秒数          →  停留几秒再继续",
        "",
        "  完整示例（Excel 中的几行）：",
        "    scene（背景图）    → 图片/背景：black",
        "    narrator（旁白）   → 对话文本：第一章",
        "    pause（暂停等待）  → 对话文本：2.0",
        "    narrator（旁白）   → 对话文本：启程之日",
        "    pause（暂停等待）  → 对话文本：2.0",
        "    jump（跳转）       → 跳转目标：start",
        "",
        "  生成的 .rpy：",
        "    scene black",
        '    "第一章"',
        "    pause 2.0",
        '    "启程之日"',
        "    pause 2.0",
        "    jump start",
        "",
        "【七】关于逐字显示（打字机效果）",
        "───────────────────────────────────────────",
        "  Ren'Py 默认就是逐字显示文本的，你不需要任何额外设置。",
        "  脚本生成的角色对话、旁白，运行时都会自动逐字打出。",
        "",
        "  如果想控制速度，在「对话文本」中使用 {cps} 标签：",
        '    原文："你好，欢迎来到这个世界！"',
        '    慢速："{cps=20}你好，欢迎来到这个世界！{/cps}"',
        "    cps = 每秒显示的字符数，默认约 40~50",
        "",
        "  常用文本标签（直接写在对话文本里即可）：",
        "    {cps=30}文字{/cps}   → 指定逐字速度",
        "    {w}                  → 等待玩家点击再继续显示",
        "    {w=1.5}              → 等待 1.5 秒自动继续",
        "    {p}                  → 段落停顿（等同于换行+等待点击）",
        "    {b}文字{/b}          → 加粗",
        "    {i}文字{/i}          → 斜体",
        "    {size=+10}文字{/size}→ 放大字号",
        "    {color=#ff0000}文字{/color} → 改变颜色",
        "",
        "【八】让玩家输入名字 · player_input 指令",
        "───────────────────────────────────────────",
        "  player_input（玩家输入）可以弹出输入框让玩家打字：",
        "    「角色名」列 → 填变量名（如 player_name）",
        "    「对话文本」列 → 填提示文字（如 请输入你的名字）",
        "    「位置/特效」列 → 可选默认值（玩家不输入时使用，如 无名）",
        "",
        "  示例 Excel 行：",
        "    player_input（玩家输入）→ 角色名：player_name",
        "                            → 对话文本：请输入你的名字：",
        "                            → 位置/特效：无名（默认值）",
        "",
        "  生成的 .rpy：",
        "    $ player_name = renpy.input(\"请输入你的名字：\").strip() or \"无名\"",
        "",
        "【九】在对话中引用变量 · [变量名] 语法",
        "───────────────────────────────────────────",
        "  对话文本中使用方括号 [ ] 包裹变量名即可自动显示变量值：",
        '    "你好，[player_name]！"           → 显示：你好，小明！',
        '    "你的分数是 [score] 分。"         → 显示：你的分数是 85 分。',
        '    "[player_name]，你的好感度是 [love]。"',
        "",
        "  注意：[ ] 里的变量必须事先通过 default / $ / player_input 定义过。",
        "  这个功能是 Ren'Py 自带的「文本插值」，脚本无需特殊处理。",
        "",
        "【十】角色表情差分 · LayeredImage 部件拼合",
        "───────────────────────────────────────────",
        "  如果你的角色由多个部件拼成（本体+眼睛+嘴），",
        "  推荐使用 Ren'Py 的 LayeredImage。",
        "",
        "  ★ 第一步：在 Ren'Py 项目的 .rpy 文件中定义 LayeredImage",
        "  （建议写在单独文件如 characters.rpy 中，Excel 不负责这步）",
        "",
        "  layeredimage eileen:",
        "      always \"eileen_base\"              # 本体（始终显示）",
        "",
        "      group eyes:                         # 眼睛组",
        "          attribute normal default \"eileen_eyes_normal\"",
        "          attribute happy \"eileen_eyes_happy\"",
        "          attribute sad \"eileen_eyes_sad\"",
        "          attribute surprise \"eileen_eyes_surprise\"",
        "",
        "      group mouth:                        # 嘴巴组",
        "          attribute normal default \"eileen_mouth_normal\"",
        "          attribute smile \"eileen_mouth_smile\"",
        "          attribute frown \"eileen_mouth_frown\"",
        "          attribute open \"eileen_mouth_open\"",
        "",
        "      group eyebrows:                     # 眉毛组（可选）",
        "          attribute normal default \"eileen_brow_normal\"",
        "          attribute angry \"eileen_brow_angry\"",
        "",
        "  ★ 第二步：用 image（定义图片）注册各部件图片",
        "  可以在 Excel 中填写：",
        "    image（定义图片） → 角色名：eileen_base",
        "                       → 图片/背景：images/eileen/base.png",
        "    image（定义图片） → 角色名：eileen_eyes_happy",
        "                       → 图片/背景：images/eileen/eyes_happy.png",
        "    ... 依次注册所有部件",
        "",
        "  ★ 第三步：在 Excel 中用 show 切换表情",
        "  「位置/特效/属性」列直接填属性组合即可：",
        "",
        "    想显示的效果        → 指令       → 角色名  → 位置/特效/属性",
        "    ─────────────────────────────────────────────────────────",
        "    默认表情（普通）     → show      → eileen  → （留空）",
        "    高兴               → show      → eileen  → happy smile",
        "    悲伤+皱眉          → show      → eileen  → sad frown",
        "    惊讶+张嘴          → show      → eileen  → surprise open",
        "    高兴+左边          → show      → eileen  → happy smile at left",
        "    生气+右边+淡入     → show      → eileen  → angry frown at right with dissolve",
        "",
        "  也就是说：属性、位置、特效都写在同一列，空格分隔即可，",
        "  脚本会原样拼接到 show 语句后面。",
        "",
        "  ★ 如果不想用 LayeredImage（更简单的传统方式）：",
        "  每个表情一个独立图片文件，用 show 的「图片/背景」列：",
        "    show（显示角色） → 角色名：eileen",
        "                     → 图片/背景：eileen_happy（图片名）",
        "                     → 位置/特效/属性：at left",
        "",
        "  生成：show eileen eileen_happy at left",
    ]

    for row_idx, text in enumerate(help_lines, start=1):
        cell = ws_help.cell(row=row_idx, column=1, value=text)
        cell.font = Font(name="微软雅黑", size=10)
        if text.startswith("═══") or text.startswith("【"):
            cell.font = Font(name="微软雅黑", bold=True, size=11)
        elif text.startswith("  ") and not text.startswith("    "):
            pass  # 二级内容保持普通字号
    ws_help.column_dimensions["A"].width = 80

    _safe_save_workbook(wb, output_path, "模板")


def generate_blank_template(output_path: str):
    """生成纯空白模板（只有表头+下拉验证，无示例数据，无使用说明）"""
    wb = Workbook()
    ws = wb.active
    ws.title = "script"

    for col_idx, (header, width) in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 24

    from openpyxl.worksheet.datavalidation import DataValidation
    cmd_formula = '"' + ",".join(COMMAND_TYPES) + '"'
    dv = DataValidation(type="list", formula1=cmd_formula, allow_blank=True)
    dv.error = "请从下拉列表选择指令类型"
    dv.errorTitle = "无效指令"
    dv.prompt = "请选择指令类型（括号内是中文说明）"
    dv.promptTitle = "指令类型"
    dv.add("B2:B10000")
    ws.add_data_validation(dv)

    ws.freeze_panes = "A2"
    _safe_save_workbook(wb, output_path, "空白模板")


# ── 转换逻辑 ──────────────────────────────────────────────


def _get_desktop_documents() -> list:
    """获取桌面和文档目录候选（兼容中英文 Windows），返回 [(路径, 友好名称), ...]"""
    from pathlib import Path as _Path

    home = _Path.home()
    seen = set()
    result = []
    for name, label in [
        ("Desktop", "桌面"), ("desktop", "桌面"), ("桌面", "桌面"),
        ("Documents", "文档"), ("documents", "文档"), ("文档", "文档"),
        ("My Documents", "文档"), ("我的文档", "文档"),
    ]:
        p = home / name
        if p.is_dir() and p not in seen:
            seen.add(p)
            result.append((p, label))
    return result


def _try_copy_to_fallbacks(src_data: bytes, stem: str, ext: str, target_parent, label: str) -> str:
    """将 bytes 内容依次尝试写入目标目录或回退目录。成功返回路径，失败返回空字符串。"""
    from pathlib import Path as _Path

    fallback_dirs = [(target_parent, "当前目录")] + _get_desktop_documents()

    for dest_dir, dir_label in fallback_dirs:
        dest = dest_dir / f"{stem}{ext}"
        try:
            dest.write_bytes(src_data)
            if dest_dir != target_parent:
                print(f"\n⚠️  当前目录写入被拦截，已自动保存到{dir_label}：")
            else:
                print(f"✅ {label}已生成：")
            print(f"   {dest}")
            return str(dest)
        except (PermissionError, OSError):
            for i in range(1, 100):
                alt = dest_dir / f"{stem}_{i}{ext}"
                try:
                    alt.write_bytes(src_data)
                    if dest_dir != target_parent:
                        print(f"\n⚠️  当前目录写入被拦截，已自动保存到{dir_label}：")
                    else:
                        print(f"✅ {label}已生成：")
                    print(f"   {alt}")
                    return str(alt)
                except (PermissionError, OSError):
                    continue
            continue
    return ""


def _safe_save_workbook(wb, output_path: str, label: str):
    """
    安全保存 Excel 工作簿：用 BytesIO 绕过 openpyxl 的内部重命名，
    直接写字节到目标路径。被拦截时依次尝试桌面、文档。返回实际保存路径。
    """
    from io import BytesIO
    from pathlib import Path as _Path

    target = _Path(output_path).resolve()
    stem = target.stem
    ext = target.suffix

    # 用 BytesIO 获取字节，避免 openpyxl 自带的 temp+rename
    buffer = BytesIO()
    try:
        wb.save(buffer)
    except Exception:
        # BytesIO 保存失败，回退到 wb.save 直接写路径
        try:
            wb.save(str(target))
            print(f"✅ {label}已生成：")
            print(f"   {target}")
            return str(target)
        except (PermissionError, OSError):
            buffer = BytesIO()
            wb.save(buffer)

    data = buffer.getvalue()

    # 尝试直接写入目标目录
    last_error = None
    for attempt in range(3):
        try:
            target.write_bytes(data)
            print(f"✅ {label}已生成：")
            print(f"   {target}")
            return str(target)
        except (PermissionError, OSError) as e:
            last_error = e
            if attempt < 2:
                import time
                time.sleep(0.3)
            continue

    # 直接写入失败 → 尝试回退目录
    print(f"\n⚠️  无法写入目标目录：{target.parent}")
    print(f"   原因：{last_error}")
    result = _try_copy_to_fallbacks(data, stem, ext, target.parent, label)
    if result:
        return result

    # 全部失败，留在临时目录
    import tempfile
    tmp_dir = _Path(tempfile.gettempdir())
    tmp_path = tmp_dir / f"{stem}{ext}"
    tmp_path.write_bytes(data)
    print(f"\n⚠️  无法写入任何目录（Defender 拦截？），文件保留在临时目录：")
    print(f"   {tmp_path}")
    return str(tmp_path)


def _safe_write_text(output_path: str, content: str):
    """
    安全写入文本文件：优先直接写入目标路径，被拦截时依次尝试桌面、文档。
    返回实际保存的路径。
    """
    from pathlib import Path as _Path

    target = _Path(output_path).resolve()
    stem = target.stem
    ext = target.suffix
    data = content.encode("utf-8")

    # 尝试直接写入目标目录
    for attempt in range(3):
        try:
            target.write_bytes(data)
            return str(target)
        except (PermissionError, OSError):
            if attempt < 2:
                import time
                time.sleep(0.3)
            continue

    # 直接写入失败 → 尝试回退目录
    result = _try_copy_to_fallbacks(data, stem, ext, target.parent, "文件")
    if result:
        return result

    # 全部失败，留在临时目录
    import tempfile
    tmp_dir = _Path(tempfile.gettempdir())
    tmp_path = tmp_dir / f"{stem}{ext}"
    tmp_path.write_bytes(data)
    print(f"\n⚠️  无法写入任何目录（Defender 拦截？），文件保留在临时目录：")
    print(f"   {tmp_path}")
    return str(tmp_path)


# ── 转换逻辑（续）──────────────────────────────────────────


def _val(row, col_idx: int, default=""):
    """安全获取单元格值"""
    if col_idx >= len(row):
        return default
    v = row[col_idx]
    return str(v).strip() if v is not None else default


def _escape_rpy(text: str) -> str:
    """转义 Ren'Py 字符串中的特殊字符（仅转义双引号）"""
    return text.replace('"', '\\"')


def _parse_cmd(raw: str) -> str:
    """从指令单元格中提取英文指令名（兼容 'jump（跳转）' 格式）"""
    return raw.split("（")[0].strip().lower()


# ── 中文术语映射 ──────────────────────────────────────────

TRANSITION_MAP = {
    "溶解": "dissolve", "褪色": "fade", "闪白": 'Fade(0.1,0.0,0.5,color="#FFFFFF")',
    "像素化": "pixellate", "横向振动": "hpunch", "纵向振动": "vpunch",
    "百叶窗": "blinds", "网格覆盖": "squares", "擦除": "wipeleft",
    "滑入": "slideleft", "滑出": "slideawayleft", "推出": "pushright",
}

POSITION_MAP = {
    "左边": "left", "右边": "right", "中间": "center",
    "真中": "truecenter", "左外": "offscreenleft", "右外": "offscreenright",
}

AUDIO_CMD_MAP = {
    "停止": "stop", "播放": "play", "循环": "loop",
}


def _translate_effect(effect: str) -> str:
    """翻译中文特效/位置为 Ren'Py 英文"""
    if not effect:
        return ""
    parts = effect.split()
    translated = []
    for p in parts:
        if p in TRANSITION_MAP:
            translated.append(TRANSITION_MAP[p])
        elif p in POSITION_MAP:
            translated.append(POSITION_MAP[p])
        else:
            translated.append(p)
    return " ".join(translated)


def _trim_role_name(name: str) -> str:
    """修剪角色名首尾空白，收集警告"""
    trimmed = name.strip()
    if trimmed != name and name.strip():
        return trimmed
    return name


# ── 表格校验 ──────────────────────────────────────────────

def check_excel(input_path: str) -> list:
    """
    转换前校验 Excel，返回警告列表 [(row_num, severity, message), ...]。
    severity: "error" 可能导致语法错误, "warn" 建议检查
    """
    from openpyxl import load_workbook

    wb = load_workbook(input_path, data_only=True)
    issues = []

    # 收集所有 label 名
    all_labels = set()

    for sheet in wb.worksheets:
        rows = list(sheet.iter_rows(min_row=2, values_only=True))
        if not rows:
            continue

        # 第一遍：收集 label 名
        for row_idx, raw_row in enumerate(rows):
            row = [str(c).strip() if c is not None else "" for c in raw_row]
            if all(c == "" for c in row):
                continue
            cmd = _parse_cmd(_val(row, 1))
            scene_label = _val(row, 0)
            dialogue = _val(row, 4)
            if cmd == "label":
                name = scene_label if scene_label else dialogue
                if name and name in all_labels:
                    issues.append((row_idx + 2, "error", f"重复的标签名：{name}"))
                elif name:
                    all_labels.add(name)

        # 第二遍：逐行检查
        in_menu = False
        menu_has_options = False
        in_if = False
        if_has_body = False
        last_role = ""

        for row_idx, raw_row in enumerate(rows):
            excel_row = row_idx + 2
            row = [str(c).strip() if c is not None else "" for c in raw_row]
            if all(c == "" for c in row):
                if in_menu and not menu_has_options:
                    issues.append((menu_start_row, "error", "菜单没有任何选项"))
                in_menu = False
                menu_has_options = False
                if in_if and not if_has_body:
                    issues.append((if_start_row, "error", "if/elif/else 没有后续指令"))
                in_if = False
                if_has_body = False
                continue

            raw_cmd = _val(row, 1)
            cmd = _parse_cmd(raw_cmd)
            scene_label = _val(row, 0)
            image_path = _val(row, 2)
            character = _val(row, 3)
            dialogue = _val(row, 4)
            option_text = _val(row, 5)
            jump_target = _val(row, 6)
            audio = _val(row, 7)
            effect = _val(row, 8)

            # 前向填充角色名
            if character.strip():
                last_role = character.strip()
            role = last_role

            # 未知指令
            if raw_cmd.strip() and cmd not in ("label", "scene", "show", "hide",
                "dialogue", "narrator", "menu", "menu_option", "jump", "call",
                "return", "$", "if", "elif", "else", "play_music", "stop_music",
                "queue_music", "play_sound", "stop_sound", "voice", "pause",
                "player_input", "window", "define_character", "default", "image"):
                if raw_cmd.strip():
                    issues.append((excel_row, "warn", f"未知指令类型：{raw_cmd}"))

            # label 无名称
            if cmd == "label":
                name = scene_label if scene_label else dialogue
                if not name:
                    issues.append((excel_row, "error", "label 没有场景标签名"))

            # scene 无背景
            if cmd == "scene":
                if not image_path and not effect:
                    issues.append((excel_row, "warn", "scene 没有指定背景图"))

            # show 无图片
            if cmd == "show":
                if not image_path and not character:
                    issues.append((excel_row, "warn", "show 没有指定图片或角色"))

            # dialogue 无文本
            if cmd == "dialogue":
                if not role or role == "narrator":
                    issues.append((excel_row, "warn", "dialogue 没有角色名"))
                if not dialogue:
                    issues.append((excel_row, "warn", "dialogue 没有对话文本"))

            # narrator 无文本
            if cmd == "narrator":
                if not dialogue:
                    issues.append((excel_row, "warn", "narrator 没有文本"))

            # menu 检测
            if cmd == "menu":
                if in_menu:
                    if not menu_has_options:
                        issues.append((menu_start_row, "error", "菜单没有任何选项"))
                in_menu = True
                menu_has_options = False
                menu_start_row = excel_row

            if cmd == "menu_option":
                if not in_menu:
                    issues.append((excel_row, "error", "menu_option 不在 menu 内"))
                else:
                    menu_has_options = True
                if not option_text:
                    issues.append((excel_row, "warn", "menu_option 没有选项文本"))
                if not jump_target:
                    issues.append((excel_row, "warn", "menu_option 没有跳转目标"))

            # 非 menu_option 结束 menu
            if cmd not in ("menu", "menu_option") and in_menu:
                if not menu_has_options:
                    issues.append((menu_start_row, "error", "菜单没有任何选项"))
                in_menu = False
                menu_has_options = False

            # if/elif/else 检测
            if cmd == "if" or cmd == "elif":
                if in_if and not if_has_body:
                    issues.append((if_start_row, "error", "if 块没有后续指令"))
                in_if = True
                if_has_body = False
                if_start_row = excel_row
                if not dialogue:
                    issues.append((excel_row, "warn", "if/elif 没有条件表达式"))
            elif cmd == "else":
                if in_if and not if_has_body:
                    issues.append((if_start_row, "error", "if 块没有后续指令"))
                in_if = True
                if_has_body = False
                if_start_row = excel_row
            elif in_if:
                if_has_body = True

            # jump/call 目标检查
            if cmd in ("jump", "call"):
                if not jump_target:
                    issues.append((excel_row, "warn", f"{cmd} 没有跳转目标"))
                elif jump_target not in all_labels:
                    issues.append((excel_row, "warn", f"跳转目标 '{jump_target}' 未定义（可能在其他 Sheet）"))

            # 音频检查
            if cmd in ("play_music", "play_sound", "queue_music", "voice"):
                if not audio:
                    issues.append((excel_row, "warn", f"{cmd} 没有音频路径"))

            # player_input 检查
            if cmd == "player_input":
                if not character:
                    issues.append((excel_row, "warn", "player_input 没有变量名"))
                if not dialogue:
                    issues.append((excel_row, "warn", "player_input 没有提示文字"))

    return issues


def _print_check_result(issues: list, max_show: int = 20):
    """打印校验结果"""
    if not issues:
        print("  Check: No issues found.")
        return

    errors = [i for i in issues if i[1] == "error"]
    warns = [i for i in issues if i[1] == "warn"]
    print(f"  Check: {len(errors)} error(s), {len(warns)} warning(s)")

    shown = 0
    for row, sev, msg in errors + warns:
        if shown >= max_show:
            print(f"  ... and {len(issues) - max_show} more")
            break
        tag = "ERROR" if sev == "error" else "WARN"
        print(f"  [{tag}] 行 {row}: {msg}")
        shown += 1


def convert_excel_to_rpy(input_path: str, output_path: str):
    """将 Excel 转为 .rpy 脚本（支持多 Sheet）"""
    wb = load_workbook(input_path, data_only=True)

    warnings = []  # (sheet_name, row_num, message)

    def _warn(msg: str, row_num: int = 0):
        sheet_name = wb.active.title if wb.active else "?"
        warnings.append(f"[{sheet_name} 行{row_num + 1}] {msg}")

    all_sheet_lines = []  # [(label, lines, header_lines)]
    all_header_lines = []  # define/image at top level
    defined_characters = set()
    role_registry = set()  # 自动收集的角色名

    for sheet in wb.worksheets:
        rows = list(sheet.iter_rows(min_row=2, values_only=True))
        if not rows:
            continue

        lines = []
        header_lines = []
        in_menu = False
        in_if_body = False
        base_indent = "    "

        # 状态追踪
        current_role = ""       # 前向填充的角色名
        current_characters = []  # 当前在场立绘缓存: [(name, alias)]
        scene_sheet_label = ""   # 首个 label 名

        def _indent():
            return base_indent + ("    " if in_if_body else "")

        def _end_if_body():
            nonlocal in_if_body
            in_if_body = False

        for row_idx, raw_row in enumerate(rows):
            row = [str(c).strip() if c is not None else "" for c in raw_row]

            # 跳过完全空行
            if all(c == "" for c in row):
                if in_menu:
                    in_menu = False
                _end_if_body()
                continue

            scene_label = _val(row, 0)
            raw_cmd = _val(row, 1)
            cmd = _parse_cmd(raw_cmd)
            image_path = _val(row, 2)
            character = _val(row, 3)
            dialogue = _val(row, 4)
            option_text = _val(row, 5)
            jump_target = _val(row, 6)
            audio = _val(row, 7)
            effect = _translate_effect(_val(row, 8))
            notes = _val(row, 9)

            # ── 角色名前向填充 & 修剪 ──
            if character.strip():
                trimmed = _trim_role_name(character)
                if trimmed != character and character.strip():
                    _warn(f"角色名首尾含空白，已修剪：'{character}' → '{trimmed}'", row_idx)
                current_role = trimmed
            role = current_role  # 当前行生效的角色名

            # ── 旁白自动检测 ──
            if role == "旁白":
                role = "narrator"

            # ── 空指令 + 有角色名 → 自动推断为 dialogue ──
            if not raw_cmd.strip() and role and role != "narrator":
                cmd = "dialogue"
            # ── 空指令 + 有对话文本 + 无角色 → 旁白 ──
            if not raw_cmd.strip() and dialogue and not role:
                cmd = "narrator"

            # ── 立绘缓存自动回收（可选） ──
            # 取消注释以下代码启用 auto-hide：
            # if cmd == "show":
            #     show_name = character if character else image_path
            #     for cached_name, cached_alias in current_characters:
            #         if cached_name != show_name:
            #             lines.append(f"{_indent()}hide {cached_name}" + (f" as {cached_alias}" if cached_alias else ""))
            #     current_characters = [(show_name, "")]
            #     if image_path and effect and " as " in effect:
            #         parts = effect.split(" as ", 1)
            #         alias = parts[1].split()[0] if len(parts) > 1 else ""
            #         if alias:
            #             current_characters = [(show_name, alias)]
            # elif cmd == "hide":
            #     hide_name = character if character else image_path
            #     current_characters = [(n, a) for n, a in current_characters if n != hide_name]
            # elif cmd == "scene":
            #     current_characters = []

            # ── define_character / default / image（提前收集到文件头）──
            if cmd == "define_character":
                name = character if character else "unknown"
                defined_characters.add(name)
                if dialogue:
                    line = f"define {name} = {dialogue}"
                else:
                    line = f'define {name} = Character("{name}")'
                header_lines.append(line)
                continue

            if cmd == "default":
                name = character if character else "unknown"
                val = dialogue if dialogue else '""'
                header_lines.append(f"default {name} = {val}")
                continue

            if cmd == "image":
                img_name = character if character else "unknown"
                img_path = image_path if image_path else ""
                if "(" in img_path:
                    header_lines.append(f'image {img_name} = {img_path}')
                elif img_path:
                    header_lines.append(f'image {img_name} = "{_escape_rpy(img_path)}"')
                else:
                    header_lines.append(f'image {img_name}')
                continue

            # ── label ──
            if cmd == "label":
                _end_if_body()
                in_menu = False
                label_name = scene_label if scene_label else dialogue
                if not label_name:
                    continue
                lines.append(f"label {label_name}:")
                continue

            # ── scene ──
            if cmd == "scene":
                img = _escape_rpy(image_path) if image_path else ""
                fx = f" {effect}" if effect else ""
                if img:
                    lines.append(f"{_indent()}scene {img}{fx}")
                elif effect:
                    lines.append(f"{_indent()}scene{fx}")
                continue

            # ── show ──
            if cmd == "show":
                char = character if character else ""
                img = _escape_rpy(image_path) if image_path else ""
                fx = f" {effect}" if effect else ""
                if char and img:
                    lines.append(f"{_indent()}show {char} {img}{fx}")
                elif char:
                    lines.append(f"{_indent()}show {char}{fx}")
                elif img:
                    lines.append(f"{_indent()}show {img}{fx}")
                continue

            # ── hide ──
            if cmd == "hide":
                char = character if character else ""
                fx = f" {effect}" if effect else ""
                if char:
                    lines.append(f"{_indent()}hide {char}{fx}")
                continue

            # ── dialogue ──
            if cmd == "dialogue":
                char = character if character else ""
                text = _escape_rpy(dialogue) if dialogue else ""
                if char and text:
                    lines.append(f'{_indent()}{char} "{text}"')
                continue

            # ── narrator ──
            if cmd == "narrator":
                text = _escape_rpy(dialogue) if dialogue else ""
                if text:
                    if notes and "centered" in notes.lower():
                        lines.append(f'{_indent()}centered "{text}"')
                    elif notes and "explicit" in notes.lower():
                        lines.append(f'{_indent()}narrator "{text}"')
                    else:
                        lines.append(f'{_indent()}"{text}"')
                continue

            # ── menu ──
            if cmd == "menu":
                _end_if_body()
                lines.append(f"{_indent()}menu:")
                in_menu = True
                continue

            # ── menu_option ──
            if cmd == "menu_option":
                opt = _escape_rpy(option_text) if option_text else ""
                target = jump_target if jump_target else ""
                if opt and target:
                    lines.append(f'{_indent()}    "{opt}":')
                    if target.startswith("jump "):
                        lines.append(f"{_indent()}        jump {target[5:]}")
                    elif target in ("return",) or target.startswith("call ") or target.startswith("$ "):
                        lines.append(f"{_indent()}        {target}")
                    else:
                        lines.append(f"{_indent()}        jump {target}")
                elif opt:
                    lines.append(f'{_indent()}    "{opt}":')
                    lines.append(f"{_indent()}        pass")
                continue

            # 非 menu_option 且在 menu 中 → 结束 menu
            if cmd != "menu_option" and in_menu:
                in_menu = False

            # ── jump ──
            if cmd == "jump":
                target = jump_target if jump_target else ""
                if target:
                    lines.append(f"{_indent()}jump {target}")
                continue

            # ── call ──
            if cmd == "call":
                target = jump_target if jump_target else ""
                if target:
                    lines.append(f"{_indent()}call {target}")
                continue

            # ── return ──
            if cmd == "return":
                lines.append(f"{_indent()}return")
                continue

            # ── $（设置变量）──
            if cmd == "$":
                code = dialogue if dialogue else ""
                if code:
                    lines.append(f"{_indent()}$ {code}")
                continue

            # ── if ──
            if cmd == "if":
                _end_if_body()
                cond = dialogue if dialogue else "True"
                lines.append(f"{_indent()}if {cond}:")
                in_if_body = True
                continue

            # ── elif ──
            if cmd == "elif":
                _end_if_body()
                cond = dialogue if dialogue else "True"
                lines.append(f"{_indent()}elif {cond}:")
                in_if_body = True
                continue

            # ── else ──
            if cmd == "else":
                _end_if_body()
                lines.append(f"{_indent()}else:")
                in_if_body = True
                continue

            # ── play_music ──
            if cmd == "play_music":
                aud = _escape_rpy(audio) if audio else ""
                if aud:
                    lines.append(f'{_indent()}play music "{aud}"')
                continue

            # ── stop_music ──
            if cmd == "stop_music":
                lines.append(f"{_indent()}stop music {effect}" if effect else f"{_indent()}stop music")
                continue

            # ── play_sound ──
            if cmd == "play_sound":
                aud = _escape_rpy(audio) if audio else ""
                if aud:
                    lines.append(f'{_indent()}play sound "{aud}"')
                continue

            # ── stop_sound ──
            if cmd == "stop_sound":
                lines.append(f"{_indent()}stop sound")
                continue

            # ── queue_music ──
            if cmd == "queue_music":
                aud = _escape_rpy(audio) if audio else ""
                if aud:
                    lines.append(f'{_indent()}queue music "{aud}"')
                continue

            # ── voice ──
            if cmd == "voice":
                aud = _escape_rpy(audio) if audio else ""
                if aud:
                    lines.append(f'{_indent()}voice "{aud}"')
                continue

            # ── window ──
            if cmd == "window":
                action = dialogue.lower() if dialogue else "show"
                if action in ("show", "hide", "auto"):
                    lines.append(f"{_indent()}window {action}")
                else:
                    lines.append(f"{_indent()}window show")
                continue

            # ── pause ──
            if cmd == "pause":
                dur = dialogue if dialogue else "1.0"
                try:
                    float(dur)
                except ValueError:
                    dur = "1.0"
                lines.append(f"{_indent()}pause {dur}")
                continue

            # ── player_input（玩家输入）──
            if cmd == "player_input":
                var_name = character if character else "input_result"
                prompt = dialogue if dialogue else "请输入："
                default_val = effect if effect else ""
                prompt_escaped = _escape_rpy(prompt)
                if default_val:
                    lines.append(f'{_indent()}$ {var_name} = renpy.input("{prompt_escaped}").strip() or "{_escape_rpy(default_val)}"')
                else:
                    lines.append(f'{_indent()}$ {var_name} = renpy.input("{prompt_escaped}").strip()')
                continue

            # 未知指令 → 跳过

            # ── 记录使用中的角色名 ──
            if role and role != "narrator":
                role_registry.add(role)
            # show 指令中的图片名也注册
            if cmd == "show" and image_path:
                img_name = image_path.split()[0] if image_path else ""
                if img_name and not img_name.startswith("bg "):
                    role_registry.add(img_name)

        # ── 每 Sheet 处理完毕 ──
        sheet_label = scene_sheet_label or sheet.title
        all_sheet_lines.append((sheet_label, lines, header_lines))

    # ── 自动生成未定义的 Character ──
    auto_defines = role_registry - defined_characters - {"narrator", ""}
    for name in sorted(auto_defines):
        all_header_lines.insert(0, f'define {name} = Character("{name}")')

    # ── 组装输出 ──
    # 首个 sheet 的 header_lines 作为全局 header
    first_sheet_headers = all_sheet_lines[0][2] if all_sheet_lines else []
    output_parts = ["# Generated by excel_to_rpy.py",
                    "# 源文件: " + os.path.basename(input_path),
                    ""]
    output_parts.extend(all_header_lines)
    output_parts.extend(first_sheet_headers)
    output_parts.append("")

    total_lines = 0
    first = True
    for label, lines, _headers in all_sheet_lines:
        if first:
            first = False
        else:
            output_parts.append("")
            # 只在 Sheet 没有自带 label 时才加
            has_label = any(l.strip().startswith("label ") for l in lines)
            if not has_label:
                output_parts.append(f"label {label}:")
        output_parts.extend(lines)
        total_lines += len(lines)

    output = "\n".join(output_parts) + "\n"
    saved_path = _safe_write_text(output_path, output)

    print(f"Complete: {saved_path}")
    print(f"   {total_lines} lines / {len(all_sheet_lines)} sheet(s)")

    # ── 警告汇总 ──
    if warnings:
        print(f"\n[WARN] {len(warnings)} issue(s):")
        for w in warnings:
            print(f"  {w}")

    # ── 自动注册的角色 ──
    if auto_defines:
        print(f"\n[INFO] Auto-defined {len(auto_defines)} character(s):")
        for n in sorted(auto_defines):
            print(f"  define {n} = Character(\"{n}\")")


# ── 入口 ───────────────────────────────────────────────────

def _strip_quotes(s: str) -> str:
    s = s.strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ('"', "'"):
        s = s[1:-1]
    return s


def _pause():
    input("\n按 Enter 键继续...")


def _interactive_convert():
    print("\n" + "-" * 40)
    path = input("请输入 Excel 文件路径（可直接拖拽文件到此处）：\n> ").strip()
    path = _strip_quotes(path)
    if not path:
        print("[ERROR] 未输入路径，返回主菜单。")
        return
    if not os.path.isfile(path):
        print(f"[ERROR] 文件不存在：{path}")
        return
    if not path.lower().endswith(".xlsx"):
        print(f"[ERROR] 不是 .xlsx 文件，请拖拽 Excel 表格。")
        print(f"       .rpy 文件请用 rpy_to_excel.py 打开。")
        return
    default_out = str(Path(path).with_suffix(".rpy"))
    out = input(f"输出路径（直接回车使用默认：{default_out}）：\n> ").strip()
    out = _strip_quotes(out) if out else default_out
    print("\n校验表格...")
    issues = check_excel(path)
    _print_check_result(issues)
    if any(i[1] == "error" for i in issues):
        go = input("\n存在错误，是否继续转换？(y/N)：").strip().lower()
        if go != "y":
            return
    convert_excel_to_rpy(path, out)


def _interactive_template():
    print("\n" + "-" * 40)
    default_path = "renpy_script_template.xlsx"
    path = input(f"模板保存路径（直接回车使用默认：{default_path}）：\n> ").strip()
    path = _strip_quotes(path) if path else default_path
    generate_template(path)


def _interactive_check():
    print("\n" + "-" * 40)
    path = input("请输入 Excel 文件路径（可直接拖拽文件到此处）：\n> ").strip()
    path = _strip_quotes(path)
    if not path:
        print("[ERROR] 未输入路径，返回主菜单。")
        return
    if not os.path.isfile(path):
        print(f"[ERROR] 文件不存在：{path}")
        return
    if not path.lower().endswith(".xlsx"):
        print(f"[ERROR] 不是 .xlsx 文件。")
        return
    issues = check_excel(path)
    _print_check_result(issues, max_show=50)


def _interactive_blank():
    """交互式生成空白模板"""
    print("\n" + "-" * 40)
    default_path = "renpy_script_blank.xlsx"
    path = input(f"空白模板保存路径（直接回车使用默认：{default_path}）：\n> ").strip()
    path = _strip_quotes(path) if path else default_path
    generate_blank_template(path)


def interactive_mode():
    while True:
        print()
        print("=" * 40)
        print("  Ren'Py Excel → .rpy 转换工具")
        print("=" * 40)
        print("  1. 转换 Excel 为 .rpy")
        print("  2. 校验 Excel 表格（不转换）")
        print("  3. 生成 Excel 模板（含示例和教学）")
        print("  4. 生成空白 Excel 模板（仅表头）")
        print("  5. 退出")
        print("-" * 40)
        choice = input("请选择 (1/2/3/4/5)：").strip()
        if choice == "1":
            _interactive_convert()
            _pause()
        elif choice == "2":
            _interactive_check()
            _pause()
        elif choice == "3":
            _interactive_template()
            _pause()
        elif choice == "4":
            _interactive_blank()
            _pause()
        elif choice == "5":
            print("再见！")
            break
        else:
            print("[ERROR] 无效选择，请输入 1、2、3、4 或 5。")


def _fix_working_directory():
    """如果当前工作目录是系统目录（如 C:\\Windows\\System32），切换到脚本所在目录。"""
    cwd = os.getcwd()
    windir = os.environ.get("SystemRoot", r"C:\Windows")
    windir = os.path.normpath(windir).lower()
    if os.path.normpath(cwd).lower().startswith(windir + os.sep) or os.path.normpath(cwd).lower() == windir:
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        os.chdir(script_dir)
        print(f"⚠️  已从系统目录（{cwd}）切换到脚本目录：{script_dir}\n")


def main():
    _fix_working_directory()
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description="Ren'Py Excel → .rpy 转换工具")
        parser.add_argument("input", nargs="?", help="输入的 .xlsx 文件路径")
        parser.add_argument("-o", "--output", help="输出的 .rpy 文件路径")
        parser.add_argument("--template", action="store_true", help="生成 Excel 模板")
        parser.add_argument("--blank", action="store_true", help="生成空白 Excel 模板（无示例数据）")
        parser.add_argument("--template-path", default="renpy_script_template.xlsx", help="模板保存路径")
        parser.add_argument("--check", action="store_true", help="仅校验表格，不转换")
        args = parser.parse_args()

        if args.template:
            generate_template(args.template_path)
            return

        if args.blank:
            generate_blank_template(args.template_path)
            return

        if not args.input:
            parser.print_help()
            print("\n提示：使用 --template 生成模板，或拖入 .xlsx 文件转换。")
            input("\n按 Enter 键退出...")
            return

        input_path = _strip_quotes(args.input)
        if not os.path.isfile(input_path):
            print(f"❌ 文件不存在：{input_path}")
            input("\n按 Enter 键退出...")
            sys.exit(1)

        output_path = args.output if args.output else str(Path(input_path).with_suffix(".rpy"))

        print("校验表格...")
        issues = check_excel(input_path)
        _print_check_result(issues)

        if args.check:
            return
        if any(i[1] == "error" for i in issues):
            print("\n存在错误，使用 --check 仅校验不转换，或添加 -o 强制输出。")
            print("继续转换中...")
        convert_excel_to_rpy(input_path, output_path)
        try:
            input("\n按 Enter 键退出...")
        except (EOFError, OSError):
            pass
        return

    # 交互模式
    try:
        interactive_mode()
    except KeyboardInterrupt:
        print("\n\n再见！")
    except Exception as e:
        print(f"\n❌ 发生错误：{e}")
        input("\n按 Enter 键退出...")


if __name__ == "__main__":
    main()
