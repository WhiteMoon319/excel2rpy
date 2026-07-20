#!/usr/bin/env python3
"""
将 .rpy 文件反向转为 excel_to_rpy.py 对应的 Excel 格式。
用法: python rpy_to_excel.py script.rpy [-o output.xlsx]
"""

import re
import sys
import os
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    print("请先安装 openpyxl：pip install openpyxl")
    sys.exit(1)

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
    "image（定义图片）",
]

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


def parse_rpy(filepath: str) -> list:
    """解析 .rpy 文件，返回行列表，每行是一个 dict"""
    with open(filepath, "r", encoding="utf-8-sig") as f:
        raw_lines = f.readlines()

    result = []
    # 先跳过空白行处理，保留有意义的行
    lines = []
    for line in raw_lines:
        stripped = line.rstrip("\n\r")
        if stripped.strip() == "":
            lines.append(("", 0, ""))
        else:
            indent = len(stripped) - len(stripped.lstrip())
            content = stripped.strip()
            lines.append((stripped, indent, content))

    in_menu = False
    in_python = False
    python_indent = 0
    in_if_body = False  # 是否在 if/elif/else 条件体中
    if_base_indent = -1

    for i, (raw, indent, content) in enumerate(lines):
        if content == "":
            if in_if_body:
                result.append({"_type": "blank"})  # 仅条件体结束需保留空行
            in_menu = False
            in_if_body = False
            continue

        if content.startswith("#"):
            continue

        # 条件体结束检测：缩进回到 if 层级或更低
        if in_if_body and indent <= if_base_indent:
            result.append({"_type": "blank"})
            in_if_body = False

        # 跳过注释
        if content.startswith("#"):
            continue

        # python: 块处理
        if in_python:
            if indent <= python_indent and not content.startswith(" "):
                in_python = False
                result.append({"_type": "blank"})  # 结束条件体
                # 不 continue，让当前行走正常解析
            else:
                # 识别 python 块内的 if/elif/else，用对应的指令类型
                if content == "else:":
                    result.append({"_type": "else", "指令类型": "else（否则）"})
                elif re.match(r"elif\s+", content) and content.rstrip().endswith(":"):
                    result.append({"_type": "elif", "指令类型": "elif（否则如果）",
                                   "对话文本": content[5:].rstrip(":").strip()})
                elif re.match(r"if\s+", content) and content.rstrip().endswith(":"):
                    result.append({"_type": "if", "指令类型": "if（条件判断）",
                                   "对话文本": content[3:].rstrip(":").strip()})
                else:
                    result.append({"_type": "$", "指令类型": "$（设置变量）",
                                   "对话文本": content})
                continue

        # Top-level define / default / image
        if indent == 0:
            if content.startswith("define "):
                m = re.match(r"define\s+(\w+)\s*=\s*(.+)", content)
                if m:
                    result.append({
                        "_type": "define_character",
                        "指令类型": "define_character（定义角色）",
                        "角色名": m.group(1),
                        "对话文本": m.group(2),
                    })
                continue

            if content.startswith("default "):
                m = re.match(r"default\s+(\w+)\s*=\s*(.+)", content)
                if m:
                    val = m.group(2).strip()
                    # Keep outer quotes for default values (e.g., default x = "")
                    result.append({
                        "_type": "default",
                        "指令类型": "default（默认变量）",
                        "角色名": m.group(1),
                        "对话文本": val,
                    })
                continue

            if content.startswith("image "):
                m = re.match(r"image\s+(.+?)\s*=\s*(.+)", content)
                if m:
                    img_name = m.group(1)
                    img_val = m.group(2).strip()
                    # Strip outer quotes if present (simple string paths)
                    if (img_val.startswith('"') and img_val.endswith('"')) or \
                       (img_val.startswith("'") and img_val.endswith("'")):
                        img_val = img_val[1:-1]
                    result.append({
                        "_type": "image",
                        "指令类型": "image（定义图片）",
                        "角色名": img_name,
                        "图片/背景": img_val,
                    })
                continue

            if content.startswith("label "):
                m = re.match(r"label\s+(\w+)\s*:", content)
                if m:
                    result.append({
                        "_type": "label",
                        "指令类型": "label（场景标签）",
                        "场景标签": m.group(1),
                    })
                in_menu = False
                continue

        # Indented content
        if indent >= 4:
            # menu:
            if re.match(r"menu\s*:", content):
                result.append({
                    "_type": "menu",
                    "指令类型": "menu（选择菜单）",
                })
                in_menu = True
                continue

            # menu option: "text":
            if in_menu and re.match(r'"([^"]*)"\s*:', content):
                m = re.match(r'"([^"]*)"\s*:', content)
                result.append({
                    "_type": "menu_option",
                    "指令类型": "menu_option（菜单选项）",
                    "选项文本": m.group(1),
                    "跳转目标": "",  # filled by next line
                })
                continue

            # menu action: jump / return / call / $ 等
            if in_menu and result and result[-1]["_type"] == "menu_option":
                if content.startswith("jump "):
                    result[-1]["跳转目标"] = "jump " + content[5:].strip()
                    continue
                elif content == "return":
                    result[-1]["跳转目标"] = "return"
                    continue
                elif content.startswith("call "):
                    result[-1]["跳转目标"] = "call " + content[5:].strip()
                    continue
                elif content.startswith("$ "):
                    result[-1]["跳转目标"] = "$ " + content[2:].strip()
                    continue
                continue

            # if / elif / else
            if re.match(r"if\s+", content) and content.rstrip().endswith(":"):
                cond = content[3:].rstrip(":").strip()
                result.append({
                    "_type": "if",
                    "指令类型": "if（条件判断）",
                    "对话文本": cond,
                })
                in_if_body = True
                if_base_indent = indent
                continue

            if re.match(r"elif\s+", content) and content.rstrip().endswith(":"):
                cond = content[5:].rstrip(":").strip()
                result.append({
                    "_type": "elif",
                    "指令类型": "elif（否则如果）",
                    "对话文本": cond,
                })
                in_if_body = True
                if_base_indent = indent
                continue

            if content == "else:":
                result.append({
                    "_type": "else",
                    "指令类型": "else（否则）",
                })
                in_if_body = True
                if_base_indent = indent
                continue

            # python: 块
            if content == "python:":
                in_python = True
                python_indent = indent
                continue

            # scene
            if content.startswith("scene "):
                rest = content[6:].strip()
                img = ""
                effect = ""
                # scene <img> with <effect>
                m = re.match(r"(.+?)\s+with\s+(.+)", rest)
                if m:
                    img = m.group(1).strip()
                    effect = "with " + m.group(2).strip()
                else:
                    img = rest
                result.append({
                    "_type": "scene",
                    "指令类型": "scene（背景图）",
                    "图片/背景": img,
                    "位置/特效/属性": effect,
                })
                continue

            # show
            if content.startswith("show "):
                rest = content[5:].strip()
                img = ""
                alias = ""
                effect = ""

                # show <img> as <alias> at <pos> with <effect>
                parts = rest.split()
                i = 0
                img_parts = []
                while i < len(parts) and parts[i] not in ("as", "at", "with"):
                    img_parts.append(parts[i])
                    i += 1
                img = " ".join(img_parts)

                remaining = parts[i:]
                effect = " ".join(remaining) if remaining else ""

                result.append({
                    "_type": "show",
                    "指令类型": "show（显示角色）",
                    "图片/背景": img,
                    "位置/特效/属性": effect,
                })
                continue

            # hide
            if content.startswith("hide "):
                rest = content[5:].strip()
                effect = ""
                m = re.match(r"(\S+)\s+(.+)", rest)
                if m:
                    rest = m.group(1)
                    effect = m.group(2)
                result.append({
                    "_type": "hide",
                    "指令类型": "hide（隐藏角色）",
                    "角色名": rest,
                    "位置/特效/属性": effect,
                })
                continue

            # stop music
            if content.startswith("stop music"):
                rest = content[10:].strip()
                effect = rest if rest else ""
                result.append({
                    "_type": "stop_music",
                    "指令类型": "stop_music（停止BGM）",
                    "位置/特效/属性": effect,
                })
                continue

            # stop sound
            if content == "stop sound":
                result.append({
                    "_type": "stop_sound",
                    "指令类型": "stop_sound（停止音效）",
                })
                continue

            # play music
            if content.startswith("play music "):
                m = re.match(r'play music\s+"([^"]*)"', content)
                audio = m.group(1) if m else ""
                result.append({
                    "_type": "play_music",
                    "指令类型": "play_music（播放BGM）",
                    "音频路径": audio,
                })
                continue

            # play sound
            if content.startswith("play sound "):
                m = re.match(r'play sound\s+"([^"]*)"', content)
                audio = m.group(1) if m else ""
                result.append({
                    "_type": "play_sound",
                    "指令类型": "play_sound（播放音效）",
                    "音频路径": audio,
                })
                continue

            # queue music
            if content.startswith("queue music "):
                m = re.match(r'queue music\s+"([^"]*)"', content)
                audio = m.group(1) if m else ""
                result.append({
                    "_type": "queue_music",
                    "指令类型": "queue_music（排队播BGM）",
                    "音频路径": audio,
                })
                continue

            # voice
            if content.startswith("voice "):
                m = re.match(r'voice\s+"([^"]*)"', content)
                audio = m.group(1) if m else ""
                result.append({
                    "_type": "voice",
                    "指令类型": "voice（配音）",
                    "音频路径": audio,
                })
                continue

            # window
            if content.startswith("window "):
                action = content[7:].strip()
                result.append({
                    "_type": "window",
                    "指令类型": "window（对话框开关）",
                    "对话文本": action,
                })
                continue

            # pause
            if content.startswith("pause "):
                dur = content[6:].strip()
                result.append({
                    "_type": "pause",
                    "指令类型": "pause（暂停等待）",
                    "对话文本": dur,
                })
                continue

            # return
            if content == "return":
                result.append({
                    "_type": "return",
                    "指令类型": "return（返回）",
                })
                continue

            # jump
            if content.startswith("jump "):
                target = content[5:].strip()
                result.append({
                    "_type": "jump",
                    "指令类型": "jump（跳转）",
                    "跳转目标": target,
                })
                continue

            # call
            if content.startswith("call "):
                target = content[5:].strip()
                result.append({
                    "_type": "call",
                    "指令类型": "call（调用子场景）",
                    "跳转目标": target,
                })
                continue

            # $ (python one-liner)
            if content.startswith("$ "):
                code = content[2:].strip()
                result.append({
                    "_type": "$",
                    "指令类型": "$（设置变量）",
                    "对话文本": code,
                })
                continue

            # centered (特殊旁白)
            if content.startswith("centered "):
                m = re.match(r'centered\s+"([^"]*)"', content)
                text = m.group(1) if m else content[9:].strip().strip('"')
                result.append({
                    "_type": "narrator",
                    "指令类型": "narrator（旁白）",
                    "对话文本": text,
                    "备注": "centered",
                })
                continue

            # narrator "text"
            if content.startswith("narrator "):
                m = re.match(r'narrator\s+"([^"]*)"', content)
                text = m.group(1) if m else content[9:].strip().strip('"')
                result.append({
                    "_type": "narrator",
                    "指令类型": "narrator（旁白）",
                    "对话文本": text,
                    "备注": "explicit",
                })
                continue

            # character "text" (dialogue)
            m_dialogue = re.match(r'(\w+)\s+"([^"]*)"', content)
            if m_dialogue:
                char = m_dialogue.group(1)
                text = m_dialogue.group(2)
                if char in ("narrator", "centered", ""):
                    result.append({
                        "_type": "narrator",
                        "指令类型": "narrator（旁白）",
                        "对话文本": text,
                    })
                else:
                    result.append({
                        "_type": "dialogue",
                        "指令类型": "dialogue（角色对话）",
                        "角色名": char,
                        "对话文本": text,
                    })
                continue

            # bare string "text" (narrator) — 可能同一行有多个字符串
            m_bare = re.match(r'"([^"]*)"', content)
            if m_bare:
                result.append({
                    "_type": "narrator",
                    "指令类型": "narrator（旁白）",
                    "对话文本": m_bare.group(1),
                })
                # 检查同一行是否还有第二个字符串
                rest = content[m_bare.end():].strip()
                m2 = re.match(r'"([^"]*)"', rest)
                if m2:
                    result.append({
                        "_type": "narrator",
                        "指令类型": "narrator（旁白）",
                        "对话文本": m2.group(1),
                    })
                continue

    return result


def write_excel(rows: list, output_path: str):
    """将解析结果写入 Excel"""
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
    cmd_formula = '"' + ",".join(COMMAND_TYPES) + '"'
    dv = DataValidation(type="list", formula1=cmd_formula, allow_blank=True)
    dv.error = "请从下拉列表选择指令类型"
    dv.errorTitle = "无效指令"
    dv.add("B2:B10000")
    ws.add_data_validation(dv)

    # 写数据行
    row_num = 1
    for item in rows:
        row_num += 1
        if item.get("_type") == "blank":
            # 写空行 — excel_to_rpy 用空行来结束 if/menu 块
            for col_idx in range(1, len(HEADERS) + 1):
                cell = ws.cell(row=row_num, column=col_idx, value="")
                cell.border = THIN_BORDER
            ws.row_dimensions[row_num].height = 8
            continue

        values = [
            item.get("场景标签", ""),
            item.get("指令类型", ""),
            item.get("图片/背景", ""),
            item.get("角色名", ""),
            item.get("对话文本", ""),
            item.get("选项文本", ""),
            item.get("跳转目标", ""),
            item.get("音频路径", ""),
            item.get("位置/特效/属性", ""),
            item.get("备注", ""),
        ]

        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=value)
            cell.font = CELL_FONT
            cell.alignment = CELL_ALIGN
            cell.border = THIN_BORDER
            if row_num % 2 == 0:
                cell.fill = ALT_FILL
        ws.row_dimensions[row_num].height = 20

    ws.freeze_panes = "A2"
    wb.save(output_path)
    print(f"Excel saved: {output_path}")
    print(f"   {row_num - 1} rows total")


def _strip_quotes(s: str) -> str:
    s = s.strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ('"', "'"):
        s = s[1:-1]
    return s


def _interactive_convert():
    print("\n" + "-" * 40)
    path = input("请输入 .rpy 文件路径（可直接拖拽文件到此处）：\n> ").strip()
    path = _strip_quotes(path)
    if not path:
        print("[ERROR] 未输入路径，返回主菜单。")
        return
    if not os.path.isfile(path):
        print(f"[ERROR] 文件不存在：{path}")
        return
    if not path.lower().endswith(".rpy"):
        print(f"[ERROR] 不是 .rpy 文件，请拖拽 .rpy 脚本。")
        print(f"       .xlsx 文件请用 excel_to_rpy.py 打开。")
        return
    default_out = str(Path(path).with_suffix(".xlsx"))
    out = input(f"输出路径（直接回车使用默认：{default_out}）：\n> ").strip()
    out = _strip_quotes(out) if out else default_out
    try:
        rows = parse_rpy(path)
    except UnicodeDecodeError:
        print(f"[ERROR] 文件编码错误，请确认是 .rpy 文本文件而非 .xlsx 二进制文件。")
        return
    write_excel(rows, out)


def interactive_mode():
    while True:
        print()
        print("=" * 40)
        print("  Ren'Py .rpy -> Excel 反向转换工具")
        print("=" * 40)
        print("  1. 转换 .rpy 为 Excel")
        print("  2. 退出")
        print("-" * 40)
        choice = input("请选择 (1/2)：").strip()
        if choice == "1":
            _interactive_convert()
            input("\n按 Enter 键继续...")
        elif choice == "2":
            print("再见！")
            break
        else:
            print("[ERROR] 无效选择，请输入 1 或 2。")


def _fix_working_directory():
    """如果当前工作目录是系统目录（如 C:\\Windows\\System32），切换到脚本所在目录。"""
    cwd = os.getcwd()
    windir = os.environ.get("SystemRoot", r"C:\Windows")
    windir = os.path.normpath(windir).lower()
    if os.path.normpath(cwd).lower().startswith(windir + os.sep) or os.path.normpath(cwd).lower() == windir:
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        os.chdir(script_dir)
        print(f"[WARN] Switched from system dir ({cwd}) to script dir: {script_dir}\n")


def main():
    _fix_working_directory()

    if len(sys.argv) < 2:
        try:
            interactive_mode()
        except KeyboardInterrupt:
            print("\n\n再见！")
        except Exception as e:
            print(f"\n[ERROR] {e}")
            input("\n按 Enter 键退出...")
        return

    input_path = sys.argv[1]
    output_path = None
    if "-o" in sys.argv:
        idx = sys.argv.index("-o")
        if idx + 1 < len(sys.argv):
            output_path = sys.argv[idx + 1]

    if not output_path:
        output_path = str(Path(input_path).with_suffix(".xlsx"))

    if not os.path.isfile(input_path):
        print(f"[ERROR] File not found: {input_path}")
        sys.exit(1)

    rows = parse_rpy(input_path)
    write_excel(rows, output_path)


if __name__ == "__main__":
    main()
