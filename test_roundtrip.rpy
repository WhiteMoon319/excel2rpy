define narrator = Character("")
define tester = Character("测试员")
define 同学A = Character("同学A")
default score = 0
default has_key = False
default money = 100
default health = 50
default love = 0

image bg room = "images/bg_room.jpg"
image bg park = "images/bg_park.jpg"
image tester happy = "images/tester_happy.png"

image test_custom = Solid("#ff0000", xsize=100, ysize=100)

label start:

    $ score = 0

    scene bg room with dissolve

    show tester happy at center

    "欢迎来到测试场景。"

    tester "你好！我是测试员。"

    "这个测试覆盖了所有指令类型。"

    $ score += 10

    if score >= 10:
        "你的分数达到了 10 分！"
        jump passed
    else:
        "分数不足，再试一次。"
        return

label passed:

    scene bg park with fade

    show tester happy at left with dissolve

    tester "恭喜通过！"

    play music "audio/bgm_win.ogg"

    window hide
    centered "{size=+20}测试通过{/size}"
    window auto

    pause 2.0

    stop music fadeout 1.0

    hide tester happy

    menu:
        "重新开始":
            jump start
        "退出":
            return

label play_sound_test:

    play sound "audio/click.ogg"
    "你听到了一个音效。"
    stop sound

    return

label voice_test:

    voice "audio/voice/line01.ogg"
    tester "这是带配音的对话。"

    return

label queue_test:

    play music "audio/bgm_calm.ogg"
    queue music "audio/bgm_tense.ogg"
    "BGM 会在当前曲结束后自动切换。"

    return

label call_test:

    "接下来调用子场景。"
    call play_sound_test
    "子场景返回了。"

    return

label player_input_test:

    $ player_name = renpy.input("请输入你的名字：", length=10).strip()
    if player_name == "":
        $ player_name = "无名"
    tester "你好，[player_name]！"

    return

label structured_variable_test:

    "———— 结构化变量指令测试 ————"

    $ money = 200
    $ score += 1
    $ health -= 5
    $ has_key = True
    $ has_key = False
    $ love = 10

    if score == 20:
        "分数正好 20。"
    else:
        "分数不是 20。"

    if score != 0:
        "分数不为 0。"

    if score > 15:
        "分数大于 15。"

    if score >= 11:
        "分数大于等于 11。"

    if health < 50:
        "生命值小于 50。"

    if health <= 45:
        "生命值小于等于 45。"

    "———— 复合条件测试 ————"

    if score >= 10 and score < 50:
        "分数在 10~49 之间。"

    if score == 100 or health <= 0:
        "满分或没血了。"

    if score >= 11 and score <= 50 and love != 0:
        "三连复合条件。"

    if money >= 100 and love >= 5:
        "复杂条件仍用传统 if。"
        return

    return
