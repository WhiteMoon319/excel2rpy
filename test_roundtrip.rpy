define narrator = Character("")
define tester = Character("测试员")
default score = 0
default has_key = False

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
