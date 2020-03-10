Attribute VB_Name = "ModuleLoadLanguage"
'================================================================================

'================================================================================

Public Sub LoadLanguageENG()
    FormMainWindow.setlanguage = "ENG"

    FormMainWindow.Caption = "Fight Landlord Card Game Assistant　v1.00　by Sam Toki"

    FormMainWindow.MenuLanguageENG.Checked = True
    FormMainWindow.MenuLanguageCHS.Checked = False
    FormMainWindow.MenuLanguageJPN.Checked = False

    FormMainWindow.MenuDoubler.Caption = "Doubler"
    FormMainWindow.MenuDoublerUndo.Caption = "←　Undo"
    FormMainWindow.MenuDoublerReset.Caption = "＊　Reset"
    FormMainWindow.MenuDice.Caption = "Dice"
    FormMainWindow.MenuDiceReset.Caption = "＊　Reset"
    FormMainWindow.MenuDiceRoll.Caption = "N/A"
    FormMainWindow.MenuAbout.Caption = "About..."
    FormMainWindow.MenuEXIT.Caption = "EXIT"
    If FormMainWindow.soundswitch = True Then FormMainWindow.MenuSoundSwitch.Caption = "Sound ON" Else FormMainWindow.MenuSoundSwitch.Caption = "Sound OFF"

    FormMainWindow.FrameDoubler.Caption = "DOUBLER"
    FormMainWindow.FrameDoubler.Font = "Arial"
    FormMainWindow.FrameDice.Caption = "DICE"
    FormMainWindow.FrameDice.Font = "Arial"
    FormMainWindow.CmdDiceRoll.Font = "Arial"
    FormMainWindow.LabelDiceNumber1.Font = "Arial"
    FormMainWindow.LabelDiceNumber2.Font = "Arial"
    Call FormMainWindow.DoublerRefresher
    Call FormMainWindow.DiceRefresher
End Sub

'================================================================================

Public Sub LoadLanguageCHS()
    FormMainWindow.setlanguage = "CHS"

    FormMainWindow.Caption = "线下斗地主棋牌辅助工具　v1.00　Sam Toki 制作"

    FormMainWindow.MenuLanguageENG.Checked = False
    FormMainWindow.MenuLanguageCHS.Checked = True
    FormMainWindow.MenuLanguageJPN.Checked = False

    FormMainWindow.MenuDoubler.Caption = "倍数"
    FormMainWindow.MenuDoublerUndo.Caption = "←　撤销"
    FormMainWindow.MenuDoublerReset.Caption = "＊　重置"
    FormMainWindow.MenuDice.Caption = "癞子"
    FormMainWindow.MenuDiceReset.Caption = "＊　重置"
    FormMainWindow.MenuDiceRoll.Caption = "N/A"
    FormMainWindow.MenuAbout.Caption = "关于..."
    FormMainWindow.MenuEXIT.Caption = "退出"
    If FormMainWindow.soundswitch = True Then FormMainWindow.MenuSoundSwitch.Caption = "声音 开" Else FormMainWindow.MenuSoundSwitch.Caption = "声音 关"

    FormMainWindow.FrameDoubler.Caption = "倍数"
    FormMainWindow.FrameDoubler.Font = "SimSun"
    FormMainWindow.FrameDice.Caption = "癞子"
    FormMainWindow.FrameDice.Font = "SimSun"
    FormMainWindow.CmdDiceRoll.Font = "SimHei"
    FormMainWindow.LabelDiceNumber1.Font = "SimHei"
    FormMainWindow.LabelDiceNumber2.Font = "SimHei"
    Call FormMainWindow.DoublerRefresher
    Call FormMainWindow.DiceRefresher
End Sub

'================================================================================

Public Sub LoadLanguageJPN()
    FormMainWindow.setlanguage = "JPN"

    FormMainWindow.Caption = "Fight Landlord カードゲームアシスタント　v1.00　by Sam Toki"

    FormMainWindow.MenuLanguageENG.Checked = False
    FormMainWindow.MenuLanguageCHS.Checked = False
    FormMainWindow.MenuLanguageJPN.Checked = True

    FormMainWindow.MenuDoubler.Caption = "番数"
    FormMainWindow.MenuDoublerUndo.Caption = "←　取り消す"
    FormMainWindow.MenuDoublerReset.Caption = "＊　リセット"
    FormMainWindow.MenuDice.Caption = "サイコロ"
    FormMainWindow.MenuDiceReset.Caption = "＊　リセット"
    FormMainWindow.MenuDiceRoll.Caption = "N/A"
    FormMainWindow.MenuAbout.Caption = "について..."
    FormMainWindow.MenuEXIT.Caption = "终了"
    If FormMainWindow.soundswitch = True Then FormMainWindow.MenuSoundSwitch.Caption = "音声 オン" Else FormMainWindow.MenuSoundSwitch.Caption = "音声 オフ"

    FormMainWindow.FrameDoubler.Caption = "番数"
    FormMainWindow.FrameDoubler.Font = "MS UI Gothic"
    FormMainWindow.FrameDice.Caption = "サイコロ"
    FormMainWindow.FrameDice.Font = "MS UI Gothic"
    FormMainWindow.CmdDiceRoll.Font = "MS UI Gothic"
    FormMainWindow.LabelDiceNumber1.Font = "MS UI Gothic"
    FormMainWindow.LabelDiceNumber2.Font = "MS UI Gothic"
    Call FormMainWindow.DoublerRefresher
    Call FormMainWindow.DiceRefresher
End Sub

'================================================================================

'================================================================================
