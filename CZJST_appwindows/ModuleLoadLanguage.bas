Attribute VB_Name = "ModuleLoadLanguage"
'================================================================================

'================================================================================

Public Sub LoadLanguageENG()
    FormMainWindow.setlanguage = "ENG"

    FormMainWindow.Caption = "Fight Landlord Card Game Assistant��v1.00��by Sam Toki"

    FormMainWindow.MenuLanguageENG.Checked = True
    FormMainWindow.MenuLanguageCHS.Checked = False
    FormMainWindow.MenuLanguageJPN.Checked = False

    FormMainWindow.MenuDoubler.Caption = "Doubler"
    FormMainWindow.MenuDoublerUndo.Caption = "����Undo"
    FormMainWindow.MenuDoublerReset.Caption = "����Reset"
    FormMainWindow.MenuDice.Caption = "Dice"
    FormMainWindow.MenuDiceReset.Caption = "����Reset"
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

    FormMainWindow.Caption = "���¶��������Ƹ������ߡ�v1.00��Sam Toki ����"

    FormMainWindow.MenuLanguageENG.Checked = False
    FormMainWindow.MenuLanguageCHS.Checked = True
    FormMainWindow.MenuLanguageJPN.Checked = False

    FormMainWindow.MenuDoubler.Caption = "����"
    FormMainWindow.MenuDoublerUndo.Caption = "��������"
    FormMainWindow.MenuDoublerReset.Caption = "��������"
    FormMainWindow.MenuDice.Caption = "���"
    FormMainWindow.MenuDiceReset.Caption = "��������"
    FormMainWindow.MenuDiceRoll.Caption = "N/A"
    FormMainWindow.MenuAbout.Caption = "����..."
    FormMainWindow.MenuEXIT.Caption = "�˳�"
    If FormMainWindow.soundswitch = True Then FormMainWindow.MenuSoundSwitch.Caption = "���� ��" Else FormMainWindow.MenuSoundSwitch.Caption = "���� ��"

    FormMainWindow.FrameDoubler.Caption = "����"
    FormMainWindow.FrameDoubler.Font = "SimSun"
    FormMainWindow.FrameDice.Caption = "���"
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

    FormMainWindow.Caption = "Fight Landlord ���`�ɥ��`�ॢ��������ȡ�v1.00��by Sam Toki"

    FormMainWindow.MenuLanguageENG.Checked = False
    FormMainWindow.MenuLanguageCHS.Checked = False
    FormMainWindow.MenuLanguageJPN.Checked = True

    FormMainWindow.MenuDoubler.Caption = "����"
    FormMainWindow.MenuDoublerUndo.Caption = "����ȡ������"
    FormMainWindow.MenuDoublerReset.Caption = "�����ꥻ�å�"
    FormMainWindow.MenuDice.Caption = "��������"
    FormMainWindow.MenuDiceReset.Caption = "�����ꥻ�å�"
    FormMainWindow.MenuDiceRoll.Caption = "N/A"
    FormMainWindow.MenuAbout.Caption = "�ˤĤ���..."
    FormMainWindow.MenuEXIT.Caption = "����"
    If FormMainWindow.soundswitch = True Then FormMainWindow.MenuSoundSwitch.Caption = "���� ����" Else FormMainWindow.MenuSoundSwitch.Caption = "���� ����"

    FormMainWindow.FrameDoubler.Caption = "����"
    FormMainWindow.FrameDoubler.Font = "MS UI Gothic"
    FormMainWindow.FrameDice.Caption = "��������"
    FormMainWindow.FrameDice.Font = "MS UI Gothic"
    FormMainWindow.CmdDiceRoll.Font = "MS UI Gothic"
    FormMainWindow.LabelDiceNumber1.Font = "MS UI Gothic"
    FormMainWindow.LabelDiceNumber2.Font = "MS UI Gothic"
    Call FormMainWindow.DoublerRefresher
    Call FormMainWindow.DiceRefresher
End Sub

'================================================================================

'================================================================================
