# -*- coding: utf-8 -*-
"""
Created on Wed Apr  6 16:11:53 2022

@author: hirkatayama
"""

import win32com.client as wincl
voice = wincl.Dispatch("SAPI.SpVoice")
voice.Speak("Pythonでこんなことできます。そぞら＠プログラミングに夢中すぎる会社員")


# speed
import win32com.client as wincl
voice = wincl.Dispatch("SAPI.SpVoice")

voice.Volume = 50 # [0 to 100]
voice.Rate = -5 # [-10 to 10]

voice.Speak("音量や速さを調整できます")


# time
import win32com.client as wincl
import datetime
import locale
locale.setlocale(locale.LC_CTYPE, "Japanese_Japan.932")

voice = wincl.Dispatch("SAPI.SpVoice")

voice.Volume = 50 # [0 to 100]
voice.Rate = -3 # [-10 to 10]

voice.Speak(datetime.datetime.now().strftime('%H時%M分です'))