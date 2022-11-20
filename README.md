# Stardisk's Player
Very simple audio player on VB6.

I think that all modern audioplayers are too big and have a lot of excessive functions. <br>
I don't need functions like playlists, music library, equalizer, bunch of playing animations, sorting by artist/album/genre/goddamn devil/etc.<br>
I just want to play music files in the current folder.<br><br>

So, I made my own audio player.<br><br>

![stardisk player v1 1](https://user-images.githubusercontent.com/24385735/202901770-50b563ec-12d9-4492-9d22-822f9bdcc558.png)<br><br>

It was written in 2013 or something, but in 2022 I've found it's sources, fix some bugs and post in there.<br><br>

## Features ##
- No installation required: the player is one .exe file (and 2 .ini files which will be created while running)
- Extremely lightweight: ~200 Kilobytes!
- Extremely fast launch.
- Works in any Windows version.
- Integrated file explorer. Just select a folder and player will play all the mp3 files inside.
- Fast file management - you can move, rename or delete mp3-file which is currently playing. Quite useful when you try to deal with folder with a lot of different mp3-files, "Downloads" for example.
- You can access local network resources by clicking "..." button near the drive selector and type in the window something like "\\\\192.168.88.247\music".
- Has compact mode which is always on top, movable and takes very little space on the screen.
- Strange function - "Play a part". You can choose start time and end time of any file and loop this part. If you don't need, you can hide it.
<br><br>
## Problems ##
- Non-standard characters in filenames like letters with umlauts or non-latin alphabets are not supported and these files couldn't be played.<br>
- 32bit.<br>
- Requires **MSVBVM60.DLL** installed in your system. But usually this file is located in C:\Windows\system32.<br>
- Requires Windows Media Player installed, because it uses its API. If you don't have it, you can put **wmp.dll** in the folder with .exe file. <br>
- Russian interface only.<br><br>

I hope that you find this audio player useful.
