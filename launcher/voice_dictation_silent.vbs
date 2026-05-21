' voice_dictation_silent.vbs
'
' Silent launcher для voice dictation: стартует pythonw.exe из venv
' без окна консоли. Используется для autostart и для ярлыков
' (см. tools/install_shortcut.ps1, tools/install_autostart.ps1).
'
' Допущения:
'   - venv лежит внутри репо: <repo>\.venv  (фоллбэк — %USERPROFILE%\.venvs\whisper)
'   - сам репозиторий лежит по %USERPROFILE%\.claude\skills\Whisper-Skill
'   - в venv установлены зависимости диктовки (sounddevice, pynput, ...)
'
' Если эти пути отличаются — поправь переменные ниже.

Option Explicit

Dim WshShell, fso, scriptDir, repoRoot, pythonw, pythonwAlt, cmd
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Репо вычисляем относительно расположения этого .vbs:
'   <repo>/launcher/voice_dictation_silent.vbs  →  <repo> = ../
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
repoRoot = fso.GetParentFolderName(scriptDir)

' 1) основной путь — venv внутри репо
pythonw = repoRoot & "\.venv\Scripts\pythonw.exe"
' 2) фоллбэк — старый путь ~/.venvs/whisper
pythonwAlt = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\.venvs\whisper\Scripts\pythonw.exe"

If Not fso.FileExists(pythonw) Then
    If fso.FileExists(pythonwAlt) Then
        pythonw = pythonwAlt
    Else
        MsgBox "pythonw.exe not found:" & vbCrLf & pythonw & vbCrLf & pythonwAlt & vbCrLf & vbCrLf & _
               "Создай venv в одном из этих путей или поправь .vbs.", _
               vbCritical, "Whisper Voice — launcher"
        WScript.Quit 1
    End If
End If

' UTF-8 для stdout/stderr — иначе print с эмодзи падает на cp1251.
WshShell.Environment("PROCESS").Item("PYTHONUTF8") = "1"
WshShell.Environment("PROCESS").Item("PYTHONIOENCODING") = "utf-8"

WshShell.CurrentDirectory = repoRoot
cmd = """" & pythonw & """ -m examples.voice_dictation"
' 0 = окно скрыто, False = не ждать завершения.
WshShell.Run cmd, 0, False
