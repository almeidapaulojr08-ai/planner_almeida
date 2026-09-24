' Abre o iniciar_bot.cmd sem janela (a tarefa agendada chama este arquivo)
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.Run "cmd /c """ & dir & "\iniciar_bot.cmd""", 0, False
