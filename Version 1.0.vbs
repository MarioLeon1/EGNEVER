set shell = createobject("wscript.shell")
strtimes = 1
if not isnumeric(strtimes) then
lol=msgbox("Shell obselete")
wscript.quit
end if
msgbox "Vamos a hacer una prueba primero esperando 10 segundos despues de actualizar "
Shell.run "chrome.exe https://google.com"
wscript.sleep(3000)
for i=1 to strtimes
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(500)
Shell.SendKeys "{Tab}"
wscript.sleep(10)
Shell.SendKeys "{Tab}"
wscript.sleep(100)
Shell.SendKeys "{Tab}"
wscript.sleep(100)
shell.sendkeys("https://passwords.google.com/" & "")
Shell.SendKeys "{Enter}"
wscript.sleep(2000)
Shell.SendKeys "{Tab}"
Shell.SendKeys "{DOWN}"
Shell.SendKeys "{DOWN}"
Shell.SendKeys "{Enter}"
wscript.sleep(2000)
shell.sendkeys("gatuno20001" & "")
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Enter}"
wscript.sleep(4000)
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Enter}"
Shell.SendKeys "%a"
Shell.SendKeys "^c" 
next