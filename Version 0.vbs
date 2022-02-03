set shell = createobject("wscript.shell")
strtimes = 1
if not isnumeric(strtimes) then
lol=msgbox("Shell obselete")
wscript.quit
end if
msgbox "Vamos a hacer una prueba primero esperando 10 segundos despues de actualizar "
wscript.sleep(1000)
for i=1 to strtimes
set wshshell = wscript.CreateObject("wscript.shell")
wshshell.run "chrome.exe https://google.com"
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Enter}"
wscript.sleep(1000)
Shell.SendKeys "{Tab}"
Shell.SendKeys "{Tab}"
shell.sendkeys(" profe me encuentro un poco mal y he preferido no ir a clase por si acaso, pero estoy aqui en el online" & "")
Shell.SendKeys "{Enter}"
wscript.sleep(1000)
wscript.sleep(75)
next