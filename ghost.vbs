set shell = createobject("wscript.shell") 
strtext = inputbox("What would you like the message to be")
strtimes = inputbox ("How many times would you like you type it?")
strtime = inputbox ("How much delay between types (in miliseconds)?")
if not isnumeric(strtimes) then
lol=msgbox("Please write a NUMBER nextime") 
wscript.quit
end if
msgbox "After you click ok the message will start in 5 seconds "
wscript.sleep(3000)
for i=1 to strtimes
shell.sendkeys(strtext & "")
Shell.SendKeys "{Enter}"
wscript.sleep(strtime)
next
