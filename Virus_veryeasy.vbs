set wshshell =wscipt.CreatorObject("wshell")
do
wscript.sleep 100
wshshell.sendkeys "{CAPSLOCK}
wscript.sleep 100
wshshell.sendkeys "{NUMLOCK}
wscript.sleep 100
wshshell.sendkeys "{SCROOLLOCK}"
loop