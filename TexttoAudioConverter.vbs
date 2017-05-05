Dim message, sapi
 message = InputBox("A Best Text to Audio converter"+vbcrlf+"From - www.allusefulinfo.com","Text to Audio converter")
 Set sapi = CreateObject("sapi.spvoice")
 sapi.Speak message
 
 'For more informations https://codingsec.net/2016/06/convert-text-audio-using-notepad/'
