Dim message, sapi
message=InputBox("What do you want me to say?","Type here the text what you want !!!")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message 