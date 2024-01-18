'https://www.nextofwindows.com/windows-trick-how-to-make-your-computer-to-speak-out-time-at-every-hour
'set up Task Scheduler to run this on the hour

Dim what, voice
what = "It is " & hour(time) & " O'clock"
Set voice = CreateObject("sapi.spvoice")
voice.Speak what
