
set objshell = CreateObject("WScript.Shell")
objshell.run "robocopy C:\Users\pengk02\Documents P:\Documents /e /r:0 /dcopy:t /mir"