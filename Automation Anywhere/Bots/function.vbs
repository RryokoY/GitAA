
Option Explicit

Function Result()
 Result = Msgbox("WSH Version: " & WScript.Version & "(" & WScript.BuildVersion & ")")
End Function