Set WinScriptHost = CreateObject( "WScript.shell" )
WinScriptHost.Run Chr(34) & "Client.bat" & Chr(34), 0
Set WinScriptHost = Nothing