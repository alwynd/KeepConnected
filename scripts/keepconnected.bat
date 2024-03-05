cd ..
del scripts\keepconnected.stop /Q
call x64\Release\KeepConnected.exe -fortclient -debug -trace > scripts\keepconnected.log
rem call x64\Release\KeepConnected.exe -fortclient > scripts\keepconnected.log

