this windows shell extension adds an "open in Command prompt" entry to the context menu of a folder or desktop background, similar to the existing "open in Terminal" option. also you can choose either to open as admin or current user.


the extension comes as .NET dll, so you'll have to install it using RegAsm.exe:
* go to C:\Windows\Microsoft.NET\\**Framework**\\(version) on 32-bit
   or to C:\Windows\Microsoft.NET\\**Framework64**\\(version) on 64-bit
* run in cmd:

```
 regasm OpenWithCmdExt.dll /codebase
```
or

```
regasm OpenWithCmdExt.dll /c
```