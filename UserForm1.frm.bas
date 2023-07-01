in file: word/vbaProject.bin - OLE stream: 'VBA/UserForm1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Private Sub UserForm_Click()

End Sub
VBA Stomping detection is experimental: please report any false positive/negative at https://github.com/decalage2/oletools/issues

+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |UserForm_Click      |Runs when the file is opened and ActiveX     |
|          |                    |objects trigger events                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|put                 |May write to a file (if combined with Open)  |
|Suspicious|output              |May write to a file (if combined with Open)  |
|Suspicious|CopyFile            |May copy a file                              |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|vbNormalFocus       |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|run                 |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|create              |May execute file or a system command through |
|          |                    |WMI                                          |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|GetObject           |May get an OLE object with a running instance|
|Suspicious|WINDOWS             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|ChrW                |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |https://stackoverflo|URL                                          |
|          |w.com/questions/7239|                                             |
|          |328/how-to-find-    |                                             |
|          |numbers-from-a-     |                                             |
|          |string#7239408      |                                             |
|IOC       |https://www.gmayor.c|URL                                          |
|          |om                  |                                             |
|IOC       |http://github.com/jd|URL                                          |
|          |dev273/chatgpt-word-|                                             |
|          |macro               |                                             |
|IOC       |https://api.openai.c|URL                                          |
|          |om/v1/chat/completio|                                             |
|          |ns                  |                                             |
|IOC       |https://docs.microso|URL                                          |
|          |ft.com/en-us/office/|                                             |
|          |vba/api/word.wdcolor|                                             |
|          |index               |                                             |
|IOC       |explorer.exe        |Executable file name                         |
|Suspicious|VBA Stomping        |VBA Stomping was detected: the VBA source    |
|          |                    |code and P-code are different, this may have |
|          |                    |been used to hide malicious code             |
+----------+--------------------+---------------------------------------------+
