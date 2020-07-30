# Class_FTP
 AutoHotkey wrapper for FTP Sessions ([msdn-docs](https://docs.microsoft.com/en-us/windows/win32/wininet/ftp-sessions))


## Examples

**Writes a file to the server.**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
FTP.PutFile(hSession, "C:\Temp\testfile.txt", "testfile.txt")
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```

**Retrieves a file from the server.**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
FTP.GetFile(hSession, "testfile.txt", "C:\Temp\testfile.txt")
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```

**Retrieves the file size of the requested FTP resource.**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
MsgBox % FTP.GetFileSize(hSession, "testfile.txt")
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```

**Deletes a file from the server.**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
FTP.DeleteFile(hSession, "testfile.txt")
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```

**Creates a new directory on the server.**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
FTP.CreateDirectory(hSession, "Test_Folder")
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```

**Changes the client's current directory on the server.**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
FTP.SetCurrentDirectory(hSession, "Test_Folder")
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```

**Returns the client's current directory on the server.**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
MsgBox % FTP.GetCurrentDirectory(hSession)
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```

**Deletes a directory on the server.**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
FTP.RemoveDirectory(hSession, "Test_Folder")
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```

**Enumerate all Files in root directory. (!!! EXPERIMENTAL !!!)**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
for k, File in FTP.FindFiles(hSession)
	MsgBox % File.FileName
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```

**Enumerate all Files in a subdirectory. (!!! EXPERIMENTAL !!!)**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
for k, File in FTP.FindFiles(hSession, "/Folder 2")
	MsgBox % File.FileName
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```

**Enumerate all Folders in root directory. (!!! EXPERIMENTAL !!!)**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
for k, Folder in FTP.FindFolders(hSession)
	MsgBox % Folder.FileName
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```

**Enumerate all Folders in a subdirectory. (!!! EXPERIMENTAL !!!)**
```AutoHotkey
hFTP := FTP.Open("AHK-FTP")
hSession := FTP.Connect(hFTP, "ftp.example.com", 21, "user", "passwd")
for k, Folder in FTP.FindFolders(hSession, "/Folder 2")
	MsgBox % Folder.FileName
FTP.Disconnect(hSession)
FTP.Close(hFTP)
```


## Questions / Bugs / Issues
Report any bugs or issues on the [AHK Thread](https://www.autohotkey.com/boards/viewtopic.php?f=6&t=79142). Same for any questions.


## Copyright and License
[The Unlicense](LICENSE)


## Donations (PayPal)
[Donations are appreciated if I could help you](https://www.paypal.me/smithz)