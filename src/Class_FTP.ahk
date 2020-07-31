; ===============================================================================================================================
; AutoHotkey wrapper for FTP Sessions
;
; Author ....: jNizM
; Released ..: 2020-07-26
; Modified ..: 2020-07-31
; Github ....: https://github.com/jNizM/Class_FTP
; Forum .....: https://www.autohotkey.com/boards/viewtopic.php?f=6&t=79142
; ===============================================================================================================================


class FTP
{

	static hWININET := DllCall("LoadLibrary", "str", "wininet.dll", "ptr")


	; ===== PUBLIC METHODS ======================================================================================================

	Close(hInternet)
	{
		if (hInternet)
			if (this.InternetCloseHandle(hInternet))
				return true
		return false
	}


	Connect(hInternet, ServerName, Port := 21, UserName := "", Password := "", FTP_PASV := 1)
	{
		if (hConnect := this.InternetConnect(hInternet, ServerName, Port, UserName, Password, FTP_PASV))
			return hConnect
		return false
	}


	CreateDirectory(hConnect, Directory)
	{
		if (DllCall("wininet\FtpCreateDirectory", "ptr", hConnect, "ptr", &Directory))
			return true
		return false
	}


	DeleteFile(hConnect, FileName)
	{
		if (DllCall("wininet\FtpDeleteFile", "ptr", hConnect, "ptr", &FileName))
			return true
		return false
	}


	Disconnect(hConnect)
	{
		if (hConnect)
			if (this.InternetCloseHandle(hConnect))
				return true
		return false
	}


	FindFiles(hConnect, SearchFile := "*.*")
	{
		static FILE_ATTRIBUTE_DIRECTORY := 0x10

		Files := []
		find := this.FindFirstFile(hConnect, hEnum, SearchFile)
		if !(find.FileAttr & FILE_ATTRIBUTE_DIRECTORY)
			Files.Push(find)

		while (find := this.FindNextFile(hEnum))
			if !(find.FileAttr & FILE_ATTRIBUTE_DIRECTORY)
				Files.Push(find)
		this.Close(hEnum)
		return Files
	}


	FindFolders(hConnect, SubDirectories := "*.*")
	{
		static FILE_ATTRIBUTE_DIRECTORY := 0x10

		Folders := []
		find := this.FindFirstFile(hConnect, hEnum, SubDirectories)
		if (find.FileAttr & FILE_ATTRIBUTE_DIRECTORY)
			Folders.Push(find)
		while (find := this.FindNextFile(hEnum))
			if (find.FileAttr & FILE_ATTRIBUTE_DIRECTORY)
				Folders.Push(find)
		this.Close(hEnum)
		return Folders
	}


	GetCurrentDirectory(hConnect)
	{
		static MAX_PATH := 260 + 8

		BUFFER_SIZE := VarSetCapacity(CurrentDirectory, MAX_PATH, 0)
		if (DllCall("wininet\FtpGetCurrentDirectory", "ptr", hConnect, "ptr", &CurrentDirectory, "uint*", BUFFER_SIZE))
			return StrGet(&CurrentDirectory)
		return false
	}


	GetFile(hConnect, RemoteFile, NewFile, OverWrite := 0, Flags := 0)
	{
		if (DllCall("wininet\FtpGetFile", "ptr", hConnect, "ptr", &RemoteFile, "ptr", &NewFile, "int", !OverWrite, "uint", 0, "uint", Flags, "uptr", 0))
			return true
		return false
	}


	GetFileSize(hConnect, FileName, SizeFormat := "auto", SizeSuffix := false)
	{
		static GENERIC_READ := 0x80000000

		if (hFile := this.OpenFile(hConnect, FileName, GENERIC_READ))
		{
			VarSetCapacity(FileSizeHigh, 8)
			if (FileSizeLow := DllCall("wininet\FtpGetFileSize", "ptr", hFile, "uint*", FileSizeHigh, "uint"))
			{
				this.InternetCloseHandle(hFile)
				return this.FormatBytes(FileSizeLow + (FileSizeHigh << 32), SizeFormat, SizeSuffix)
			}
			this.InternetCloseHandle(hFile)
		}
		return false
	}


	Open(Agent, Proxy := "", ProxyBypass := "")
	{
		if (hInternet := this.InternetOpen(Agent, Proxy, ProxyBypass))
			return hInternet
		return false
	}


	PutFile(hConnect, LocaleFile, RemoteFile, Flags := 0)
	{
		if (DllCall("wininet\FtpPutFile", "ptr", hConnect, "ptr", &LocaleFile, "ptr", &RemoteFile, "uint", Flags, "uptr", 0))
			return true
		return false
	}


	RemoveDirectory(hConnect, Directory)
	{
		if (DllCall("wininet\FtpRemoveDirectory", "ptr", hConnect, "ptr", &Directory))
			return true
		return false
	}


	RenameFile(hConnect, ExistingFile, NewFile)
	{
		if (DllCall("wininet\FtpRenameFile", "ptr", hConnect, "ptr", &ExistingFile, "ptr", &NewFile))
			return true
		return false
	}


	SetCurrentDirectory(hconnect, Directory)
	{
		if (DllCall("wininet\FtpSetCurrentDirectory", "ptr", hConnect, "ptr", &Directory))
			return true
		return false
	}


	; ===== PRIVATE METHODS =====================================================================================================

	FileAttributes(Attributes)
	{
		static FILE_ATTRIBUTE := { 0x1: "READONLY", 0x2: "HIDDEN", 0x4: "SYSTEM", 0x10: "DIRECTORY", 0x20: "ARCHIVE", 0x40: "DEVICE", 0x80: "NORMAL"
						, 0x100: "TEMPORARY", 0x200: "SPARSE_FILE", 0x400: "REPARSE_POINT", 0x800: "COMPRESSED", 0x1000: "OFFLINE"
						, 0x2000: "NOT_CONTENT_INDEXED", 0x4000: "ENCRYPTED", 0x8000: "INTEGRITY_STREAM", 0x10000: "VIRTUAL"
						, 0x20000: "NO_SCRUB_DATA", 0x40000: "RECALL_ON_OPEN", 0x400000: "RECALL_ON_DATA_ACCESS" }
		GetFileAttributes := []
		for k, v in FILE_ATTRIBUTE
			if (k & Attributes)
				GetFileAttributes.Push(v)
		return GetFileAttributes
	}


	FindData(ByRef WIN32_FIND_DATA, SizeFormat := "auto", SizeSuffix := false)
	{
		static MAX_PATH := 260
		static MAXDWORD := 0xffffffff

		addr := &WIN32_FIND_DATA
		FIND_DATA := []
		FIND_DATA["FileAttr"]          := NumGet(addr + 0, "uint")
		FIND_DATA["FileAttributes"]    := this.FileAttributes(NumGet(addr + 0, "uint"))
		FIND_DATA["CreationTime"]      := this.FileTime(NumGet(addr +  4, "uint64"))
		FIND_DATA["LastAccessTime"]    := this.FileTime(NumGet(addr + 12, "uint64"))
		FIND_DATA["LastWriteTime"]     := this.FileTime(NumGet(addr + 20, "uint64"))
		FIND_DATA["FileSize"]          := this.FormatBytes((NumGet(addr + 28, "uint") * (MAXDWORD + 1)) + NumGet(addr + 32, "uint"), SizeFormat, SizeSuffix)
		FIND_DATA["FileName"]          := StrGet(addr + 44, "utf-16")
		FIND_DATA["AlternateFileName"] := StrGet(addr + 44 + MAX_PATH * (A_IsUnicode ? 2 : 1), "utf-16")
		return FIND_DATA
	}


	FindFirstFile(hConnect, ByRef hFind, SearchFile := "*.*", SizeFormat := "auto", SizeSuffix := false)
	{
		VarSetCapacity(WIN32_FIND_DATA, (A_IsUnicode ? 592 : 320), 0)
		if (hFind := DllCall("wininet\FtpFindFirstFile", "ptr", hConnect, "str", SearchFile, "ptr", &WIN32_FIND_DATA, "uint", 0, "uint*", 0))
			return this.FindData(WIN32_FIND_DATA, SizeFormat, SizeSuffix)
		VarSetCapacity(WIN32_FIND_DATA, 0)
		return false
	}


	FindNextFile(hFind, SearchFile := "*.*", SizeFormat := "auto", SizeSuffix := false)
	{
		VarSetCapacity(WIN32_FIND_DATA, (A_IsUnicode ? 592 : 320), 0)
		if (DllCall("wininet\InternetFindNextFile", "ptr", hFind, "ptr", &WIN32_FIND_DATA))
			return this.FindData(WIN32_FIND_DATA, SizeFormat, SizeSuffix)
		VarSetCapacity(WIN32_FIND_DATA, 0)
		return false
	}


	FileTime(addr)
	{
		this.FileTimeToSystemTime(addr, SystemTime)
		this.SystemTimeToTzSpecificLocalTime(&SystemTime, LocalTime)
		return Format("{:04}{:02}{:02}{:02}{:02}{:02}"
					, NumGet(LocalTime,  0, "ushort")
					, NumGet(LocalTime,  2, "ushort")
					, NumGet(LocalTime,  6, "ushort")
					, NumGet(LocalTime,  8, "ushort")
					, NumGet(LocalTime, 10, "ushort")
					, NumGet(LocalTime, 12, "ushort"))
	}


	FileTimeToSystemTime(FileTime, ByRef SystemTime)
	{
		VarSetCapacity(SystemTime, 16, 0)
		if (DllCall("FileTimeToSystemTime", "int64*", FileTime, "ptr", &SystemTime))
			return true
		return false
	}


	FormatBytes(bytes, SizeFormat := "auto", suffix := false)
	{
		static SFBS_FLAGS_ROUND_TO_NEAREST_DISPLAYED_DIGIT    := 0x0001
		static SFBS_FLAGS_TRUNCATE_UNDISPLAYED_DECIMAL_DIGITS := 0x0002
		static S_OK := 0

		if (SizeFormat = "auto")
		{
			size := VarSetCapacity(buf, 1024, 0)
			if (DllCall("shlwapi\StrFormatByteSizeEx", "int64", bytes, "int", SFBS_FLAGS_ROUND_TO_NEAREST_DISPLAYED_DIGIT, "str", buf, "uint", size) = S_OK)
				output := buf
		}
		else if (SizeFormat = "kilobytes" || SizeFormat = "kb")
			output := Round(bytes / 1024, 2) . (suffix ? " KB" : "")
		else if (SizeFormat = "megabytes" || SizeFormat = "mb")
			output := Round(bytes / 1024**2, 2) . (suffix ? " MB" : "")
		else if (SizeFormat = "gigabytes" || SizeFormat = "gb")
			output := Round(bytes / 1024**3, 2) . (suffix ? " GB" : "")
		else if (SizeFormat = "terabytes" || SizeFormat = "tb")
			output := Round(bytes / 1024**4, 2) . (suffix ? " TB" : "")
		else
			output := Round(bytes, 2) . (suffix ? " Bytes" : "")
		return output
	}


	InternetCloseHandle(hInternet)
	{
		if (DllCall("wininet\InternetCloseHandle", "ptr", hInternet))
			return true
		return false
	}


	InternetConnect(hInternet, ServerName, Port := 21, UserName := "", Password := "", FTP_PASV := 1)
	{
		static INTERNET_DEFAULT_FTP_PORT := 21
		static INTERNET_SERVICE_FTP      := 1
		static INTERNET_FLAG_PASSIVE     := 0x08000000

		if (hConnect := DllCall("wininet\InternetConnect", "ptr",    hInternet
														 , "ptr",    &ServerName
														 , "ushort", (Port = 21 ? INTERNET_DEFAULT_FTP_PORT : Port)
														 , "ptr",    (UserName ? &UserName : 0)
														 , "ptr",    (Password ? &Password : 0)
														 , "uint",   INTERNET_SERVICE_FTP
														 , "uint",   (FTP_PASV ? INTERNET_FLAG_PASSIVE : 0)
														 , "uptr",   0
														 , "ptr"))
			return hConnect
		return false
	}


	InternetOpen(Agent, Proxy := "", ProxyBypass := "")
	{
		static INTERNET_OPEN_TYPE_DIRECT := 1
		static INTERNET_OPEN_TYPE_PROXY  := 3

		if (hInternet := DllCall("wininet\InternetOpen", "ptr",  &Agent
													   , "uint", (Proxy ? INTERNET_OPEN_TYPE_PROXY : INTERNET_OPEN_TYPE_DIRECT)
													   , "ptr",  (Proxy ? &Proxy : 0)
													   , "ptr",  (ProxyBypass ? &ProxyBypass : 0)
													   , "uint", 0
													   , "ptr"))
			return hInternet
		return false
	}


	OpenFile(hConnect, FileName, Access)
	{
		static FTP_TRANSFER_TYPE_BINARY := 2

		if (hFTPSESSION := DllCall("wininet\FtpOpenFile", "ptr", hConnect, "ptr", &FileName, "uint", Access, "uint", FTP_TRANSFER_TYPE_BINARY, "uptr", 0))
			return hFTPSESSION
		return false
	}


	SystemTimeToTzSpecificLocalTime(SystemTime, ByRef LocalTime)
	{
		VarSetCapacity(LocalTime, 16, 0)
		if (DllCall("SystemTimeToTzSpecificLocalTime", "ptr", 0, "ptr", SystemTime, "ptr", &LocalTime))
			return true
		return false
	}

}

; ===============================================================================================================================