; ===============================================================================================================================
; AutoHotkey wrapper for FTP Sessions
;
; Author ....: jNizM
; Released ..: 2020-07-26
; Modified ..: 2020-07-26
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


	Open(Agent, Proxy := "", ProxyBypass := "")
	{
		if (hInternet := this.InternetOpen(Agent, Proxy, ProxyBypass))
			return hInternet
		return false
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


	GetFileSize(hConnect, FileName, Format := "auto", Suffix := false)
	{
		static GENERIC_READ := 0x80000000

		if (hFile := this.OpenFile(hConnect, FileName, GENERIC_READ))
		{
			VarSetCapacity(FileSizeHigh, 8)
			if (FileSizeLow := DllCall("wininet\FtpGetFileSize", "ptr", hFile, "uint*", FileSizeHigh, "uint"))
			{
				this.InternetCloseHandle(hFile)
				return this.FormatBytes(FileSizeLow + (FileSizeHigh << 32), Format, Suffix)
			}
			this.InternetCloseHandle(hFile)
		}
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
														 , "ushort", INTERNET_DEFAULT_FTP_PORT ;(Port = 21) ? INTERNET_DEFAULT_FTP_PORT : Port
														 , "ptr",    &UserName ;(UserName) ? &UserName : 0
														 , "ptr",    &Password ;(Password) ? &Password : 0
														 , "uint",   INTERNET_SERVICE_FTP
														 , "uint",   INTERNET_FLAG_PASSIVE ;(FTP_PASV) ? INTERNET_FLAG_PASSIVE : 0
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
													   , "uint", (Proxy) ? INTERNET_OPEN_TYPE_PROXY : INTERNET_OPEN_TYPE_DIRECT
													   , "ptr",  (Proxy) ? &Proxy : 0
													   , "ptr",  (ProxyBypass) ? &ProxyBypass : 0
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


	FormatBytes(bytes, mode := "auto", suffix := false)
	{
		static SFBS_FLAGS_ROUND_TO_NEAREST_DISPLAYED_DIGIT    := 0x0001
		static SFBS_FLAGS_TRUNCATE_UNDISPLAYED_DECIMAL_DIGITS := 0x0002
		static S_OK := 0

		if (mode = "auto")
		{
			size := VarSetCapacity(buf, 1024, 0)
			if (DllCall("shlwapi\StrFormatByteSizeEx", "int64", bytes, "int", SFBS_FLAGS_ROUND_TO_NEAREST_DISPLAYED_DIGIT, "str", buf, "uint", size) = S_OK)
				output := buf
		}
		else if (mode = "kilobytes" || mode = "kb")
			output := Round(bytes / 1024, 2) . (suffix ? " KB" : "")
		else if (mode = "megabytes" || mode = "mb")
			output := Round(bytes / 10204**2, 2) . (suffix ? " MB" : "")
		else if (mode = "gigabytes" || mode = "gb")
			output := Round(bytes / 10204**3, 2) . (suffix ? " GB" : "")
		else if (mode = "terabytes" || mode = "tb")
			output := Round(bytes / 10204**4, 2) . (suffix ? " TB" : "")
		else
			output := Round(bytes, 2) . (suffix ? " Bytes" : "")
		return output
	}

}

; ===============================================================================================================================