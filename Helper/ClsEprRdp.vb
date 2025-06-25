Imports System.Runtime.InteropServices
Imports Epr.EprRdp.Globals
Partial Public Class ClsMain
	Public Function GetRemoteCommandsFolder(ByVal root As String) As <MarshalAs(UnmanagedType.BStr)> String
		Return $"{root}{Constants.REMOTE_COMMANDS_FOLDER}\"
	End Function

	Public Function GetRemoteDataFolder(ByVal root As String) As <MarshalAs(UnmanagedType.BStr)> String
		Return $"{root}{Constants.REMOTE_DATA_FOLDER}\"
	End Function

	Public Function GetGviaDataFolder(ByVal root As String) As <MarshalAs(UnmanagedType.BStr)> String
		Return $"{root}{Constants.REMOTE_DATA_FOLDER}\{Constants.GVIA_DATA_FOLDER}\"
	End Function

	Public Function GetEmailsFolder(ByVal root As String) As <MarshalAs(UnmanagedType.BStr)> String
		Return $"{root}{Constants.REMOTE_DATA_FOLDER}\{Constants.EMAILS_FOLDER}\"
	End Function

	Public Function GetGetMailFolder(ByVal root As String) As <MarshalAs(UnmanagedType.BStr)> String
		Return $"{root}{Constants.REMOTE_DATA_FOLDER}\{Constants.GETMAIL_FOLDER}\"
	End Function

	Public Function GetHinDataFolder(ByVal root As String) As <MarshalAs(UnmanagedType.BStr)> String
		Return $"{root}{Constants.REMOTE_DATA_FOLDER}\{Constants.HIN_DATA_FOLDER}\"
	End Function

	Public Function GetFormatFolder(ByVal root As String) As <MarshalAs(UnmanagedType.BStr)> String
		Return $"{root}{Constants.FORMAT_UNIQUE_ID}\"
	End Function

	Public Function GetAddToEmailFolder(ByVal root As String) As <MarshalAs(UnmanagedType.BStr)> String
		Return $"{root}{Constants.REMOTE_COMMANDS_FOLDER}\{Constants.ADDTOEMAIL_FOLDER}\"
	End Function
End Class
