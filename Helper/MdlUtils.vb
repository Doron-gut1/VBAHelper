Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.IO
Imports System.Reflection

Friend Module MdlUtils

	Function AssemblyResolve(sender As Object, e As ResolveEventArgs) As Assembly
		Dim resourceFullName As String

		resourceFullName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) & "\" & String.Format("{0}.dll", e.Name.Split(","c)(0))  ' Environment.CurrentDirectory & "\" & String.Format("{0}.dll", e.Name.Split(","c)(0))

		Return Assembly.LoadFile(resourceFullName)
	End Function

	Function ExtractMustInnerException(ByRef ExceptionIn As Exception) As Exception
		Dim Ret As Exception = ExceptionIn

		While Ret.InnerException IsNot Nothing
			Ret = Ret.InnerException
		End While

		Return Ret
	End Function

	Function ToBytes(ByVal instance As Stream) As Byte()
		Dim capacity As Integer = If(instance.CanSeek, Convert.ToInt32(instance.Length), 0)

		Using result As New MemoryStream(capacity)
			Dim readLength As Integer
			Dim buffer(4096) As Byte

			Do
				readLength = instance.Read(buffer, 0, buffer.Length)
				result.Write(buffer, 0, readLength)
			Loop While readLength > 0

			Return result.ToArray()
		End Using
	End Function

	Public Function TraceException(exceptionMsg As String,
		<System.Runtime.CompilerServices.CallerMemberName> Optional memberName As String = Nothing,
		<System.Runtime.CompilerServices.CallerFilePath> Optional sourcefilePath As String = Nothing,
		<System.Runtime.CompilerServices.CallerLineNumber()> Optional sourceLineNumber As Integer = 0) As String

		'System.Diagnostics.Trace.WriteLine("message: " & message)
		'System.Diagnostics.Trace.WriteLine("member name: " & memberName)
		'System.Diagnostics.Trace.WriteLine("source file path: " & sourcefilePath)
		'System.Diagnostics.Trace.WriteLine("source line number: " & sourceLineNumber)
		Return $"{memberName} Exception - {exceptionMsg} ({New FileInfo(sourcefilePath).Name} line:{sourceLineNumber})"
	End Function

	Public Sub AddDbErrors(ByVal sqlConnString As String, ByVal errDesc As String,
			<System.Runtime.CompilerServices.CallerMemberName> ByVal Optional memberName As String = Nothing,
			<System.Runtime.CompilerServices.CallerFilePath> ByVal Optional sourcefilePath As String = Nothing,
			<System.Runtime.CompilerServices.CallerLineNumber()> ByVal Optional sourceLineNumber As Integer = 0, ByVal Optional user As String = "")
		AddDbErrors(sqlConnString, user, "0", errDesc, sourcefilePath, "", sourceLineNumber, "", 0, "", Environment.MachineName, "", "", 0, memberName)
	End Sub

	Private Sub AddDbErrors(ByVal sqlConnString As String, ByVal user As String, ByVal errnum As String, ByVal errdesc As String, ByVal modulname As String, ByVal objectname As String, ByVal errline As Integer, ByVal strinrow As String, ByVal moduletype As Integer, ByVal moredtls As String, ByVal comp As String, ByVal ver As String, ByVal callStack As String, ByVal jobnum As Integer, ByVal subname As String)
		Try

			Using sqlConn As SqlConnection = New SqlConnection(sqlConnString)
				sqlConn.Open()

				Using sqlCmd As SqlCommand = New SqlCommand("AddDbErrors", sqlConn)
					sqlCmd.CommandType = System.Data.CommandType.StoredProcedure
					sqlCmd.Parameters.Add("@user", SqlDbType.NVarChar, 50).Value = user
					sqlCmd.Parameters.Add("@errnum", SqlDbType.NText).Value = errnum
					sqlCmd.Parameters.Add("@errdesc", SqlDbType.NVarChar, 255).Value = (If(errdesc.Length > 255, errdesc.Substring(0, 255), errdesc))
					sqlCmd.Parameters.Add("@modulname", SqlDbType.NVarChar, 50).Value = modulname
					sqlCmd.Parameters.Add("@objectname", SqlDbType.NVarChar, 50).Value = objectname
					sqlCmd.Parameters.Add("@errline", SqlDbType.Int).Value = errline
					sqlCmd.Parameters.Add("@strinrow", SqlDbType.NText).Value = strinrow
					sqlCmd.Parameters.Add("@moduletype", SqlDbType.Int).Value = moduletype
					sqlCmd.Parameters.Add("@moredtls", SqlDbType.NVarChar, -1).Value = moredtls
					sqlCmd.Parameters.Add("@comp", SqlDbType.NVarChar, 100).Value = comp
					sqlCmd.Parameters.Add("@Ver", SqlDbType.NVarChar, 100).Value = ver
					sqlCmd.Parameters.Add("@CallStack", SqlDbType.NVarChar, -1).Value = callStack
					sqlCmd.Parameters.Add("@jobnum", SqlDbType.Int).Value = jobnum
					sqlCmd.Parameters.Add("@subname", SqlDbType.NVarChar, 100).Value = subname
					sqlCmd.ExecuteNonQuery()
				End Using
			End Using

		Catch ex As Exception
		End Try
	End Sub

	Public Function TruncateString(str As String, length As Integer) As String
		' If argument is too big, return the original string.
		' ... Otherwise take a substring from the string's start index.

		If length > str.Length Then
			Return str
		Else
			Return str.Substring(0, length)
		End If
	End Function

End Module
