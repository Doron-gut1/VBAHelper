Imports System.Reflection
Imports System.Runtime.InteropServices
Imports RGiesecke.DllExport

Public Class ClsGateway

	<DllExport()>
	Public Shared Function GetInstance() As <MarshalAs(UnmanagedType.IDispatch)> Object
		AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf AssemblyResolve
		Return New ClsMain()
	End Function

End Class
