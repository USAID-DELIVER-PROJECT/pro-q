Option Strict Off
Option Explicit On
Module basUserName
	' Objects for GetUserName_TSB()
	Declare Function TSB_API_GetUserName Lib "advapi32.dll"  Alias "GetUserNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	
	Public Function GetUserName() As String
		' Comments  : Retrieves the name of the user logged into Windows
		' Parameters: none
		' Returns   : string user name
		'
		Dim lngLen As Integer
		Dim strBuf As String
		
		Const MaxUserName As Short = 255
		
		strBuf = Space(MaxUserName)
		
		lngLen = MaxUserName
		
		If CBool(TSB_API_GetUserName(strBuf, lngLen)) Then
			GetUserName = Left(strBuf, lngLen - 1)
		Else
			GetUserName = ""
		End If
		
	End Function
End Module