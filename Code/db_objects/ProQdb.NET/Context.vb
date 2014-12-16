Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ContextDB_NET.ContextDB")> Public Class ContextDB
	' TODO: Declare local ADO Recordset object. For example:
    'Private WithEvents rs as dataset
	
	Private Sub Class_GetDataMember(ByRef DataMember As String, ByRef Data As Object)
		' TODO:  Return the appropriate recordset based on DataMember. For example:
		
		'Select Case DataMember
		'Case ""             ' Default
		'    Set Data = Nothing
		'Case Else           ' Default
		'    Set Data = rs
		'End Select
	End Sub
End Class