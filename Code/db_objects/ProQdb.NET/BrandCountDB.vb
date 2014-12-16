Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("BrandCountDB_NET.BrandCountDB")> Public Class BrandCountDB
	'BrandCountDB.cls
	'
	'this class represents a row in the collection of Responses
	'
	'lbailey
	'12 june 2002
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Private m_lngBrandCount As Integer
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'-----------------------------------------------------------------
	'Standard Accessor Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	Public Function GetBrandCount() As String
		GetBrandCount = CStr(m_lngBrandCount)
	End Function
	'-----------------------------------------------------------------
	'Standard manipulator Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	Public Sub SetBrandCount(ByRef lngBrandCount As String)
		m_lngBrandCount = CInt(lngBrandCount)
	End Sub
End Class