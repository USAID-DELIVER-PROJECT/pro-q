Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("BabyDB_NET.BabyDB")> Public Class BabyDB
	'BabyDB.cls
	'
	'this class represents a row in the collection
	'
	'lblanken
	'1 oct 2002
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Private m_strName As String
	Private m_strID As String
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'-----------------------------------------------------------------
	'Standard Accessor Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	Public Function GetName() As String
		GetName = m_strName
	End Function
	
	Public Function GetID() As String
		GetID = m_strID
	End Function
	
	'-----------------------------------------------------------------
	'Standard manipulator Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	Public Sub SetName(ByRef strName As String)
		m_strName = strName
	End Sub
	
    Public Sub SetID(ByRef strID As String)
        m_strID = strID
    End Sub
End Class