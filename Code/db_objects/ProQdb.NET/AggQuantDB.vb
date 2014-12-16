Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("AggQuantDB_NET.AggQuantDB")> Public Class AggQuantDB
	'AggQuantDB.cls
	'
	'this class represents a row in the collection of Answers
	'
	'lblankenship
	'16 january 2003
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Private m_strQuantID As String
	Private m_strUseID As String
	Private m_strQuantName As String
	Private m_lngScriptUseID As Integer
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'-----------------------------------------------------------------
	'Standard Accessor Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	Public Function GetQuantID() As String
		GetQuantID = m_strQuantID
	End Function
	
	Public Function GetUseID() As String
		GetUseID = m_strUseID
	End Function
	
	Public Function GetQuantName() As String
		GetQuantName = m_strQuantName
	End Function
	
	Public Function GetScriptUseID() As Integer
		GetScriptUseID = m_lngScriptUseID
	End Function
	
	'-----------------------------------------------------------------
	'Standard manipulator Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	Public Sub SetQuantID(ByRef strQuantID As String)
		m_strQuantID = strQuantID
	End Sub
	
	Public Sub SetUseID(ByRef strUseID As String)
		m_strUseID = strUseID
	End Sub
	
	Public Sub SetQuantName(ByRef strQuantName As String)
		m_strQuantName = strQuantName
	End Sub
	
	Public Sub SetScriptUseID(ByRef lngScriptUseID As Integer)
		m_lngScriptUseID = lngScriptUseID
	End Sub
End Class