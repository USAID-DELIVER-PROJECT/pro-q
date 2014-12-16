Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("SelectedQuantityDB_NET.SelectedQuantityDB")> Public Class SelectedQuantityDB
	'+
	'SelectedQuantityDB.cls
	'
	'each record in this table represents a distinct brand associated
	'with its quantity info.  these objects are joins of other tables,
	'and don't exist directly in the back end.  in other words, this
	'is not a table wrapper
	'
	'lbailey
	'6 june 2002
	'-
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private properties
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'properties
	Private m_strID As Object
	Private m_strQuantificationID As Object
	Private m_strSelectedBrandID As Object
	Private m_strBrandID As Object
	Private m_strKitID As Object
	Private m_intTestsPerKit As Object
	Private m_dblKitCost As Object
	Private m_lngCount As Object
	Private m_intCategory As Object
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'+
	'standard manipulator functions
	'
	'lbailey
	'16 may 2002
	'-
	Public Sub SetID(ByRef strID As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strID = strID
	End Sub
	
	Public Sub SetQuantificationID(ByRef strQuantificationID As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strQuantificationID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strQuantificationID = strQuantificationID
	End Sub
	
	Public Sub SetSelectedBrandID(ByRef strSelectedBrandID As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strSelectedBrandID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strSelectedBrandID = strSelectedBrandID
	End Sub
	
	Public Sub SetBrandID(ByRef strBrandID As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strBrandID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strBrandID = strBrandID
	End Sub
	
	Public Sub SetKitID(ByRef strKitID As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strKitID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strKitID = strKitID
	End Sub
	
	Public Sub SetTestsPerKit(ByRef intTestsPerKit As Short)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_intTestsPerKit. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_intTestsPerKit = intTestsPerKit
	End Sub
	
	Public Sub SetKitCost(ByRef dblKitCost As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_dblKitCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_dblKitCost = dblKitCost
	End Sub
	
	Public Sub SetCount(ByRef lngCount As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_lngCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_lngCount = lngCount
	End Sub
	
	Public Sub SetCategory(ByRef intCategory As Short)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_intCategory. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_intCategory = intCategory
	End Sub
	
	
	
	'+
	'standard accessor
	'
	'lbailey
	'6 june 2002
	'-
	Public Function GetID() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetID = m_strID
	End Function
	
	Public Function GetQuantificationID() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strQuantificationID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetQuantificationID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetQuantificationID = m_strQuantificationID
	End Function
	
	Public Function GetSelectedBrandID() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strSelectedBrandID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetSelectedBrandID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetSelectedBrandID = m_strSelectedBrandID
	End Function
	
	Public Function GetBrandID() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strBrandID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetBrandID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetBrandID = m_strBrandID
	End Function
	
	Public Function GetKitID() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strKitID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetKitID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetKitID = m_strKitID
	End Function
	
	Public Function GetTestsPerKit() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_intTestsPerKit. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetTestsPerKit. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetTestsPerKit = m_intTestsPerKit
	End Function
	
	Public Function GetKitCost() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_dblKitCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetKitCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetKitCost = m_dblKitCost
	End Function
	
	Public Function GetCount() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_lngCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCount = m_lngCount
	End Function
	
	Public Function GetCategory() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_intCategory. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetCategory. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCategory = m_intCategory
	End Function
End Class