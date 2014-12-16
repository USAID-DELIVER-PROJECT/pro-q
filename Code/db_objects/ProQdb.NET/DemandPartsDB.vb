Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("DemandPartsDB_NET.DemandPartsDB")> Public Class DemandPartsDB
	'DemandPartDB Holds a The sub parts of a Demand Calcualtion
	
	Private m_strQuantificationID As String
	Private m_strMethodologyID As String
	Private m_strBrandID As String
	Private m_varA As Object
	Private m_varB As Object
	Private m_varC As Object
	Private m_varD As Object
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'Intialize the object here
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'Terminate the object here
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Accessor Functions
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function getMethodologyID() As String
		getMethodologyID = m_strMethodologyID
	End Function
	
	Public Function getQuantificationID() As String
		getQuantificationID = m_strQuantificationID
	End Function
	
	Public Function GetBrandID() As String
		GetBrandID = m_strBrandID
	End Function
	
	Public Function getA() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getA = m_varA
	End Function
	
	Public Function getB() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varB. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getB. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getB = m_varB
	End Function
	
	Public Function getC() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varC. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getC. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getC = m_varC
	End Function
	
	Public Function getD() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varD. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getD. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getD = m_varD
	End Function
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Manipulator Functions
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Sub setMethodologyID(ByRef strID As String)
		m_strMethodologyID = strID
	End Sub
	
	Public Sub setQuantificationID(ByRef strID As String)
		m_strQuantificationID = strID
	End Sub
	
	Public Sub SetBrandID(ByRef strID As String)
		m_strBrandID = strID
	End Sub
	
	Public Sub setA(ByRef varX As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varA = varX
	End Sub
	
	Public Sub setB(ByRef varX As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varB. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varB = varX
	End Sub
	
	Public Sub setC(ByRef varX As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varC. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varC = varX
	End Sub
	
	Public Sub setD(ByRef varX As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varD. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varD = varX
	End Sub
End Class