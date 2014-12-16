Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("DemandDB_NET.DemandDB")> Public Class DemandDB
	'DemandDB Holds a DemandCalcualtion
	
	Private m_strQuantificationID As String
	Private m_strMethodologyID As String
	Private m_strBrandID As String
	Private m_varResult As Object
	Private m_varDemand2 As Object
	Private m_varValue1 As Object
	
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
	
	Public Function getBrandID() As String
		getBrandID = m_strBrandID
	End Function
	
	Public Function getResult() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varResult. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getResult. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getResult = m_varResult
	End Function
	
	Public Function getDemand2() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varDemand2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getDemand2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getDemand2 = m_varDemand2
	End Function
	Public Function getValue1() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varValue1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getValue1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getValue1 = m_varValue1
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
	
	Public Sub setBrandID(ByRef strID As String)
		m_strBrandID = strID
	End Sub
	
	Public Sub setResult(ByRef varResult As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varResult. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varResult. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varResult = varResult
	End Sub
	
	Public Sub setDemand2(ByRef varD2 As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varD2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varDemand2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varDemand2 = varD2
	End Sub
	
	Public Sub setValue1(ByRef varV1 As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varV1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varValue1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varValue1 = varV1
	End Sub
End Class