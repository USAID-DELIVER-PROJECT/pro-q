Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("LogisticsConsDB_NET.LogisticsConsDB")> Public Class LogisticsConsDB
	'LogisiticsDB Holds a The Logisitcs Considerations Data
	
	Private m_strQuantificationID As String
	Private m_strBrandID As String
	Private m_varLeadTime As Object
	Private m_varBufferStock As Object
	Private m_varShipments As Object
	Private m_varStorageCapacity As Object
	Private m_varColdStorage As Object
	Private m_varQtyOnHand As Object
	Private m_varQtyOnOrder As Object
	
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
	Public Function getQuantificationID() As String
		getQuantificationID = m_strQuantificationID
	End Function
	
	Public Function getBrandID() As String
		getBrandID = m_strBrandID
	End Function
	
	Public Function getLeadTime() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varLeadTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getLeadTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getLeadTime = m_varLeadTime
	End Function
	
	Public Function getBufferStock() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varBufferStock. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getBufferStock. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getBufferStock = m_varBufferStock
	End Function
	Public Function getShipments() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getShipments = m_varShipments
	End Function
	Public Function getStorageCapacity() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varStorageCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getStorageCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getStorageCapacity = m_varStorageCapacity
	End Function
	Public Function getColdStorage() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varColdStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getColdStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getColdStorage = m_varColdStorage
	End Function
	Public Function getQtyOnHand() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQtyOnHand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getQtyOnHand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getQtyOnHand = m_varQtyOnHand
	End Function
	Public Function getQtyOnOrder() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQtyOnOrder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getQtyOnOrder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getQtyOnOrder = m_varQtyOnOrder
	End Function
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Manipulator Functions
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Public Sub setQuantificationID(ByRef strID As String)
		m_strQuantificationID = strID
	End Sub
	
	Public Sub setBrandID(ByRef strID As String)
		m_strBrandID = strID
	End Sub
	
	Public Sub setLeadTime(ByRef varLeadTime As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varLeadTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varLeadTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varLeadTime = varLeadTime
	End Sub
	
	Public Sub setBufferStock(ByRef varBufferStock As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varBufferStock. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varBufferStock. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varBufferStock = varBufferStock
	End Sub
	
	Public Sub setShipments(ByRef varShipments As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varShipments = varShipments
	End Sub
	
	Public Sub setStorageCapacity(ByRef varStorageCapacity As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varStorageCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varStorageCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varStorageCapacity = varStorageCapacity
	End Sub
	
	Public Sub setColdStorage(ByRef varColdStorage As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varColdStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varColdStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varColdStorage = varColdStorage
	End Sub
	Public Sub setQtyOnHand(ByRef varQtyOnHand As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varQtyOnHand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQtyOnHand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varQtyOnHand = varQtyOnHand
	End Sub
	
	Public Sub setQtyOnOrder(ByRef varQtyOnOrder As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varQtyOnOrder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQtyOnOrder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varQtyOnOrder = varQtyOnOrder
	End Sub
End Class