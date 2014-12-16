Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("BrandQtyCalculatorCol_NET.BrandQtyCalculatorCol")> Public Class BrandQtyCalculatorCol
	Implements System.Collections.IEnumerable
	
	'local variable to hold collection
	Private m_objBQC As BrandQtyCalculator
	Private mCol As Collection
	
	Private m_varStorageCapacity As Object
	Private m_varColdStorage As Object
	
	Public Function Add(Optional ByRef sKey As String = "") As BrandQtyCalculator
		'create a new object
		Dim objNewMember As BrandQtyCalculator
		objNewMember = New BrandQtyCalculator
		
		'set the properties passed into the method
		If Len(sKey) = 0 Then
			mCol.Add(objNewMember)
		Else
			mCol.Add(objNewMember, sKey)
		End If
		
		
		'return the object created
		Add = objNewMember
		'UPGRADE_NOTE: Object objNewMember may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objNewMember = Nothing
		
		
	End Function
	
	Default Public ReadOnly Property Item(ByVal vntIndexKey As Object) As BrandQtyCalculator
		Get
			'used when referencing an element in the collection
			'vntIndexKey contains either the Index or Key to the collection,
			'this is why it is declared as a Variant
			'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
			Item = mCol.Item(vntIndexKey)
		End Get
	End Property
	
	
	Public ReadOnly Property Count() As Integer
		Get
			'used when retrieving the number of elements in the
			'collection. Syntax: Debug.Print x.Count
			Count = mCol.Count()
		End Get
	End Property
	
	'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
	'Public ReadOnly Property NewEnum() As stdole.IUnknown
		'Get
			'this property allows you to enumerate
			'this collection with the For...Each syntax
			'NewEnum = mCol._NewEnum
		'End Get
	'End Property
	
	Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: Uncomment and change the following line to return the collection enumerator. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"'
        GetEnumerator = mCol.GetEnumerator
	End Function
	
	
	Public Sub Remove(ByRef vntIndexKey As Object)
		'used when removing an element from the collection
		'vntIndexKey contains either the Index or Key, which is why
		'it is declared as a Variant
		'Syntax: x.Remove(xyz)
		
		
		mCol.Remove(vntIndexKey)
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'creates the collection when this class is created
		mCol = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'destroys collection when this class is terminated
		'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mCol = Nothing
		'UPGRADE_NOTE: Object m_objBQC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objBQC = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function GetBAdjustedVolumeForAll() As Object
		' Comments  :  Returns the total volume for all BrandQtyCalcs
		' Parameters:  -
		' Returns   :  Variant - Total of volumes
		' Created   :  01-July-2002 jleiner
		'----------------------------------------------------------------------
		
		Dim varVolume As Object
		
		For	Each m_objBQC In mCol
			'UPGRADE_WARNING: Couldn't resolve default property of object varVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varVolume = varVolume + m_objBQC.getBdjustedVolume
		Next m_objBQC
		
		'UPGRADE_WARNING: Couldn't resolve default property of object varVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetBAdjustedVolumeForAll. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetBAdjustedVolumeForAll = varVolume
		
	End Function
	
	Public Function GetBAdjustedVolumePerShipment() As Object
		' Comments  :  Returns the Total Volume for the sum of the shipments
		' Parameters:  -
		' Returns   :  Variant - Total of volumes
		' Created   :  01-July-2002 jleiner
		'----------------------------------------------------------------------
		
		Dim varVolume As Object
		Dim varShipments As Object
		
		For	Each m_objBQC In mCol
			'UPGRADE_WARNING: Couldn't resolve default property of object varVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varVolume = varVolume + m_objBQC.getBdjustedVolume
			'UPGRADE_WARNING: Couldn't resolve default property of object varShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varShipments = varShipments + m_objBQC.GetShipments
		Next m_objBQC
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NulltoValue(varShipments, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If NulltoValue(varShipments, 0) = 0 Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetBAdjustedVolumePerShipment. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetBAdjustedVolumePerShipment = System.DBNull.Value
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object varShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetBAdjustedVolumePerShipment = System.Math.Round(varVolume / varShipments, 2)
		End If
	End Function
	
	Public Sub SetStorageCapacity(ByRef varCapacity As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varStorageCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varStorageCapacity = varCapacity
	End Sub
	
	Public Sub SetColdStorage(ByRef varColdStorage As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varColdStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varColdStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varColdStorage = varColdStorage
	End Sub
	
	Public Function GetStorageCapacity() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varStorageCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetStorageCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetStorageCapacity = m_varStorageCapacity
	End Function
	
	Public Function GetColdStorage() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varColdStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetColdStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetColdStorage = m_varColdStorage
	End Function
 
	Public Function GetVolumeByType(ByRef fIsColdStorage As Boolean) As Object
		
		Dim varVolume As Object
		
		For	Each m_objBQC In mCol
			If m_objBQC.GetSelectedBrand.GetBrand.GetColdStorage = fIsColdStorage Then
				'UPGRADE_WARNING: Couldn't resolve default property of object varVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				varVolume = varVolume + m_objBQC.getBdjustedVolume
			End If
		Next m_objBQC
		
		'UPGRADE_WARNING: Couldn't resolve default property of object varVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetVolumeByType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetVolumeByType = varVolume
		
	End Function
	
	Public Function CompareShipmentsToStorage(ByRef fIsColdStorage As Boolean) As Object
		' Comments  :  Compare the Capacity for Storage against the expected shipment size
		' Parameters:  -
		' Returns   :  Variant = remaing space, if negative shipment is larger
		' Created   :  01-July-2002 jleiner
		'----------------------------------------------------------------------
		
		Dim varVolume As Object
		Dim varShipments As Object
		Dim varCapacity As Object
		
		For	Each m_objBQC In mCol
			If m_objBQC.GetSelectedBrand.GetBrand.GetColdStorage = fIsColdStorage Then
				'UPGRADE_WARNING: Couldn't resolve default property of object varVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				varVolume = varVolume + m_objBQC.getBdjustedVolume
				'UPGRADE_WARNING: Couldn't resolve default property of object varShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				varShipments = varShipments + m_objBQC.GetShipments
			End If
		Next m_objBQC
		
		If fIsColdStorage = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_varColdStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varCapacity = m_varColdStorage
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object m_varStorageCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varCapacity = m_varStorageCapacity
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NulltoValue(varShipments, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If NulltoValue(varShipments, 0) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object varCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CompareShipmentsToStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CompareShipmentsToStorage = varCapacity
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object varShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CompareShipmentsToStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CompareShipmentsToStorage = varCapacity - System.Math.Round(varVolume / varShipments, 2)
		End If
		
	End Function
End Class