Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("BrandQtyCalculator_NET.BrandQtyCalculator")> Public Class BrandQtyCalculator
	' Brand Quantity Required Stuff
	' Stores all numbers from demand up through the process
	'
	' 24-June-2002  jleiner
	'---------------------------------------------------------------------
	
	' Storage Variables
	Private m_varDemand_Initial As Object
	Private m_varDemand_Level2 As Object
	Private m_varServiceCapacity As Object
	'Private m_varDemand_SCFiltered As Variant
	Private m_varWastage As Object
	Private m_varQC As Object
	'Private m_varDemand_Adjusted As Variant
	Private m_varLeadTime As Object
	Private m_varBufferStock As Object
	Private m_varShipments As Object
	Private m_varStorageCapacity As Object
	Private m_varColdStorage As Object
	Private m_varQtyonHand As Object
	Private m_varQtyonOrder As Object
	
	'Brand Related
	Private m_objBrand As SelectedBrand
	
	Private m_strBrandID As String
	Private m_strKitID As String
	Private m_lngTestsPerKit As Integer
	Private m_dblKitVolume As Double
	Private m_dblCostPerTest As Double
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Calculated Functions
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Public Function GetDemand_SCFiltered() As Object
		
		Dim varDemandSC As Object
		
		'Return Lessor of the 2 values
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(m_varDemand_Initial) Or IsDbNull(m_varServiceCapacity) Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varDemandSC. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varDemandSC = System.DBNull.Value
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object m_varServiceCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object m_varDemand_Initial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varDemandSC. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varDemandSC = IIf(m_varDemand_Initial < m_varServiceCapacity, m_varDemand_Initial, m_varServiceCapacity)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object varDemandSC. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetDemand_SCFiltered = System.Math.Round(NulltoValue(varDemandSC, 0), 0)
		
	End Function
	Public Function GetDemand_adjusted() As Object
        GetDemand_adjusted = System.Math.Round(GetDemand_SCFiltered() * (1 + NulltoValue(GetWastage(), 0) + NulltoValue(GetQC(), 0)), 0)
	End Function
	
	Public Function GetAMAD() As Object
		'Average Monthly Adjusted Demand
		'UPGRADE_WARNING: Couldn't resolve default property of object GetAMAD. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetAMAD = GetDemand_adjusted / 12
	End Function
	
	Public Function GetLeadTimeStock() As Object
		' Lead Time Stock Calculation
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varLeadTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetLeadTimeStock = System.Math.Round(NulltoValue(m_varLeadTime, 0) * GetAMAD(), 0)
	End Function
	Public Function GetBufferStockUnits() As Object
		'Buffer Stock calculation to units
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varBufferStock. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetBufferStockUnits = System.Math.Round(NulltoValue(m_varBufferStock, 0) * GetAMAD(), 0)
	End Function
	Public Function getBdjustedDemand() As Object
		'Buffer Stock Adjusted Demand
		'UPGRADE_WARNING: Couldn't resolve default property of object GetLeadTimeStock(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getBdjustedDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getBdjustedDemand = GetDemand_adjusted + GetLeadTimeStock() + GetBufferStockUnits
	End Function
	
	Public Function getBdjustedVolume() As Object
		'Volume of the BufferStock Adjusted Demand
		Dim varBVolume As Object
		
		With m_objBrand.GetKit
			If .GetTestsPerKit = 0 Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object varBVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				varBVolume = System.DBNull.Value
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object varBVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				varBVolume = System.Math.Round((getBdjustedDemand / .GetTestsPerKit) + 0.5, 0) * .GetVolume
			End If
		End With
		
		'UPGRADE_WARNING: Couldn't resolve default property of object varBVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getBdjustedVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getBdjustedVolume = varBVolume
		
	End Function
	
	Public Function getQuantityRequired() As Object
		' Get Quantity Required
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQtyonOrder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQtyonHand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getBdjustedDemand(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object getQuantityRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        getQuantityRequired = getBdjustedDemand() - (NulltoValue(m_varQtyonHand, 0) + NulltoValue(m_varQtyonOrder, 0))
	End Function
	Public Function getKitsRequired() As Object
		'Get the kit Required to fill the test
		Dim varTestsPerKit As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object varTestsPerKit. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		varTestsPerKit = m_objBrand.GetTestsPerKit
		If m_objBrand.GetKit.GetTestsPerKit = 0 Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object getKitsRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			getKitsRequired = System.DBNull.Value
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object varTestsPerKit. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			getKitsRequired = System.Math.Round(getQuantityRequired / varTestsPerKit + 0.5, 0)
			'UPGRADE_WARNING: Couldn't resolve default property of object getKitsRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If getKitsRequired < 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object getKitsRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				getKitsRequired = 0
			End If
		End If
	End Function
	Public Function getKitTotalCost() As Object
		
		Dim varKitsRequired As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object getKitsRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object varKitsRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		varKitsRequired = getKitsRequired
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(varKitsRequired) Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object getKitTotalCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			getKitTotalCost = System.DBNull.Value
			'UPGRADE_WARNING: Couldn't resolve default property of object varKitsRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf varKitsRequired < 0 Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object getKitTotalCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			getKitTotalCost = 0
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object varKitsRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object getKitTotalCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			getKitTotalCost = varKitsRequired * m_objBrand.GetKitCost
		End If
	End Function
	
	Public Function getCostofQtyRequired() As Object
		'Cost of Qty Required
		Dim varQtyRequired As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object getQuantityRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object varQtyRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		varQtyRequired = getQuantityRequired
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(varQtyRequired) Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object getCostofQtyRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			getCostofQtyRequired = System.DBNull.Value
			'UPGRADE_WARNING: Couldn't resolve default property of object varQtyRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf varQtyRequired < 0 Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object getCostofQtyRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			getCostofQtyRequired = 0
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object varQtyRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object getCostofQtyRequired. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			getCostofQtyRequired = varQtyRequired * m_dblCostPerTest
		End If
	End Function
	Function GetShipmentVolume() As Object
		' Returns the Shipment Volume in Meters
		
		Dim varShipVolume As Object
		
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(m_varShipments) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetShipmentVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetShipmentVolume = 0
		Else
			With m_objBrand.GetKit
				If .GetTestsPerKit = 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object GetShipmentVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetShipmentVolume = 0
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object m_varShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object getQuantityRequired(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object varShipVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varShipVolume = (System.Math.Round((getQuantityRequired() / .GetTestsPerKit) + 0.5, 0) * (.GetVolume / 1000000)) / m_varShipments
					
					'UPGRADE_WARNING: Couldn't resolve default property of object varShipVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If varShipVolume < 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object GetShipmentVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetShipmentVolume = 0
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object varShipVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetShipmentVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetShipmentVolume = varShipVolume
					End If
				End If
			End With
		End If
		
	End Function
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Accessors
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function GetDemand_Initial() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varDemand_Initial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetDemand_Initial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetDemand_Initial = m_varDemand_Initial
	End Function
	
	Public Function GetDemand_Level2() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varDemand_Level2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetDemand_Level2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetDemand_Level2 = m_varDemand_Level2
	End Function
	
	Public Function GetServiceCapacity() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varServiceCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetServiceCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetServiceCapacity = m_varServiceCapacity
	End Function
	
	'Public Function GetDemand_SCFiltered() As Variant
	'    GetDemand_SCFiltered = m_varDemand_SCFiltered
	'End Function
	
	Public Function GetWastage() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varWastage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetWastage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetWastage = m_varWastage
	End Function
	
	
	Public Function GetQC() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQC. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetQC. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetQC = m_varQC
	End Function
	
	Public Function GetLeadTime() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varLeadTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetLeadTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetLeadTime = m_varLeadTime
	End Function
	
	Public Function GetBufferStock() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varBufferStock. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetBufferStock. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetBufferStock = m_varBufferStock
	End Function
	
	Public Function GetShipments() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetShipments = m_varShipments
	End Function
	
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
	
	Public Function GetQtyOnHand() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQtyonHand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetQtyOnHand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetQtyOnHand = m_varQtyonHand
	End Function
	
	Public Function GetQtyOnOrder() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQtyonOrder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetQtyOnOrder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetQtyOnOrder = m_varQtyonOrder
	End Function
	
	Public Function GetSelectedBrand() As SelectedBrand
		GetSelectedBrand = m_objBrand
	End Function
	
	Public Function GetKitVolume() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object GetKitVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetKitVolume = m_dblKitVolume
	End Function
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Manipulator Subs
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Sub SetDemand_Initial(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varDemand_Initial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varDemand_Initial = varValue
	End Sub
	
	Public Sub SetDemand_Level2(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varDemand_Level2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varDemand_Level2 = varValue
	End Sub
	
	Public Sub SetServiceCapacity(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varServiceCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varServiceCapacity = varValue
	End Sub
	
	Public Sub SetWastage(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varWastage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varWastage = varValue
	End Sub
	
	Public Sub SetQC(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQC. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varQC = varValue
	End Sub
	
	Public Sub SetLeadTime(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varLeadTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varLeadTime = varValue
	End Sub
	
	Public Sub SetBufferStock(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varBufferStock. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varBufferStock = varValue
	End Sub
	
	Public Sub SetShipments(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varShipments. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varShipments = varValue
	End Sub
	
	Public Sub SetStorageCapacity(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varStorageCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varStorageCapacity = varValue
	End Sub
	
	Public Sub SetColdStorage(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varColdStorage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varColdStorage = varValue
	End Sub
	
	Public Sub SetQtyOnHand(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQtyonHand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varQtyonHand = varValue
	End Sub
	
	Public Sub SetQtyOnOrder(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varQtyonOrder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_varQtyonOrder = varValue
	End Sub
	
	Public Sub SetSelectedBrand(ByRef objSB As SelectedBrand)
		m_objBrand = objSB
	End Sub
	
	'Public Sub SetTestsPerKit(varValue As Variant)
	'     m_lngTestsPerKit = varValue
	'End Sub
	
	Public Sub SetKitVolume(ByRef varValue As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_dblKitVolume = varValue
	End Sub
End Class