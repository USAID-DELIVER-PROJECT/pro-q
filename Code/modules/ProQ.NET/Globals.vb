Option Strict Off
Option Explicit On
'UPGRADE_NOTE: Globals was upgraded to Globals_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
Module Globals_Renamed
	'Globals.bas
	'
	'things that should be visible to every class in the project.
	'
	'lbailey
	'1 june 2002
	'
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            enumerations
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'used as categories for the collections of selected brands in
	'the brand quantities
	Public Enum QuantityCategory
		None = -1
		MethodologyLogistics = 0
		Demand = 1
		AdjustedDemand = 2
		QuantityRequired = 3
		Costs = 4
		CostsKits = 5
		CostsCustoms = 6
		CostsStorage = 7
		Funds = 8
		FundsKits = 9
		FundsCustoms = 10
		FundsStorage = 11
		Protocol = 12
	End Enum
	
	
	'NB: this is going away in favor of the brand object
	'lb - 8 june 2002
	'
	'a struct in which to keep counts of brands
	Public Structure BrandInfo
		Dim objBrand As Brand
		Dim lCount As Integer
		Dim sngPercent As Single
	End Structure
	
	'NB: replaced by enum QuantityCategory
	'lb - 8 june 2002
	'
	'used by sections and quantities
	Private Enum SectionType
		Protocol = 0
		Demand = 1
		AdjustedDemand = 2
		Required = 3
		Costs = 4
		Funds = 5
	End Enum
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'global variables
	'
	'global variables follow the "g_strFoo" form.
	Public g_objGUID As New jsiGuidGenerator.GUIDGen
	Public g_objEditor As New Editor
	Public g_objDM As New DemandCalculator
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'global functions and procedures
	'
	'this might be a good place for a registry-reading routine.  i found
	'the "settings" functions.  these allow us to get/set the registry.
	'I didn't know these things existed.
	
	
	'+
	'NulltoValue()
	'
	'If an expression is null replace it with the passed in value, else
	'return the expression
	'
	'jleiner
	'27 june 2002
	'-
	Public Function NulltoValue(ByRef varValue1 As Object, ByRef varValue2 As Object) As Object
		
		Dim varReturn As Object
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        If IsDBNull(varValue1) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object varValue2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object varReturn. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varReturn = varValue2
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object varValue1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object varReturn. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varReturn = varValue1
        End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object varReturn. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object NulltoValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NulltoValue = varReturn
		
	End Function
End Module