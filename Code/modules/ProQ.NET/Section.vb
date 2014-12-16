Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Section_NET.Section")> Public Class Section
	
	'+
	'Section.cls
	'
	'represents one of the sections of the quantification (initial demand,
	'demand, requirments, costs, or funding).  the idea I'm trying to
	'capture here is that this is a pure virtual base class for the
	'various sections.  each section will need to override the calc function.
	'how do you do that in vb?
	'
	'lbailey
	'27 may 2002
	'-
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	
	'Private Const SECTION_INITIAL_DEMAND = 0
	'Private Const SECTION_DEMAND = 1
	'Private Const SECTION_REQUIREMENTS = 2
	'Private Const SECTION_COSTS = 3
	'Private Const SECTION_FUNDS = 4
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
    Private m_qcType As Globals_Renamed.QuantityCategory
	Private m_strName As String 'Demand, Quantity Required, etc
	Private m_aoMethodologies() As Methodology
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'public properties
	
	'none
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	'+
	'CalcInitDemand()
	'
	'implementation of the formula used to calculate initial demand.
	'
	'lbailey
	'27 may 2002
	'-
	Private Function CalcInitDemand(ByRef aValues() As Object) As Object
		Dim dblResult As Double
		dblResult = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcInitDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcInitDemand = dblResult
	End Function
	
	
	'+
	'CalcDemand()
	'
	'implementation of the formula used to calculate demand.
	'
	'lbailey
	'27 may 2002
	'-
	Private Function CalcDemand(ByRef aValues() As Object) As Object
		Dim dblResult As Double
		dblResult = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcDemand = dblResult
	End Function
	
	
	
	'+
	'CalcRequirements()
	'
	'implementation of the formula used to calculate requirments.
	'
	'lbailey
	'27 may 2002
	'-
	Private Function CalcRequirements(ByRef aValues() As Object) As Object
		Dim dblResult As Double
		dblResult = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcRequirements. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcRequirements = dblResult
	End Function
	
	
	
	'+
	'CalcCosts()
	'
	'implementation of the formula used to calculate initial demand.
	'
	'lbailey
	'27 may 2002
	'-
	Private Function CalcCosts(ByRef aValues() As Object) As Object
		Dim dblResult As Double
		dblResult = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcCosts. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcCosts = dblResult
	End Function
	
	
	
	'+
	'CalcFunds()
	'
	'implementation of the formula used to calculate funds.
	'
	'lbailey
	'27 may 2002
	'-
	Private Function CalcFunds(ByRef aValues() As Object) As Object
		Dim dblResult As Double
		dblResult = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcFunds. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcFunds = dblResult
	End Function
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	'GetName()
	'SetName()
	'AddMethodology(meth params)
	'DeleteMethodology(name)
	
	
	'+
	'Calculate()
	'
	'No inheritance in vb.  bummer.  what we can do, is look at the
	'type of object we are, and then use the appropriate private function.
	'this isn't quite the same thing as inheriting and overriding the
	'base class, but it should look the same to the client, and
	'shouldn't be hard to implement.
	'
	'lbailey
	'27 may 2002
	'-
	Public Function Calculate(ByRef aValues() As Object) As Object
		'a holder for the results of our calculations
		Dim dblResult As Double
		'select the right formula
		Select Case m_qcType
			Case Globals_Renamed.QuantityCategory.Demand
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcInitDemand(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblResult = CalcInitDemand(aValues)
			Case Globals_Renamed.QuantityCategory.AdjustedDemand
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcDemand(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblResult = CalcDemand(aValues)
			Case Globals_Renamed.QuantityCategory.QuantityRequired
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcRequirements(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblResult = CalcRequirements(aValues)
			Case Globals_Renamed.QuantityCategory.Costs
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcRequirements(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblResult = CalcRequirements(aValues)
			Case Globals_Renamed.QuantityCategory.Funds
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcRequirements(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblResult = CalcRequirements(aValues)
			Case Else
				'what should be the default?
				dblResult = 0
		End Select
		'return the result
		'UPGRADE_WARNING: Couldn't resolve default property of object Calculate. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Calculate = dblResult
		
		
	End Function
	
	Public Sub SetSectionType(ByRef qcType As Integer)
		
		m_qcType = qcType
		
		Select Case m_qcType
			Case Globals_Renamed.QuantityCategory.Demand
				m_strName = "Demand"
			Case Globals_Renamed.QuantityCategory.AdjustedDemand
				m_strName = "Adjusted Demand"
			Case Globals_Renamed.QuantityCategory.QuantityRequired
				m_strName = "Quantity Required"
			Case Globals_Renamed.QuantityCategory.Costs
				m_strName = "Financial Requirements"
			Case Globals_Renamed.QuantityCategory.Funds
				m_strName = "Budget Reconcilliation"
		End Select
	End Sub
	
	Public Function GetSectionType() As Integer
		GetSectionType = m_qcType
	End Function
	
	Public Function GetSectionName() As String
		GetSectionName = m_strName
	End Function
End Class