Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Context_NET.Context")> Public Class Context
	'Context
	'
	'This class provides the client with information about the client's
	'current status
	'
	'lbailey
	'17 may 2002
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                           constants
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                        private properties
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Private m_objAggregation As Aggregation
	Private m_objQuantification As Quantification
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                          private methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'constructor
	'
	'by default, the class should be initialized such that it's pointing
	'to the top(head) of the script.
	'
	'lbailey
	'17 may 2002
	'
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'nothing
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	
	'destructor
	'
	'clean up
	'
	'lbailey
	'17 may 2002
	'
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'+
	'standard manipulators
	'
	'lbailey
	'10 june 2002
	'-
	Public Sub SetAggregation(ByRef objAggregation As Aggregation)
		m_objAggregation = objAggregation
	End Sub
	
	Public Sub SetQuantitification(ByRef objQuantification As Quantification)
		m_objQuantification = objQuantification
	End Sub
	
	
	
	'+
	'standard accessors
	'
	'lbailey
	'10 june 2002
	'-
	Public Function GetAggregation() As Aggregation
		GetAggregation = m_objAggregation
	End Function
	
	Public Function GetQuantification() As Quantification
		GetQuantification = m_objQuantification
	End Function
	
	Public Function GetCountryName() As String
		GetCountryName = m_objAggregation.GetCountryName
	End Function
End Class