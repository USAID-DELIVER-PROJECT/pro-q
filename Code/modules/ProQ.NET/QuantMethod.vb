Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("QuantMethod_NET.QuantMethod")> Public Class QuantMethod
	
	'+
	'QuantMethod.cls
	'A Quantification/Methodology Record for the selected methodology
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	Private m_strGuidID As String
	Private m_strQuantificationID As String
	Private m_strMethodologyID As String
	Private m_fSelectedMethod As Boolean
	
	
	Public m_objMethodology As New Methodology
	
	Private m_varDemand As Object
	Private m_varDemand2 As Object
	
	Private m_fNewRecord As Boolean
	Private m_fIsDirty As Boolean
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'public properties
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	Public Function Load(ByRef strID As String) As Short
		
		Dim objMethodologyDB As New ProQdb.QuantMethodDB
		
		With objMethodologyDB
			.Load(strID)
			m_strGuidID = .GetID()
			m_strQuantificationID = .GetQuantificationID()
			m_strMethodologyID = .GetMethodologyID()
			m_fSelectedMethod = .GetSelectedMethod()
		End With
		
		m_objMethodology.Load(m_strMethodologyID)
		
		m_fIsDirty = False
		m_fNewRecord = False
		
		Load = S_OK
		
	End Function
	
	Public Function Save() As Short
		
		Dim objMethodologyDB As New ProQdb.QuantMethodDB
		
		With objMethodologyDB
			If m_fNewRecord = True Then
				.SetID(m_strGuidID)
				.SetQuantificationID(m_strQuantificationID)
				.SetMethodologyID(m_strMethodologyID)
				.SetSelectedMethod(m_fSelectedMethod)
				'Add The New Record
				.Create()
			Else
				'Update Record
				.Load(m_strGuidID)
				.SetID(m_strGuidID)
				.SetQuantificationID(m_strQuantificationID)
				.SetMethodologyID(m_strMethodologyID)
				.SetSelectedMethod(m_fSelectedMethod)
				'Save the Record
				.Update()
			End If
		End With
		
		m_fNewRecord = False
		m_fIsDirty = False
		
		Save = S_OK
		
		'UPGRADE_NOTE: Object objMethodologyDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objMethodologyDB = Nothing
		
	End Function
	
	
	Public Function SetID(ByRef strValue As String) As Short
		'set the member
		m_strGuidID = strValue
		'set the return code
		SetID = S_OK
	End Function
	
	Public Sub Create()
		'Create the and call the guidgen to create.
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUID.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_strGuidID = g_objGUID.GetGUID()
		'm_dtmCreated = Now()
		m_fNewRecord = True
		
	End Sub
	Public Sub Delete()
		' Call the DELETE Function from the JSIHIVDB object to get this one
		Dim objMethodologyDB As New ProQdb.QuantMethodDB
		objMethodologyDB.Delete(m_strGuidID)
		
	End Sub
	
	'GetID()
	Public Function GetID() As String
		GetID = m_strGuidID
	End Function
	
	Public Function SetQuantificationID(ByRef strValue As String) As Short
		'set the member
		m_strQuantificationID = strValue
		'set the return code
		SetQuantificationID = S_OK
	End Function
	
	'GetQuantificationID()
	Public Function GetQuantificationID() As String
		GetQuantificationID = m_strQuantificationID
	End Function
	' SetMethodologyID
	Public Function SetMethodologyID(ByRef strValue As String) As Short
		'set the member
		m_strMethodologyID = strValue
		m_objMethodology.Load(strValue)
		'set the return code
		SetMethodologyID = S_OK
	End Function
	'GetMethodologyID()
	Public Function GetMethodologyID() As String
		GetMethodologyID = m_strMethodologyID
	End Function
	
	' SetSelectedMethod
	Public Function SetSelectedMethod(ByRef fValue As Boolean) As Short
		'set the member
		m_fSelectedMethod = fValue
		'set the return code
		SetSelectedMethod = S_OK
	End Function
	'GetSelectedMethod()
	Public Function GetSelectedMethod() As Boolean
		GetSelectedMethod = m_fSelectedMethod
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		m_fNewRecord = False
		m_fIsDirty = False
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object m_objMethodology may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objMethodology = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class