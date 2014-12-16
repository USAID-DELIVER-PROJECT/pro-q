Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Use_NET.Use")> Public Class Use
	
	'+
	'Use.cls
	'The Use that Quantification is for.
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	Private m_strGuidID As String
	Private m_strName As String
	Private m_strNotes As String
	
	Private m_fNewRecord As Boolean
	
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
		
		Dim objUseDB As New ProQdb.UseDB
		
		With objUseDB
			.Load(strID)
			m_strGuidID = .GetID()
			m_strName = .GetName()
			m_strNotes = .GetNotes()
		End With
		
		Load = S_OK
		
	End Function
	
	Public Function Save() As Short
		
		Dim objUseDB As New ProQdb.UseDB
		
		With objUseDB
			If m_fNewRecord = True Then
				.SetID(m_strGuidID)
				.SetName(m_strName)
				.SetNotes(m_strNotes)
				
				'Add The New Record
				.Create()
			Else
				'Update Record
				.Load(m_strGuidID)
				.SetID(m_strGuidID)
				.SetName(m_strName)
				.SetNotes(m_strNotes)
				
				'Save the Record
				.Update()
			End If
		End With
		
		m_fNewRecord = False
		
		Save = S_OK
		
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
		Dim objUseDB As New ProQdb.UseDB
		objUseDB.Delete(m_strGuidID)
		
	End Sub
	
	'GetID()
	Public Function GetID() As String
		GetID = m_strGuidID
	End Function
	
	
	Public Function SetName(ByRef strValue As String) As Short
		'set the member
		m_strName = strValue
		'set the return code
		SetName = S_OK
	End Function
	
	'GetName()
	Public Function GetName() As String
		GetName = m_strName
	End Function
	' SetNotes
	Public Function SetNotes(ByRef strValue As String) As Short
		'set the member
		m_strNotes = strValue
		'set the return code
		SetNotes = S_OK
	End Function
	
	'GetNotes()
	Public Function GetNotes() As String
		GetNotes = m_strNotes
	End Function
End Class