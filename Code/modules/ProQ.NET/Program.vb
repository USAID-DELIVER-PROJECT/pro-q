Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Program_NET.Program")> Public Class Program
	
	'+
	'Program.cls
	'The Program that ProQ is being run on
	
	
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
		
		Dim objProgramDB As New ProQdb.ProgramDB
		
		With objProgramDB
			.Load(strID)
			m_strGuidID = .GetID()
			m_strName = .GetName()
			m_strNotes = .GetNotes()
		End With
		
		Load = S_OK
		
	End Function
	
	Public Function Save() As Short
		
		Dim objProgramDB As New ProQdb.ProgramDB
		
		With objProgramDB
			If m_fNewRecord = True Then
				.SetID(m_strGuidID)
				.SetName(m_strName)
				.SetNotes(m_strNotes)
				
				'Add The New Record
				.Create()
			Else
				'Update Record
				.SetID(m_strGuidID)
				.SetName(m_strName)
				.SetNotes(m_strNotes)
				
				'Save the Record
				.Update()
			End If
		End With
		
		m_fNewRecord = False
		m_fIsDirty = False
		
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
		
		' Call the DELETE Function from the DB object to get this one
		Dim objProgramDB As New ProQdb.ProgramDB
		
		If m_fNewRecord = False Then
			objProgramDB.Delete(m_strGuidID)
		End If
		
		'UPGRADE_NOTE: Object objProgramDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objProgramDB = Nothing
	End Sub
	
	'GetID()
	Public Function GetID() As String
		GetID = m_strGuidID
	End Function
	
	Public Function SetName(ByRef strValue As String) As Short
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_strName, strValue)
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
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_strNotes, strValue)
		'set the member
		m_strNotes = strValue
		'set the return code
		SetNotes = S_OK
	End Function
	
	'GetNotes()
	Public Function GetNotes() As String
		GetNotes = m_strNotes
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_fIsDirty = False
		m_fNewRecord = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	Public Function GetIsNew() As Boolean
		GetIsNew = m_fNewRecord
	End Function
	
	Public Function GetIsDirty() As Boolean
		GetIsDirty = m_fIsDirty
	End Function
	
	Public Sub SetIsDirty(ByRef fDirty As Boolean)
		m_fIsDirty = fDirty
	End Sub
End Class