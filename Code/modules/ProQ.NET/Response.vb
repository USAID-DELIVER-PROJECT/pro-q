Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Response_NET.Response")> Public Class Response
	
	'response.cls
	'
	'the response object represents the responses to the questions. though
	'we have determined that the responses can easily live in the same table
	'that holds the data which represents the script (which is also the same
	'data that the treeview uses), the response is physically a different
	'"thing", and has different needs, and so therefore is a different
	'object, regardless of where it lives in the db.
	'
	'right.  so, who uses the response?  does it need to know the section?
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Const S_OK As Short = 0
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private m_objResponseDB As New ProQdb.ResponseDB
	Private m_objScriptRelationshipDB As New ProQdb.ScriptRelationshipDB
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'public properties
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'(none)
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function Load(ByRef lngID As Integer) As Object
		' Comments    : Loads the specified response record
		' Parameters  : lngID - ID of the response to load
		' Returns     : None
		' Created     : 05-Jun-02 LKB
		'---------------------------------------------------------
		
		
		'Load the specified entry in the table
		m_objResponseDB.Load(lngID)
		
		'get the associated questionID from the ScriptRelationshipDB
		m_objScriptRelationshipDB.Load(GetScriptRelationshipID)
		
	End Function
	Public Sub Update()
		m_objResponseDB.Update()
	End Sub
	
	'-----------------------------------------------------------------
	'Standard Accessor Functions
	' Modified   : 05-Jun-02 LKB
	'-----------------------------------------------------------------
	Public Function GetID() As Integer
		GetID = CInt(m_objResponseDB.GetID)
	End Function
	Public Function GetTreeviewID() As Integer
		GetTreeviewID = m_objResponseDB.GetTreeviewID
	End Function
	Public Function GetScriptRelationshipID() As Integer
		GetScriptRelationshipID = m_objResponseDB.GetScriptRelationshipID
	End Function
	Public Function GetAnswer() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_objResponseDB.GetAnswer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetAnswer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetAnswer = m_objResponseDB.GetAnswer
	End Function
	Public Function GetStatus() As Integer
		GetStatus = m_objResponseDB.GetStatus
	End Function
	Public Function GetTypeID() As String
		GetTypeID = m_objResponseDB.GetTypeID
	End Function
	Public Function GetEnabled() As Boolean
		GetEnabled = m_objResponseDB.GetEnabled
	End Function
	Public Function GetNotes() As String
		GetNotes = m_objResponseDB.GetNotes
	End Function
	Public Function GetQuestionID() As Integer
		GetQuestionID = m_objScriptRelationshipDB.GetQuestionID
	End Function
	Public Function GetFalseGroup() As Integer
		GetFalseGroup = m_objScriptRelationshipDB.GetFalseGroup
	End Function
	Public Function GetTrueGroup() As Integer
		GetTrueGroup = m_objScriptRelationshipDB.GetTrueGroup
	End Function
	Public Function GetGroupParent() As Integer
		GetGroupParent = CInt(m_objScriptRelationshipDB.GetGroupParent)
	End Function
	Public Function GetMasterParent() As Integer
		GetMasterParent = CInt(m_objScriptRelationshipDB.GetMasterParent)
	End Function
	Public Function GetMasterGrandParent() As Integer
		GetMasterGrandParent = CInt(m_objScriptRelationshipDB.GetMasterGrandParent)
	End Function
	'-----------------------------------------------------------------
	'Standard Manipulator Functions
	' Modified   : 05-Jun-02 LKB
	'-----------------------------------------------------------------
	Public Function SetAnswer(ByRef varAnswer As Object) As Object
		m_objResponseDB.SetAnswer(varAnswer)
	End Function
	Public Function SetStatus(ByRef lngStatus As Integer) As Object
		m_objResponseDB.SetStatus(lngStatus)
	End Function
	Public Function SetEnabled(ByRef fEnabled As Boolean) As Object
		m_objResponseDB.SetEnabled(fEnabled)
	End Function
	Public Function SetNotes(ByRef strValue As String) As Short
		m_objResponseDB.SetNotes(strValue)
	End Function
End Class