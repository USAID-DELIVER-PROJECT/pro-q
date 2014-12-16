Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Question_NET.Question")> Public Class Question
	
	'question.cls
	
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
	Private m_objQuestionDB As New ProQdb.QuestionDB
	
	
	
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
		' Comments    : Loads the specified question
		' Parameters  : lngID - ID of the question to load
		' Returns     : None
		' Created     : 05-Jun-02 LKB
		'---------------------------------------------------------
		
		'Note that this is assuming that we would already know the question ID we want
		'by knowing the response ID we can get the question ID and then load the question
		
		
		'Load the specified entry in the table
		m_objQuestionDB.Load(lngID)
		
	End Function
	
	'-----------------------------------------------------------------
	'Standard Accessor Functions
	' Modified   : 05-Jun-02 LKB
	'-----------------------------------------------------------------
	Public Function GetID() As Integer
		GetID = m_objQuestionDB.GetID
	End Function
	Public Function GetTitle() As String
		GetTitle = m_objQuestionDB.GetName
	End Function
	Public Function GetText() As String
		GetText = m_objQuestionDB.GetText
	End Function
	Public Function GetOverride() As Boolean
		GetOverride = m_objQuestionDB.GetOverride
	End Function
	Public Function GetORType() As Integer
		GetORType = m_objQuestionDB.GetORType
	End Function
	Public Function GetEditRuleID() As Integer
		GetEditRuleID = m_objQuestionDB.GetEditRuleID
	End Function
End Class