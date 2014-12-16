Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("AnswerDB_NET.AnswerDB")> Public Class AnswerDB
	'Answer.cls
	'
	'this class represents a row in the collection of Answers
	'
	'lbailey
	'12 june 2002
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Private m_strAnswer As String
	Private m_strLabel As String
	Private m_lParentIndex As Integer
	Private m_strActionID As String
	Private m_fNotes As Boolean
	Private m_strComments As String
	Private m_fPercent As Boolean
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'-----------------------------------------------------------------
	'Standard Accessor Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	Public Function GetAnswer() As String
		GetAnswer = m_strAnswer
	End Function
	
	Public Function GetLabel() As String
		GetLabel = m_strLabel
	End Function
	
	Public Function GetParentIndex() As Integer
		GetParentIndex = m_lParentIndex
	End Function
	
	Public Function GetActionID() As String
		GetActionID = m_strActionID
	End Function
	Public Function GetNotes() As String
		GetNotes = CStr(m_fNotes)
	End Function
	
	Public Function GetComment() As String
		GetComment = m_strComments
	End Function
	Public Function GetIsPercent() As Integer
		GetIsPercent = m_fPercent
	End Function
	
	'-----------------------------------------------------------------
	'Standard manipulator Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	Public Sub SetAnswer(ByRef strAnswer As String)
		m_strAnswer = strAnswer
	End Sub
	
	Public Sub SetLabel(ByRef strLabel As String)
		m_strLabel = strLabel
	End Sub
	
    Public Sub SetParentIndex(ByRef lParentIndex As Integer)
        m_lParentIndex = lParentIndex
    End Sub
	
	Public Sub SetActionID(ByRef strActionID As String)
		m_strActionID = strActionID
	End Sub
	Public Sub SetNotes(ByRef fNotes As String)
		m_fNotes = CBool(fNotes)
	End Sub
	
	Public Sub SetComment(ByRef strNotes As String)
		m_strComments = strNotes
	End Sub
    Public Sub SetIsPercent(ByRef fPercent As Boolean)
        m_fPercent = fPercent
    End Sub
End Class