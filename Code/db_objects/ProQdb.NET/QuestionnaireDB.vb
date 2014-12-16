Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("QuestionnaireDB_NET.QuestionnaireDB")> Public Class QuestionnaireDB
	'Private Const m_strTable = "qselQuestionnaire"
	
	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    'Private m_objRS As DataSet 
	
	Private m_lngParentIndex As Integer
	Private m_lngIndex As Integer
	Private m_lngType As Integer
	Private m_lngGroupType As Integer
	Private m_strLabel As String
	Private m_strParentType As String
	Private m_strTreeType As String
	Private m_strBrandName As String
	Private m_strName As String
	Private m_strQuestion As String
	Private m_dblAnswer As Object
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'Class_Initialize() (constructor)
	'
	'connect to the db and load the data into memory
	'
	'lbailey
	'16 may 2002
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'get a connection to the db
		'Set m_objConn = GetConnection(DB_DSN)
		
		'load the data into memory
		'm_objRS.Open strTable, m_objConn, adOpenDynamic, adLockPessimistic, adCmdTable
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'+
	'Class_Terminate() (destructor)
	'
	'cleans up whatever needs to be cleaned up when this object is
	'released.
	'
	'lbailey
	'6 june 2002
	'-
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'release the connections to the db
		'm_objRS.Close
		'Set m_objConn = Nothing
		
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
	'standard manipulator functions
	'
	'jleiner
	'17 Sep 2002
	'-
	Public Sub SetIndex(ByRef lngID As Integer)
		m_lngIndex = lngID
	End Sub
	
	Public Sub SetParentIndex(ByRef lngID As Integer)
		m_lngParentIndex = lngID
	End Sub
	Public Sub SetType(ByRef lngType As Integer)
		m_lngType = lngType
	End Sub
	Public Sub SetGroupType(ByRef lngType As Integer)
		m_lngGroupType = lngType
	End Sub
	
	Public Sub SetLabel(ByRef strLabel As String)
		m_strLabel = strLabel
	End Sub
	
	Public Sub SetParentType(ByRef strParentType As String)
		m_strParentType = strParentType
	End Sub
	
	Public Sub SetTreeType(ByRef strTreeType As String)
		m_strTreeType = strTreeType
	End Sub
	
	Public Sub SetBrandName(ByRef strBrandName As String)
		m_strBrandName = strBrandName
	End Sub
	
	Public Sub SetName(ByRef strName As String)
		m_strName = strName
	End Sub
	Public Sub SetQuestion(ByRef strQuestion As String)
		m_strQuestion = strQuestion
	End Sub
	Public Sub SetAnswer(ByRef varAnswer As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object varAnswer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_dblAnswer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_dblAnswer = varAnswer
	End Sub
	
	'+
	'standard accessor functions
	'
	'jleiner
	'17 Sep 2002
	'-
	Public Function GetIndex() As Integer
		GetIndex = m_lngIndex
	End Function
	
	Public Function GetParentIndex() As Integer
		GetParentIndex = m_lngParentIndex
	End Function
	
	'UPGRADE_NOTE: GetType was upgraded to GetType_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetType_Renamed() As Integer
		GetType_Renamed = m_lngType
	End Function
	
	Public Function GetGroupType() As Integer
		GetGroupType = m_lngGroupType
	End Function
	
	Public Function GetLabel() As String
		GetLabel = m_strLabel
	End Function
	
	Public Function GetBrandName() As String
		GetBrandName = m_strBrandName
	End Function
	Public Function GetParentType() As String
		GetParentType = m_strParentType
	End Function
	
	Public Function getTreeType() As String
		getTreeType = m_strTreeType
	End Function
	
	Public Function GetName() As String
		GetName = m_strName
	End Function
	
	Public Function GetQuestion() As String
		GetQuestion = m_strQuestion
	End Function
	
	Public Function GetAnswer() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_dblAnswer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetAnswer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetAnswer = m_dblAnswer
	End Function
End Class