Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolPatternTest_NET.ProtocolPatternTest")> Public Class ProtocolPatternTest
	'ProtocolPatternTest.cls
	'
	'the pattern test is one of the elements of a protocol patternlevel.
	'
	'lbailey
	'7 june 2002
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'connection to the db
	Private m_objDB As ProQdb.ProtocolPatternTestDB
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	
	'+
	'Create()
	'
	'we won't be creating levels until we write the admin screens.
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub Create()
		'stub
	End Sub
	
	
	
	
	'+
	'Load(byref inVal as variant)
	'
	'determines whether db object or a string is being passed in, and then
	'calls the appropriate load function to load the object
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub Load(ByRef strID As String)
		
		'instantiate the object
		m_objDB = New ProQdb.ProtocolPatternTestDB
		'load from id
		LoadByID(strID)
		
	End Sub
	
	
	
	'+
	'LoadByID()
	'
	'creates and loads the db object referenced by the given id
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub LoadByID(ByRef strID As String)
		
		m_objDB.Load(strID)
		
	End Sub
	
	
	
	'+
	'LoadFromObject()
	'
	'copies the values out of the db object into this object
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub LoadFromObject(ByRef objDB As ProQdb.ProtocolPatternTestDB)
		
		'copy the inval object
		m_objDB = objDB
		
	End Sub
	
	
	
	'+
	'Delete()
	'
	'we won't need this until we write the admin tool
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub Delete(ByRef strID As String)
		'stub
	End Sub
	
	
	
	
	Public Sub SetLevelID(ByRef strID As String)
		m_objDB.SetLevelID(strID)
	End Sub
	
	Public Sub SetName(ByRef strName As String)
		m_objDB.SetName(strName)
	End Sub
	
	
	
	Public Function GetID() As String
		GetID = m_objDB.GetID
	End Function
	
	Public Function GetLevel() As String
		GetLevel = m_objDB.GetLevelID
	End Function
	
    Public Function GetName() As String
        'GetName = m_objDB.GetName
        If m_objDB Is Nothing Then
            GetName = ""
        Else
            GetName = m_objDB.GetName
        End If
    End Function
End Class