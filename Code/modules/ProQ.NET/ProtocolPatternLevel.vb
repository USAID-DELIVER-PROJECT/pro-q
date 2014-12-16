Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolPatternLevel_NET.ProtocolPatternLevel")> Public Class ProtocolPatternLevel
	'ProtocolPatternLevel.cls
	'
	'the protocol holds both all of the lookup information, and all of
	'the user-input info regarding the number of samples to be tested,
	'and the numbers and types of tests that will be needed to accomplish
	'this.
	'
	'lbailey
	'7 june 2002
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'connection to the db
	Private m_objDB As ProQdb.ProtocolPatternLevelDB
	
	'collection of tests associated with this protocol
	Private m_cTests As New Collection
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'public properties
	
	'none
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	
	
	'+
	'LoadTests()
	'
	'loads the pattern tests under the patternlevel
	'
	'lbailey
	'1 june 2002
	'-
	Private Sub LoadTests()
		
		'setup the parameter array for the sp
		Dim aParms(1, 1) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object aParms(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParms(0, 0) = New Guid(m_objDB.GetID)
		'UPGRADE_WARNING: Couldn't resolve default property of object aParms(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParms(0, 1) = DbType.Guid
		
		'instantiate the dbcoll object and pass it the parm array
		Dim cDBTests As Collection
		Dim objPatternTestDBColl As New ProQdb.ProtocolPatternTestDBCollection
        Dim objDBTest As ProQdb.ProtocolPatternTestDB
        cDBTests = objPatternTestDBColl.GetCollection(SP_GET_PATTERNTEST_BY_LEVEL, aParms)

        For Each objDBTest In cDBTests
            Dim objTest As New ProtocolPatternTest
            objTest.LoadFromObject(objDBTest)
            m_cTests.Add(objTest)
            'UPGRADE_NOTE: Object objTest may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            objTest = Nothing
        Next objDBTest
	End Sub
	
	
	
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
		'load from id
		LoadByID(strID)
		
	End Sub
	
	
	
	'+
	'LoadByID()
	'
	'creates and loads the db object referenced by the given id, and then
	'constructs child objects
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub LoadByID(ByRef strID As String)
		
		m_objDB = New ProQdb.ProtocolPatternLevelDB
		m_objDB.Load(strID)
		
		'now, load the levels
		LoadTests()
		
	End Sub
	
	
	
	'+
	'LoadFromObject()
	'
	'copies the values out of the db object into this object, and then
	'constructs child objects
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub LoadFromObject(ByRef objDB As ProQdb.ProtocolPatternLevelDB)
		
		'copy the inval object
		m_objDB = objDB
		
		LoadTests()
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
	
	Public Sub SetName(ByRef strName As String)
		m_objDB.SetName(strName)
	End Sub
	
	Public Sub SetLevel(ByRef bytLevel As Byte)
		m_objDB.SetLevel(bytLevel)
	End Sub
	
	Public Sub SetRefLevel(ByRef bytRefLevel As Byte)
		m_objDB.SetRefLevel(bytRefLevel)
	End Sub
	
	Public Sub SetPatternID(ByRef strPatternID As String)
		m_objDB.SetPatternID(strPatternID)
	End Sub
	
	Public Sub SetPercent(ByRef sngPercent As Single)
		m_objDB.SetPercent(sngPercent)
	End Sub
	Public Sub SetPercentStructure(ByRef intStructure As Short)
		m_objDB.SetPercentStructure(intStructure)
	End Sub
	
	Public Function GetID() As String
		GetID = m_objDB.GetID
	End Function
	
	Public Function GetName() As String
		GetName = m_objDB.GetName
	End Function
	
	Public Function GetLevel() As String
        'GetLevel = CStr(m_objDB.GetLevel)
        On Error GoTo GetLevel_Error
        GetLevel = CStr(m_objDB.GetLevel)
        Exit Function
GetLevel_ERROR:
        GetLevel = 0
	End Function
	
	Public Function GetRefLevel() As String
        'GetRefLevel = CStr(m_objDB.GetRefLevel)
        On Error GoTo GetRefLevel_ERROR
        GetRefLevel = CStr(m_objDB.GetRefLevel)
        Exit Function
GetRefLevel_ERROR:
        GetRefLevel = 0
    End Function
	
	Public Function GetPatternID() As String
		GetPatternID = m_objDB.GetPatternID
	End Function
	
	Public Function GetPercent() As String

        On Error GoTo GetPercent_ERROR
        GetPercent = (m_objDB.GetPercent)
        Exit Function
GetPercent_ERROR:
        GetPercent = 0
    End Function
	Public Function GetPercentStructure() As Short
        'GetPercentStructure = m_objDB.GetPercentStructure()
        On Error GoTo GetPercentStructure_ERROR
        GetPercentStructure = m_objDB.GetPercentStructure()
        Exit Function
GetPercentStructure_ERROR:
        GetPercentStructure = 0
    End Function
	
	Public Function GetTests() As Collection
		GetTests = m_cTests
	End Function
End Class