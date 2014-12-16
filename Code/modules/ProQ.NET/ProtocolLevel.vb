Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolLevel_NET.ProtocolLevel")> Public Class ProtocolLevel
	'ProtocolLevel.cls
	'
	'This class manages each level of the protocol, accumlating all
	'of the brands in the various tests at this level.
	'
	'question to answer: by brand, how many tests will there be?
	'
	'lbailey
	'31 may 2002
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'related db object
	Private m_objDB As New ProQdb.ProtocolLevelDB
	
	'the collection of children
	Private m_cTests As New Collection
	
	'the number of samples that to base the calculation on
	Private m_lReferenceCount As Integer
	'the percentage of ref counts that we'll be testing
	Private m_sngPercent As Single
	'the number of samples to be tested at this level
	Private m_lCount As Integer
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	
	
	'+
	'LoadTests()
	'
	'
	'lbailey
	'1 june 2002
	'-
	Private Sub LoadTests()
		
		'setup the parameter array for the sp
		Dim aParms(1, 1) As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object aParms(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParms(0, 0) = New Guid(m_objDB.GetID())
		'UPGRADE_WARNING: Couldn't resolve default property of object aParms(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParms(0, 1) = DbType.Guid
		
		'instantiate the dbcoll object and pass it the parm array
		Dim objChildDBColl As New ProQdb.ProtocolTestDBCollection
		Dim cDBObjects As Collection
		cDBObjects = objChildDBColl.GetCollection(SP_GET_PROTOCOLTESTS_BY_LEVEL, aParms)
		
		'now throw the dbobjects into the level objects
		Dim objDBObject As ProQdb.ProtocolTestDB
        For Each objDBObject In cDBObjects
            Dim objChild As New ProtocolTest

            'create a middle tier object
            'load it with the db object
            objChild.LoadFromObject(objDBObject)
            'put it in the collection
            m_cTests.Add(objChild)
            'release the instance
            'UPGRADE_NOTE: Object objChild may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            objChild = Nothing

        Next objDBObject
		
	End Sub
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	
	'+
	'LoadFromID()
	'
	'creates and loads the db object referenced by the given id, and then
	'constructs child objects
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub LoadByID(ByRef strID As String)
		
		'load the object
		'Set m_objDB = objDB.Load(varValue)
		m_objDB.Load(strID)
		
		'now, load the children
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
	Public Sub LoadFromObject(ByRef objDB As ProQdb.ProtocolLevelDB)
		
		'copy the inval object
		m_objDB = objDB
		
		'now, load the children
		LoadTests()
		
	End Sub
	
	
	
	
	'+
	'Create()
	'
	'goes and gets the pattern and creates a record in the protocol
	'level table.
	'
	'lbailey
	'5 june 2002
	'-
	Public Function Create(ByRef strProtocolID As String, ByRef objPatternLevel As ProtocolPatternLevel, ByRef sngDiscordancy As Single, ByRef sngPrevalence As Single) As String
		
		'create the db object
		m_objDB.Create()
		m_objDB.SetLevel(CShort(objPatternLevel.GetLevel))
		'm_objDB.SetPercent objPatternLevel.GetPercent
		m_objDB.SetProtocolID(strProtocolID)
		m_objDB.SetRefLevel(CShort(objPatternLevel.GetRefLevel))
		m_objDB.SetPercentStructure(objPatternLevel.GetPercentStructure)
		
		SetDefaultPercent(UseDiscordance(objPatternLevel), sngDiscordancy, sngPrevalence)
		
		'now, go make babies (add tests to the level)
		'now add the tests to the levels
		Dim cTests As Collection
		cTests = objPatternLevel.GetTests()
		Dim objPatternTest As ProtocolPatternTest
        For Each objPatternTest In cTests
            Dim objTest As New ProtocolTest
            'create a new test tied to this level
            objTest.Create(m_objDB.GetID(), objPatternTest)
            'addtest to this level
            AddTest(objTest)
            'UPGRADE_NOTE: Object objTest may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            objTest = Nothing
        Next objPatternTest
		
		Create = m_objDB.GetID
		
	End Function
	
	
	
	'+
	'Load()
	'
	'loads up all of the data about this level, then calls load on the
	'tests at this level
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub Load(ByRef strID As String)
		'load from id
		LoadByID(strID)
	End Sub
	
	
	
	'+
	'Save()
	'
	'pushes the data out to the backend
	'
	'lbailey
	'27 june 2002
	'-
	Public Sub Save()
		
		'these guys aren't being saved right now
		'Private m_lReferenceCount As Long
		'Private m_lCount As Long
		
		'Percent not copying over to ProtocolLevel Table.  Had to relookup what the percent should be - LKB 8/27/02
		m_sngPercent = m_objDB.GetPercent
		m_objDB.SetPercent(m_sngPercent)
		m_objDB.Update()
		
		Dim objProtocolTest As ProtocolTest
		For	Each objProtocolTest In m_cTests
			objProtocolTest.Save()
		Next objProtocolTest
		
		
	End Sub
	Public Function UseDiscordance(ByRef objPatternLevel As ProtocolPatternLevel) As Boolean
		'Captures the name of the patterntest and determines if it is C or greater.  If true then
		'we will need to add discordance to HIV Prevalence rate for level two of protocol.  Level 3
		'will always be discordance.
		'
		'lblanken 8/27/2002
		
		'UseDiscordance = False
		'Dim cTests As Collection
		'Set cTests = objPatternLevel.GetTests()
		'Dim objPatternTest As ProtocolPatternTest
		'For Each objPatternTest In cTests
		'    If Asc(objPatternTest.GetName) >= 67 Then
		'        UseDiscordance = True
		'    End If
		'Next
		UseDiscordance = True
		
	End Function
	
	
	'+
	'AddTest()
	'
	'adds a test to the collection of tests in this obj
	'
	'lbailey
	'8 june 2002
	'-
	Public Sub AddTest(ByRef objTest As ProtocolTest)
		m_cTests.Add(objTest)
	End Sub
	
	
	'+
	'GetTest()
	'
	'returns the test referenced by the id passed in
	'
	'lbailey
	'1 june 2002
	'-
	Public Function GetTest(ByRef strTestID As String) As ProtocolTest
		
		Dim objReturn As New ProtocolTest
		Dim objTest As ProtocolTest
		
		'loop through the collection, finding the one with this id
        For Each objTest In m_cTests
            'UPGRADE_WARNING: Couldn't resolve default property of object m_cTests.objTest. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

            If (strTestID = objTest.GetID()) Then
                'set the return value
                objReturn = objTest
                'leave this loop
                Exit For
            End If
        Next objTest
		
		GetTest = objReturn
		
	End Function
	
	
	
	
	'+
	'standard manipulator procedures
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub SetLevel(ByRef intLevel As Short)
		m_objDB.SetLevel(intLevel)
	End Sub
	
	Public Sub SetRefLevel(ByRef intRefLevel As Short)
		m_objDB.SetRefLevel(intRefLevel)
	End Sub
	
	Public Sub SetProtocolID(ByRef strProtocolID As String)
		m_objDB.SetProtocolID(strProtocolID)
		'CalculateCount
	End Sub
	
	Public Sub SetDefaultPercent(ByRef fDiscordance As Boolean, ByRef sngDiscordancy As Single, ByRef sngPrevalence As Single)
		
		'set default values for the levels and the structure used to create them
		
		'Select Case GetLevel()
		'Case 1
		'    m_objDB.SetPercent 100
		'    m_objDB.SetPercentStructure 0
		'Case 2
		'    'level two is always AIDS prev
		'    If fDiscordance Then
		'        m_objDB.SetPercent sngPrevalence + sngDiscordancy
		'        m_objDB.SetPercentStructure 2
		'    Else
		'        m_objDB.SetPercent sngPrevalence
		'        m_objDB.SetPercentStructure 1
		'    End If
		'Case Else
		'    m_objDB.SetPercent sngDiscordancy
		'    m_objDB.SetPercentStructure 3
		'End Select
		
		Select Case GetPercentStructure
			Case 0
				'No Factoring (100% Tests occur)
				m_objDB.SetPercent(100)
			Case 1
				'Prevalence Only
				m_objDB.SetPercent(sngPrevalence)
			Case 2
				'Prevalence + Discordancy
				m_objDB.SetPercent(sngPrevalence + sngDiscordancy)
			Case 3
				'Discordancy Only
				m_objDB.SetPercent(sngDiscordancy)
		End Select
		
	End Sub
	Public Sub SetPercent(Optional ByRef sngPercent As Single = NO_INPUT, Optional ByRef sngDiscordancy As Single = PCT_DISCORDANCE)
		'set default values for the levels if none is specified
		
		If (sngPercent = NO_INPUT) Then
			m_objDB.SetPercent(GetPercent)
		Else
			Select Case GetPercentStructure
				Case 0
					'0 Level
				Case 1
					'Percent Only
					m_objDB.SetPercent(sngPercent)
				Case 2
					'Percent + Discordancy
					m_objDB.SetPercent(sngPercent + sngDiscordancy)
				Case 3
					'Discordancy Only
					m_objDB.SetPercent(sngDiscordancy)
			End Select
			
			
		End If
	End Sub
	
	'+
	'standard accessor functions
	'
	'lbailey
	'1 june 2002
	'-
	Public Function GetID() As String
		GetID = m_objDB.GetID()
	End Function
	
	Public Function GetLevel() As Short
		GetLevel = m_objDB.GetLevel()
	End Function
	
	Public Function GetPercent() As Single
		GetPercent = m_objDB.GetPercent()
	End Function
	Public Function GetPercentStructure() As Short
		GetPercentStructure = m_objDB.GetPercentStructure()
	End Function
	Public Function GetRefLevel() As Short
		GetRefLevel = m_objDB.GetRefLevel()
	End Function
	
	Public Function GetTests() As Collection
		GetTests = m_cTests
	End Function
End Class