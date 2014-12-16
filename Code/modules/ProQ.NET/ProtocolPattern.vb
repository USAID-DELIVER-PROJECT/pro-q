Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolPattern_NET.ProtocolPattern")> Public Class ProtocolPattern
	'ProtocolPattern.cls
	'
	'the protocol holds both all of the lookup information, and all of
	'the user-input info regarding the number of samples to be tested,
	'and the numbers and types of tests that will be needed to accomplish
	'this.
	'
	'lbailey
	'1 june 2002
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'connection to the db
	Private m_objDB As ProQdb.ProtocolPatternDB
	
	'collection of levels associated with this protocol
	Private m_cLevels As New Collection
	
	'quantities associated with this protocol
	Private m_objQuantity As New Quantity
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	
	
	'+
	'LoadLevels()
	'
	'adds the children to the object
	'
	'lbailey
	'1 june 2002
	'-
	Private Sub LoadLevels()
		
		'setup the parameter array for the sp
		Dim aParms(1, 1) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object aParms(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParms(0, 0) = New Guid(m_objDB.GetID())
		'UPGRADE_WARNING: Couldn't resolve default property of object aParms(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParms(0, 1) = DbType.Guid
		
		'instantiate the dbcoll object and pass it the parm array
		Dim objLevelCollection As New ProQdb.ProtocolPatternLevelDBCollection
		Dim cLevelDB As Collection
		'get the collection of db objects
		cLevelDB = objLevelCollection.GetCollection(SP_GET_PROTPATLEV_BY_PROTPAT_ID, aParms)
		
		'create a collection of objects from the dbobjects
		Dim objLevelDB As ProQdb.ProtocolPatternLevelDB
        For Each objLevelDB In cLevelDB
            Dim objLevel As New ProtocolPatternLevel
            objLevel.LoadByID(objLevelDB.GetID)
            'put each object in our collection
            m_cLevels.Add(objLevel)
            'UPGRADE_NOTE: Object objLevel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            objLevel = Nothing
        Next objLevelDB
		
	End Sub
	
	
	
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
		
		'instantiate a db object
		m_objDB = New ProQdb.ProtocolPatternDB
		
		'go gets its info from the db
		m_objDB.Load(strID)
		
		LoadLevels()
		
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
	Public Sub LoadFromObject(ByRef objDB As ProQdb.ProtocolPatternDB)
		
		'copy the inval object
		m_objDB = objDB
		
		LoadLevels()
		
	End Sub
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	
	'+
	'Create()
	'
	'we won't be creating protocols patterns until we write the admin
	'screens.
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
	
	
	
	'+
	'standard manipulators
	'
	'lbailey
	'7 june 2002
	'-
	Public Sub SetName(ByRef strName As String)
		m_objDB.SetName(strName)
	End Sub
	
	Public Sub SetNotes(ByRef strNotes As String)
		m_objDB.SetNotes(strNotes)
	End Sub
	
	Public Sub SetDiagram(ByRef oleDiagram As String)
		m_objDB.SetDiagram(oleDiagram)
	End Sub
	
	
	
	'+
	'standard accessors
	'
	'lbailey
	'7 june 2002
	'-
	Public Function GetID() As String
		GetID = m_objDB.GetID()
	End Function
	
	Public Function GetName() As String
		GetName = m_objDB.GetName()
	End Function
	
	Public Function GetNotes() As String
		GetNotes = m_objDB.GetNotes()
	End Function
	
	Public Function GetDiagram() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object GetDiagram. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetDiagram = m_objDB.GetDiagram()
	End Function
	
	Public Function GetLevels() As Collection
		GetLevels = m_cLevels
	End Function
	
	'+
	'GetTests()
	'
	'returns a collection of tests that belong to this pattern
	'
	'lbailey
	'15 june 2002
	'-
	Public Function GetTests() As Collection
		
		'declare a holder for all of the tests
		Dim cPatternTests As New Collection
		'let the compiler know that we have collections of levels
		Dim objLevel As ProtocolPatternLevel
		'iterate through the levels...
		Dim cTests As New Collection
		Dim objTest As ProtocolPatternTest
        For Each objLevel In m_cLevels
            '...getting all of the tests from each level...
            cTests = objLevel.GetTests()
            For Each objTest In cTests
                Dim objProtocolTest As ProtocolPatternTest
                objProtocolTest = objTest
                '...and then put the tests in our collection.
                cPatternTests.Add(objProtocolTest)
                'UPGRADE_NOTE: Object objProtocolTest may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objProtocolTest = Nothing
            Next objTest
        Next objLevel
		
		GetTests = cPatternTests
		
	End Function
	
	
	
	'+
	'GetWeight()
	'
	'returns a value which represents the number of tests per sample
	'which will be retqured for the current protocol.
	'
	'lbailey
	'19 june 2002
	'-
	Public Function GetWeight() As Single
		
		Dim sngWeight As Single ' the result
		Dim nLevelTests As Short
		Dim sngPrevPct As Single 'the number of samples * pct
		Dim sngPct As Single 'the cur level's pct
        Dim objLevel As ProtocolPatternLevel 'the cur level
		sngWeight = 0# 'start clean
		sngPrevPct = 1# 'first levels tests all samps
		
		SortLevels()
		
		For	Each objLevel In m_cLevels
			
			'calc the percentage for this level
			sngPct = (CDbl(objLevel.GetPercent()) / 100) * sngPrevPct
			'remember it for the next loop
			sngPrevPct = sngPct
			
			'get the number of tests at this level
			nLevelTests = objLevel.GetTests().Count()
			
			'multiply out the number of tests at this level and accum
			sngWeight = sngWeight + (nLevelTests * sngPct)
			
		Next objLevel
		
		GetWeight = sngWeight
		
	End Function
	
	
	
	'+
	'SortLevels()
	'
	'a very simple sort routine to ensure that the collection of levels
	'is sorted by level number.  if it isn't, then "getweight" tends to
	'generate some rather fanciful weights.
	'
	'lbailey
	'24 june 2002
	'-
	Public Sub SortLevels()
		
		'holder for the sorted collection
		Dim cTemp As New Collection
		
		'the objects used in the sorting process
		Dim objLevel As New ProtocolPatternLevel
		Dim objSortLevel As New ProtocolPatternLevel
		
		'the loops used to sort
		Dim nLevels As Short
		nLevels = m_cLevels.Count()
		Dim nCurLevel As Short
		For nCurLevel = 1 To nLevels
			
			For	Each objLevel In m_cLevels
				If (CDbl(objLevel.GetLevel()) = nCurLevel) Then
					objSortLevel = objLevel
				End If
			Next objLevel
			cTemp.Add(objSortLevel)
			
		Next nCurLevel
		
		m_cLevels = cTemp
	End Sub
End Class