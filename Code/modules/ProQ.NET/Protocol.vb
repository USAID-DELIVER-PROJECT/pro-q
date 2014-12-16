Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Protocol_NET.Protocol")> Public Class Protocol
	'Protocol.cls
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
	'constants
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	Private m_strID As String
	Private m_strQuantificationID As String
	Private m_fIsNew As Boolean
	
	'connection to the db
	Private m_objDB As ProQdb.ProtocolDB
	
	'collection of levels associated with this protocol
	Private m_cLevels As New Collection
	
	'all of the brands and counts for this protocol
	Private m_objSelectedBrands As New SelectedBrands
	
	Private m_objPattern As ProtocolPattern
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'public properties
	
	'none
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	
	'+
	'Class_Initialize()
	'
	'invoked automatically on class creation.  consider this the
	'constructor
	'
	'lbailey
	'2 june 2002
	'-
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'get a connection to the backend
		m_objDB = New ProQdb.ProtocolDB
		m_fIsNew = True
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	
	'+
	'LoadLevels()
	'
	'
	'lbailey
	'1 june 2002
	'-
	Private Sub LoadLevels()
		
		'setup the parameter array for the sp
		Dim aParams(1, 1) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(m_objDB.GetID())
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid
		
		'instantiate the dbcoll object and pass it the parm array
		Dim objProtLevelDBColl As New ProQdb.ProtocolLevelDBCollection
		Dim cDBObjects As Collection
		cDBObjects = objProtLevelDBColl.GetCollection(SP_GET_PROTOCOLLEVELS_BY_PROTOCOLID, aParams)
		
		'now throw the dbobjects into the level objects
		Dim objDBObject As ProQdb.ProtocolLevelDB
        For Each objDBObject In cDBObjects
            Dim objProtocolLevel As New ProtocolLevel
            'create a middle tier object
            'load it with the db object
            'objProtocolLevel.Load objDBObject
            objProtocolLevel.LoadFromObject(objDBObject)
            'put it in the collection
            m_cLevels.Add(objProtocolLevel)
            'release the instance
            'UPGRADE_NOTE: Object objProtocolLevel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            objProtocolLevel = Nothing
        Next objDBObject
		
	End Sub
	
	
	'+
	'LoadSelectedBrand()
	'
	'creates a selected brand manager object and uses it to load the collection
	'of selected brands associated with this protocol quantification
	'
	'lbailey
	'1 july 2002
	'-
	Public Sub LoadSelectedBrands()
		m_objSelectedBrands.Load(m_strQuantificationID)
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
		
		'load the object
		m_objDB.Load(strID)
		
		'now, load the levels
		LoadLevels()
		
		'now load the selected brands
		LoadSelectedBrands()
		
		m_fIsNew = False
		
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
	Public Sub LoadFromObject(ByRef objDB As ProQdb.ProtocolDB)
		
		'copy the inval object
		m_objDB = objDB
		
		'now, load the levels
		LoadLevels()
		
		'now load the selected brands
		LoadSelectedBrands()
		
		m_fIsNew = False
		
	End Sub
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	
	'+
	'Create()
	'
	'creates a new protocol patterned after the specified pattern
	'for the specified quantification.
	'
	'lbailey
	'7 june 2002
	'-
	Public Sub Create(ByRef objQuantification As Quantification, ByRef objPattern As ProtocolPattern)
		
		m_objPattern = objPattern
		
		m_strQuantificationID = objQuantification.GetID
		
		Dim cPatternLevels As Collection
		Dim objPatternLevel As ProtocolPatternLevel
        If (Len(objQuantification.GetProtocolID()) > 0) Then

            Debug.Print("the prot id is " & objQuantification.GetProtocolID())

            'the given quant alread has a protocol
            Err.Description = "This quantification already has a protocol."
            'Err.Raise 0
            Exit Sub
        Else
            'the given quant does not yet have a prot,so create one
            m_objDB = New ProQdb.ProtocolDB
            m_objDB.Create()

            'stuff the fields
            m_objDB.SetDiagram(objPattern.GetDiagram)
            m_objDB.SetName(objPattern.GetName)
            m_objDB.SetNotes(objPattern.GetNotes)
            m_objDB.SetPatternID(objPattern.GetID)
            m_objDB.SetSamples(0)

            'use the input pattern to create this prot
            cPatternLevels = objPattern.GetLevels()
            'use this pattern to add levels to this protocol
            For Each objPatternLevel In cPatternLevels
                Dim objProtocolLevel As New ProtocolLevel
                'create a level
                objProtocolLevel.Create(GetID(), objPatternLevel, objQuantification.getDiscordancy, objQuantification.getPrevalence)
                'and add it to this protocol
                m_cLevels.Add(objProtocolLevel)
                'UPGRADE_NOTE: Object objProtocolLevel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objProtocolLevel = Nothing
            Next objPatternLevel

            'now tell the quanification that we exist
            objQuantification.SetProtocolID(m_objDB.GetID())

        End If
		
		m_fIsNew = False
		
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
		
		'load the pattern object
		m_objPattern = New ProtocolPattern
		m_objPattern.LoadByID(m_objDB.GetPatternID)
		
		m_fIsNew = False
		'TODO: do we need to load the "selected brand manager" here?
		
		
	End Sub
	
	
	'+
	'Update()
	
	'Updates the database with this object's properties
	'
	'lbailey
	'8 june 2002
	'-
	Public Sub Update()
		
		'save the protocol object
		m_objDB.Update()
		
		'save the protocol levels
		Dim objProtocolLevel As ProtocolLevel
		For	Each objProtocolLevel In m_cLevels
			objProtocolLevel.Save()
		Next objProtocolLevel
		
		'save the selected brands
		Calculate()
		ConsolidateBrands()
		m_objSelectedBrands.Save()
		
	End Sub
	
	Public Sub Save()
		'save everything
		Update()
	End Sub
	
	
	
	'+
	'GetLevel()
	'
	'looks for the level specified, and returns it.  if it doesn't find it,
	'bad things happen. (the function returns "nothing")
	'
	'lbailey
	'
	Public Function GetLevel(ByRef intLevelNum As Short) As ProtocolLevel
		
		'this will hold the results
		Dim objProtocolLevel As ProtocolLevel
		
		'ensure that we've been asked for a valid level
		Dim objLevel As ProtocolLevel
		If (intLevelNum > m_cLevels.Count()) Then
			'if we haven't just return the first one
			SortLevels()
			objProtocolLevel = m_cLevels.Item(1)
		Else
			'look for the specified level
			For	Each objLevel In m_cLevels
				If (objLevel.GetLevel() = intLevelNum) Then
					'this is the one we want
					objProtocolLevel = objLevel
					'stop looking
					Exit For
				End If
			Next objLevel
		End If
		
		'return the results
		GetLevel = objProtocolLevel
		
	End Function
	
	
	
	'+
	'Calculate()
	'
	'
	Public Sub Calculate(Optional ByRef lSamples As Integer = -1)
		
		'calculate the number of each brand in each test in each level
		If lSamples = -1 Then
			lSamples = GetNumberSamples()
		End If
		
		'ensure that the levels are in the right order so that the loops
		'can work smoothly
		SortLevels()
		
		'declare the variables that we'll be using
		Dim lRefCount As Object
		Dim lCount As Integer
		Dim nRefLevel As Short
		Dim aCounts() As Integer
		ReDim aCounts(m_cLevels.Count())
		Dim sngPercent As Single
		
		Dim objTest As ProtocolTest
		Dim objBrand As ProtocolBrand
		
		
		'set first level count
		aCounts(1) = GetNumberSamples()
		
		'-------------------------------------------------------------
		' Get the counts for the individual First Level Tests.
		' - not sure if this is the best way, probably could occur
		' - in loop with other levels.  But seems to work - jleiner
		'-------------------------------------------------------------
		'UPGRADE_WARNING: Couldn't resolve default property of object m_cLevels().GetTests. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each objTest In m_cLevels.Item(1).GetTests()
			For	Each objBrand In objTest.GetBrands()
				'get the percentage for that brand
				sngPercent = objBrand.GetPercent() / 100
				'multiply it out
                lCount = sngPercent * aCounts(1)
				'set the count
				objBrand.SetCount(lCount)
			Next objBrand
		Next objTest
		
		
		'if there are more levels, then calc the counts for these levels
		Dim i As Short
		If (m_cLevels.Count() > 1) Then
			For i = 2 To m_cLevels.Count()
				
				'get refid of level we want to take the pct of
				'UPGRADE_WARNING: Couldn't resolve default property of object m_cLevels().GetRefLevel. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nRefLevel = m_cLevels.Item(i).GetRefLevel()
				'get the count from that level
				'UPGRADE_WARNING: Couldn't resolve default property of object lRefCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lRefCount = aCounts(nRefLevel)
				'from this level, get the percentage
				'UPGRADE_WARNING: Couldn't resolve default property of object m_cLevels().GetPercent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sngPercent = m_cLevels.Item(i).GetPercent() / 100
				
				'calculate the count for this level
				'UPGRADE_WARNING: Couldn't resolve default property of object lRefCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				aCounts(i) = sngPercent * lRefCount
				
				'now work out the count per brand
				'UPGRADE_WARNING: Couldn't resolve default property of object m_cLevels().GetTests. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For	Each objTest In m_cLevels.Item(i).GetTests()
					For	Each objBrand In objTest.GetBrands()
						'get the percentage for that brand
						sngPercent = objBrand.GetPercent() / 100
						'multiply it out
						lCount = sngPercent * aCounts(i)
						'set the count
						objBrand.SetCount(lCount)
					Next objBrand
				Next objTest
			Next i
		End If
		
	End Sub
	
	
	'+
	'GetRawBrands()
	'
	'returns a collection of all the brands in all of the tests in all of
	'the levels, without accumlating any of the values
	'
	'lbailey
	'27 june 2002
	'-
	Public Function GetRawBrands() As Object
		
		Dim objProtocolLevel As ProtocolLevel
		Dim objProtocolTest As ProtocolTest
		Dim objProtocolBrand As ProtocolBrand
		Dim cBrands As New Collection
		
		'in every level...
		For	Each objProtocolLevel In m_cLevels
			'...get the list of tests from the level...
			For	Each objProtocolTest In objProtocolLevel.GetTests()
				'...get each brand from the test...
				For	Each objProtocolBrand In objProtocolTest.GetBrands()
					'... and put it in the new collection
					cBrands.Add(objProtocolBrand)
				Next objProtocolBrand
			Next objProtocolTest
		Next objProtocolLevel
		
		'return the collection
		GetRawBrands = cBrands
		
	End Function
	
	
	
	'+
	'ResetBrands()
	'
	'removes any brands that are no longer present in the protocol, and sets
	'all of the counts down to zero
	'
	'lbailey
	'27 june 2002
	'-
	Private Sub ResetBrands()
		
		'get a collection of all of the brand objects in the protocol
		Dim cBrands As New Collection
		cBrands = GetRawBrands()
		Dim fMatchFound As Boolean
		
		'a new set of selected brands that we'll copy valid elements into
		Dim objSelectedBrands As New SelectedBrands
		
		'go see if our selected brands are still in the protocol
		Dim objSelectedBrand As SelectedBrand
		Dim objProtocolBrand As ProtocolBrand
		Dim strSelBrandID As Object
		Dim strProtBrandID As String
		Dim strSelBrandKitID As Object
		Dim strProtBrandKitID As String
        'Dim strSelBrandGenericCode As String
        'Dim strProtBrandGenericCode As String
		
		'Pass the Quantification ID to the new SelectedBrands Object
		objSelectedBrands.SetQuantificationID(m_strQuantificationID)
		
		For	Each objSelectedBrand In m_objSelectedBrands.GetSelectedBrands()
			'compare each sel brand to all of the prot brands until
			'you find a match
			fMatchFound = False
			
			For	Each objProtocolBrand In cBrands
				
				'get the brand ids from each of the two objects
				'UPGRADE_WARNING: Couldn't resolve default property of object strSelBrandID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strSelBrandID = objSelectedBrand.GetBrandID()
				'UPGRADE_WARNING: Couldn't resolve default property of object strSelBrandKitID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strSelBrandKitID = objSelectedBrand.GetKitID
				strProtBrandID = objProtocolBrand.GetBrand().GetID()
				strProtBrandKitID = objProtocolBrand.GetKit.GetID
				
				'if the selbrand is in the prot, then copy it to new coll
				'UPGRADE_WARNING: Couldn't resolve default property of object strSelBrandKitID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object strSelBrandID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (strSelBrandID = strProtBrandID) And (strSelBrandKitID = strProtBrandKitID) And objSelectedBrand.getGeneric = objProtocolBrand.getGeneric Then
					objSelectedBrands.AddBrand(objSelectedBrand)
					fMatchFound = True
					Exit For
				End If
				
			Next objProtocolBrand
			
			If fMatchFound = False Then
				'Delete the SelectedBrand
				objSelectedBrand.Delete()
			End If
		Next objSelectedBrand
		
		'now, zero out the counts
		For	Each objSelectedBrand In objSelectedBrands.GetSelectedBrands()
			objSelectedBrand.SetCount(0)
		Next objSelectedBrand
		
		m_objSelectedBrands = objSelectedBrands
		
	End Sub
	
	
	
	'+
	'ConsolidateBrands()
	'
	'resets the list of selected brands, consolidates the tests under the levels,
	'and consolidates the levels in this protocol
	'
	'lbailey
	'5 june 2002
	'-
	Public Sub ConsolidateBrands()
		
		Dim objLevel As ProtocolLevel
		Dim objTest As ProtocolTest
		Dim objProtocolBrand As ProtocolBrand
		
		'clear the list of brands
		'm_objSelectedBrands.RemoveAllBrands
		ResetBrands()
		
		'in every level...
		For	Each objLevel In m_cLevels
			'...get the list of tests from the level...
			For	Each objTest In objLevel.GetTests()
				'...get each brand from the test...
				For	Each objProtocolBrand In objTest.GetBrands()
					'... and add it to the collection
					m_objSelectedBrands.AccumulateBrand(objProtocolBrand)
				Next objProtocolBrand
			Next objTest
		Next objLevel
	End Sub
	
	
	
	'+
	'Delete()
	'
	'removes the protocol from the db
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub Delete(Optional ByRef fAsk As Boolean = False)
		
		If fAsk = True Then
			If MsgBox("Are you sure you want to delete the selected Protocol? ", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "DELETE?") = MsgBoxResult.No Then
				Exit Sub
			End If
		End If
		
		m_objDB.Delete()
		
	End Sub
	
	
	
	'+
	'GetTests()
	'
	'returns a collection of tests that belong to this protocol
	'
	'lbailey
	'15 june 2002
	'-
	Public Function GetTests() As Collection
		
		'declare a holder for all of the tests
		Dim cProtocolTests As New Collection
		'let the compiler know that we have collections of levels
		Dim objLevel As ProtocolLevel
		'iterate through the levels...
        Dim objTest As ProtocolTest
        For Each objLevel In m_cLevels
            Dim cTests As New Collection
            '...getting all of the tests from each level...
            cTests = objLevel.GetTests()
            For Each objTest In cTests
                Dim objProtocolTest As ProtocolTest
                objProtocolTest = objTest
                '...and then put the tests in our collection.
                cProtocolTests.Add(objProtocolTest)
                'UPGRADE_NOTE: Object objProtocolTest may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objProtocolTest = Nothing
            Next objTest
            'UPGRADE_NOTE: Object cTests may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            cTests = Nothing
        Next objLevel
		
		GetTests = cProtocolTests
		
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
		Dim objLevel As ProtocolLevel 'the cur level
		sngWeight = 0# 'start clean
		sngPrevPct = 1# 'first levels tests all samps
		
		'ensure that the "previous level" is what it should be
		SortLevels()
		
		For	Each objLevel In m_cLevels
			
			'calc the percentage for this level
			sngPct = (objLevel.GetPercent() / 100) * sngPrevPct
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
	Private Sub SortLevels()
		
		'holder for the sorted collection
		Dim cTemp As New Collection
		
		'the objects used in the sorting process
		Dim objLevel As New ProtocolLevel
		Dim objSortLevel As New ProtocolLevel
		
		'the loops used to sort
		Dim nLevels As Short
		nLevels = m_cLevels.Count()
		Dim nCurLevel As Short
		For nCurLevel = 1 To nLevels
			
			For	Each objLevel In m_cLevels
				If (objLevel.GetLevel() = nCurLevel) Then
					objSortLevel = objLevel
				End If
			Next objLevel
			cTemp.Add(objSortLevel)
			
		Next nCurLevel
		
		m_cLevels = cTemp
	End Sub
	Public Sub UpdatePercents(ByRef sngDiscordancy As Single, ByRef sngPrevalence As Single)
		
		'the objects used in the updating process
		Dim objLevel As New ProtocolLevel
		
		'for each level in the protocol, update the default percent and save
		For	Each objLevel In m_cLevels
			Select Case objLevel.GetPercentStructure
				Case 2
					objLevel.SetDefaultPercent(True, sngDiscordancy, sngPrevalence)
				Case Else
					objLevel.SetDefaultPercent(True, sngDiscordancy, sngPrevalence)
			End Select
			objLevel.Save()
		Next objLevel
		
	End Sub
	
	
	
	'+
	'standard accessors
	'
	'lbailey
	'2 june 2002
	'-
	Public Sub SetName(ByRef strName As String)
		m_objDB.SetName(strName)
	End Sub
	
	Public Sub SetNotes(ByRef strNotes As String)
		m_objDB.SetNotes(strNotes)
	End Sub
	
	Public Sub SetDiagram(ByRef strDiagram As String)
		m_objDB.SetDiagram(strDiagram)
	End Sub
	
	Public Sub SetNumberSamples(ByRef lNumberSamples As Integer)
		m_objDB.SetSamples(lNumberSamples)
		Calculate()
        ConsolidateBrands()
	End Sub
	
	Public Sub SetSelectedBrands(ByRef objSelectedBrands As SelectedBrands)
		m_objSelectedBrands = objSelectedBrands
	End Sub
	
	Public Sub SetPatternID(ByRef strPatternID As String)
		m_objDB.SetPatternID((strPatternID))
	End Sub
	
	Public Sub SetQuantificationID(ByRef strQuantificationID As String)
		m_strQuantificationID = strQuantificationID
	End Sub
	
	Public Sub SetIsNew(ByRef fIsNew As Boolean)
		m_fIsNew = fIsNew
	End Sub
	
	'+
	'standard manipulators
	'
	'lbailey
	'2 june 2002
	'-
	Public Function GetID() As String
		GetID = m_objDB.GetID()
	End Function
	
	Public Function GetPatternID() As String
		GetPatternID = m_objDB.GetPatternID()
	End Function
	
	Public Function GetName() As String
		GetName = m_objDB.GetName()
	End Function
	
	Public Function GetNotes() As String
		GetNotes = m_objDB.GetNotes()
	End Function
	
	Public Function GetDiagram() As String
		GetDiagram = m_objDB.GetDiagram()
	End Function
	
	Public Function GetNumberSamples() As Integer
		GetNumberSamples = m_objDB.GetSamples()
	End Function
	
	Public Function GetSelectedBrands() As SelectedBrands
		GetSelectedBrands = m_objSelectedBrands
	End Function
	
	Public Function GetLevels() As Collection
		GetLevels = m_cLevels
	End Function
	
	Public Function IsNew() As Boolean
		IsNew = m_fIsNew
	End Function
End Class