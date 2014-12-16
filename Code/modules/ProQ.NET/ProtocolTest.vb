Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolTest_NET.ProtocolTest")> Public Class ProtocolTest
	'ProtocolStep.cls
	'
	'This class manages each step of each level of the protocol, accumlating all
	'of the brands in this tests at this level.
	'
	'question to answer: by brand, how many tests will there be?
	'
	'lbailey
	'3 june 2002
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'the persistent data
	Private m_objDB As ProQdb.ProtocolTestDB
	
	'a collection of brands associated with this test
	Private m_cBrands As New Collection
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	
	'+
	'LoadBrands()
	'
	'gets a collection of all of the brands associated with this test
	'
	'lbailey
	'1 june 2002
	'-
	Private Sub LoadBrands()
		
		'setup the parameter array for the sp
		Dim aParms(1, 1) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object aParms(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParms(0, 0) = New Guid(m_objDB.GetID())
		'UPGRADE_WARNING: Couldn't resolve default property of object aParms(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParms(0, 1) = DbType.Guid
		
		'instantiate the dbcoll object and pass it the parm array
		Dim objChildDBColl As New ProQdb.ProtocolBrandCollectionDB
		Dim cDBObjects As New Collection
		Dim objDBObject As New ProQdb.ProtocolBrandDB
		cDBObjects = objChildDBColl.GetCollection(SP_GET_PROTOCOLBRANDS_BY_TEST, aParms)
		
		'now throw the dbobjects into the level objects
        For Each objDBObject In cDBObjects
            'create a middle tier object
            Dim objChild As New ProtocolBrand
            'load it with the db object
            objChild.Load(objDBObject)

            'put it in the collection
            m_cBrands.Add(objChild)

            'UPGRADE_NOTE: Object objChild may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            objChild = Nothing
        Next objDBObject
		
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
	Private Sub LoadFromID(ByRef strID As String)
		
		'load the object
		m_objDB.Load(strID)
		
		'now, load the levels
		LoadBrands()
		
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
	Public Sub LoadFromObject(ByRef objProtocolTestDB As ProQdb.ProtocolTestDB)
		
		'copy the inval object
		m_objDB = objProtocolTestDB
		
		'now, load the levels
		LoadBrands()
		
	End Sub
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	
	
	'+
	'Create()
	'
	'creates this object, and ties it to its parent
	'
	'lbailey
	'7 june 2002
	'-
	Public Sub Create(ByRef strProtocolLevelID As String, ByRef objPatternTest As ProtocolPatternTest)
		
		'create a db object
		m_objDB = New ProQdb.ProtocolTestDB
		m_objDB.Create()
		m_objDB.SetLevelID(strProtocolLevelID)
        m_objDB.SetName(objPatternTest.GetName)
		
		'create a single empty brand for the test
		Dim objProtocolBrand As New ProtocolBrand
		objProtocolBrand.Create(m_objDB.GetID())
		m_cBrands.Add(objProtocolBrand)
		
	End Sub
	
	
	
	'+
	'Load()
	'
	'loads up all of the data about this test, then calls load on the
	'brands associated with this test
	'
	'lbailey
	'4 june 2002
	'-
	Public Sub Load(ByRef strID As String)
		'load from id
		LoadFromID(strID)
	End Sub
	
	
	
	'+
	'GetProtocolBrand()
	'
	'returns the test referenced by the id passed in.  A protocolbrand is a brand
	'plus its usage info (particular to this quantification)
	'
	'lbailey
	'1 june 2002
	'-
	Public Function GetBrand(ByRef strID As String) As ProtocolBrand
		
		Dim objBrand As ProtocolBrand
		For	Each objBrand In m_cBrands
			If (objBrand.GetID() = strID) Then
				GetBrand = objBrand
				Exit Function
			End If
		Next objBrand
		
		'return null if we couln't find what they were looking for
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        GetBrand = Nothing
		
	End Function
	
	
	
	'+
	'AddBrand()
	'
	'adds a brand to the test.  this creates a protocolbrand from a
	'a brand - not a protocolbrand. all of the non-protocolbrand fields
	'are set from local values.  the brand is then added to the quantity
	'
	'lbailey
	'5 june 2002
	'-
	Public Function AddBrand(ByRef strBrandID As String, Optional ByRef sngPercent As Single = 0, Optional ByRef fGeneric As Boolean = False) As String
		
		'create a protocolbrand object
		Dim objProtocolBrand As New ProtocolBrand
		'create an object and give it a parent id
		objProtocolBrand.Create(m_objDB.GetID())
		'feed it the info it needs
		objProtocolBrand.SetBrand(strBrandID)
		objProtocolBrand.SetPercent(sngPercent)
		objProtocolBrand.SetGeneric(fGeneric)
		
		'get the brand info into memory
		objProtocolBrand.GetBrand().Load(strBrandID)
		
		'add this object to the collection
		m_cBrands.Add(objProtocolBrand)
		
		'return the id of the new protocol brand
		AddBrand = objProtocolBrand.GetID()
		
	End Function
	
	
	
	'+
	'AddProtocolBrand()
	'
	'creates and adds a protocolbrand to the test.  all of the
	'non-protocolbrand fields are set from local values.
	'the brand is then added to the quantity
	'
	'lbailey
	'1 july 2002
	'-
	Public Function AddProtocolBrand(Optional ByRef sngPercent As Single = 0, Optional ByRef fGeneric As Boolean = False) As String
		
		'create a protocolbrand object
		Dim objProtocolBrand As New ProtocolBrand
		'create an object and give it a parent id
		objProtocolBrand.Create(m_objDB.GetID())
		'feed it the info it needs
		objProtocolBrand.SetPercent(sngPercent)
		objProtocolBrand.SetGeneric(fGeneric)
		
		'add this object to the collection
		m_cBrands.Add(objProtocolBrand)
		
		'return the id of the new protocol brand
		AddProtocolBrand = objProtocolBrand.GetID()
		
	End Function
	
	
	'+
	'Save()
	'
	'pushes the data out to the db
	'
	'lbailey
	'27 june 2002
	'-
	Public Sub Save()
		'the persistent data
		'Private m_objDB As ProtocolTestDB
		m_objDB.Update()
		
		'a collection of brands associated with this test
		'Private m_cBrands As New Collection
		Dim objProtocolBrand As ProtocolBrand
		For	Each objProtocolBrand In m_cBrands
			objProtocolBrand.Save()
		Next objProtocolBrand
		
	End Sub
	
	
	
	'+
	'DeleteBrand
	'
	'looks for the specified brand by ID in the collection of
	'brands associated with this test.  if found, it is deleted.
	'
	'lbailey
	'22 june 2002
	'-
    Public Sub DeleteBrand(ByRef strProtocolBrandID As String)
        Dim i As Short
        i = 1
        Dim objProtocolBrand As ProtocolBrand
        For Each objProtocolBrand In m_cBrands

            If (objProtocolBrand.GetID() = strProtocolBrandID) Then
                'remove the item from the collection of
                m_cBrands.Remove(i)
                'delete the brand object
                objProtocolBrand.Delete()
                'UPGRADE_NOTE: Object objProtocolBrand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objProtocolBrand = Nothing
            End If
            i = i + 1
        Next objProtocolBrand
    End Sub
	
	
	
	'+
	'DeleteBrandByName
	'
	'looks for the specified brand by name in the collection of
	'brands associated with this test.  if found, it is deleted.
	'
	'lbailey
	'22 june 2002
	'-
    Public Sub DeleteBrandByName(ByRef strProtocolBrandName As String)

        Dim i As Short
        i = 1
        Dim objProtocolBrand As ProtocolBrand
        For Each objProtocolBrand In m_cBrands
            'Dim objProtocolBrand As ProtocolBrand
            If (objProtocolBrand.GetBrand().GetName() = strProtocolBrandName) Then
                'remove the item from the collection of
                m_cBrands.Remove(i)
                'delete the brand object
                objProtocolBrand.Delete()
                'UPGRADE_NOTE: Object objProtocolBrand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objProtocolBrand = Nothing
            End If
            i = i + 1
        Next objProtocolBrand
    End Sub
	
	
	
	'+
	'standard manipulator procedures
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub SetName(ByRef strName As String)
		m_objDB.SetName(strName)
	End Sub
	
	Public Sub SetNotes(ByRef strNotes As String)
		m_objDB.SetNotes(strNotes)
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
	
	Public Function GetLevelID() As String
		GetLevelID = m_objDB.GetLevelID()
	End Function
	
	Public Function GetName() As String
		GetName = m_objDB.GetName()
	End Function
	
	Public Function GetNotes() As String
		GetNotes = m_objDB.GetNotes()
	End Function
	
	Public Function GetBrands() As Collection
		GetBrands = m_cBrands
	End Function
End Class