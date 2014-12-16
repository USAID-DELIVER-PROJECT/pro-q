Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolBrand_NET.ProtocolBrand")> Public Class ProtocolBrand
	'ProtocolTest
	'
	'Represents a particulat test in a level of a protocol.  these are not
	'lookups, but are actually data that has been gathered as part of an
	'aggregation.  the protocol itself may be a pattern that we can lookup,
	'but this stuff is the brands related to the step, percentages of the
	'overall number of tests for each brand, and the number of tests
	'conducted at this step (not loaded for qa, wastage, or any of that)
	'
	'
	'lbailey
	'31 may 2002
	'
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'the persistent data
	Private m_objDB As New ProQdb.ProtocolBrandDB
	
	'information pertaining to this product
	Private m_objBrand As New Brand
	Private m_objKit As New Kit
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	
	'+
	'destructor
	'
	'cleans up any loose ends (memory, db conns, etc)
	'
	'lbailey
	'31 may 2002
	'-
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'clean up memory
		'delete each brand from the array
		'delete the array
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	
	'+
	'LoadFromID()
	'
	'creates and loads the db object referenced by the given id
	'
	'lbailey
	'4 june 2002
	'-
	Private Sub LoadFromID(ByRef varValue As Object)
		
		'load the object
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_objDB.Load(varValue)
		
		'load the brand
		LoadBrand()
		
		'loads the kit
		LoadKit()
	End Sub
	
	
	
	'+
	'LoadFromObject()
	'
	'copies the values out of the db object into this object
	'
	'lbailey
	'4 june 2002
	'-
	Private Sub LoadFromObject(ByRef varValue As Object)
		
		'copy the inval object
		m_objDB = varValue
		
		'load the brand
		LoadBrand()
		
		'loads the kit
		LoadKit()
		
		
	End Sub
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                        public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	
	'+
	'Create()
	'
	'the protocolbrand needs to know which test it belongs to
	'
	'lbailey
	'-
	Public Sub Create(ByRef strProtocolTestID As String)
		m_objDB.Create()
		m_objDB.SetTestID(strProtocolTestID)
	End Sub
	
	
	
	'+
	'Load()
	'
	'loads up all of the data about this brand
	'
	'lbailey
	'4 june 2002
	'-
    Public Sub Load(ByRef varValue As Object)
        'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        'If (IsReference(varValue) = True) Then
        'load from db object
        LoadFromObject(varValue)
        'Else
        'load from id
        'LoadFromID(varValue)
        ' End If
    End Sub
    Public Sub Load(ByRef varValue As String)
        'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        'If (IsReference(varValue) = True) Then
        'load from db object
        'LoadFromObject(varValue)
        'Else
        'load from id
        LoadFromID(varValue)
        'End If
    End Sub

	
	
	'+
	'LoadBrand()
	'
	'loads the branddb object that this object represents an
	'instance of
	'
	'lbailey
	'12 june 2002
	'-
	Public Sub LoadBrand()
		
		'load the brand object
		m_objBrand.Load(m_objDB.GetBrandID())
		
	End Sub
	
	
	
	'+
	'LoadKit()
	'
	'loads the branddb object that this object represents an
	'instance of
	'
	'lbailey
	'12 june 2002
	'-
	Public Sub LoadKit()
		
		'load the brand object
		Dim strKitID As String
		strKitID = m_objDB.GetKitID()
		If (Len(strKitID) > 1) Then
			m_objKit.Load(m_objDB.GetKitID())
		End If
		
	End Sub
	
	
	
	'+
	'Save()
	'
	'pushes the data out to the db
	'
	'lbailey
	'28 june 2002
	'-
	Public Sub Save()
		m_objDB.Update()
	End Sub
	
	
	
	'DeleteBrand()
	'
	'deletes this object from the db
	'
	'lbailey
	'31 may 2002
	'
	Public Sub Delete()
		m_objDB.Delete()
	End Sub
	
	
	
	'+
	'Child Object Accessors
	'
	'lbailey
	'04 july 2002
	'-
	Public Function GetBrand() As Brand
		GetBrand = m_objBrand
	End Function
	
	Public Function GetKit() As Kit
		GetKit = m_objKit
	End Function
	
	
	
	'+
	'Child Object Manipulators
	'
	'lbailey
	'04 july 2002
	'-
	
	'should be a set kit
	'should be a set brand
	
	
	'+
	'Protocol Brand Accessors
	'
	'lbailey
	'4 july 2002
	'-
	Public Function GetID() As String
		GetID = m_objDB.GetID()
	End Function
	
	Public Function GetName() As String
		GetName = m_objDB.GetName()
	End Function
	
	Public Function IsGeneric() As Boolean
		IsGeneric = m_objDB.getGeneric()
	End Function
	
	Public Function getGeneric() As Boolean
		getGeneric = m_objDB.getGeneric()
	End Function
	Public Function getGenericCode() As String
		getGenericCode = m_objDB.getGenericCode
	End Function
	
	Public Function GetPercent() As Single
		GetPercent = m_objDB.GetPercent()
	End Function
	
	Public Function GetCount() As Integer
		GetCount = m_objDB.GetCount()
	End Function
	
	Public Function GetNotes() As String
		GetNotes = m_objDB.GetNotes()
	End Function
	
	
	'+
	'Protocol Brand manipulators
	'
	'lbailey
	'4 july 2002
	'-
	
	Public Sub SetName(ByRef strName As String)
		m_objDB.SetName(strName)
	End Sub
	
	Public Sub SetGeneric(ByRef fGeneric As Boolean)
		m_objDB.SetGeneric(fGeneric)
	End Sub
	Public Sub SetGenericCode(ByRef strCode As String)
		m_objDB.SetGenericCode(strCode)
	End Sub
	
	
	Public Sub SetKit(ByRef strKitID As String)
		m_objDB.SetKitID(strKitID)
		LoadKit()
	End Sub
	
	Public Sub SetBrand(ByRef strBrandID As String)
		m_objDB.SetBrandID(strBrandID)
		If (Len(strBrandID) > 1) Then
			m_objBrand.Load(strBrandID)
			
			'Dim objBrand As New Brand
			'objBrand.Load strBrandID
			'Set m_objBrand = objBrand
			'Set objBrand = Nothing
		End If
	End Sub
	
	Public Sub SetCount(ByRef lngCount As Integer)
		m_objDB.SetCount(lngCount)
	End Sub
	
	Public Sub SetPercent(ByRef sngPercent As Single)
		m_objDB.SetPercent(sngPercent)
	End Sub
	
	Public Sub SetNotes(ByRef strNotes As String)
		m_objDB.SetNotes(strNotes)
	End Sub
End Class