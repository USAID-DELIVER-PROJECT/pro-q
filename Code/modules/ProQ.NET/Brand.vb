Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Brand_NET.Brand")> Public Class Brand
	'Brand.cls
	'
	'represents the lookup information about a brand.  this object
	'does not hold any info that is relevant to the current
	'quantification
	'
	'lbailey
	'5 june 2002
	'
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'persistent data
	Private m_objDB As New ProQdb.BrandDB
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	
	'+
	'LoadFromID()
	'
	'reaches down to the db objects and gets the specified brand
	'
	'lbailey
	'1 june 2002
	'-
	Private Sub LoadFromID(ByRef strID As String)
		
		'get the info from the db into the db wrapper
        m_objDB.Load((strID))
		
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
	Private Sub LoadFromObject(ByRef varValue As Object)
		
		'copy the inval object
		m_objDB = varValue
		
	End Sub
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	
	'+
	'Load(byref inVal as variant)
	'
	'determines whether db object or a string is being passed in, and then
	'calls the appropriate load function to load the object
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub Load(ByRef varValue As Object)
        'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        'Dim strTest As String
        'strTest = TypeName(varValue)
        'If (strTest <> "String") Then
        'load from db object
        LoadFromObject(varValue)
        'Else
        'load from id
        'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'LoadFromID(CStr(varValue))
        'End If
    End Sub
	
    Public Sub Load(ByRef varValue As String)
        'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        'Dim strTest As String
        'strTest = TypeName(varValue)
        'If (strTest <> "String") Then
        'load from db object
        'LoadFromObject(varValue)
        'Else
        'load from id
        'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        LoadFromID(CStr(varValue))
        'End If
    End Sub
	'+
	'Create()
	'
	'creates a new object (and generates reference id)
	'
	'lbailey
	'8 june 2002
	'-
	Public Function Create() As String
		Create = m_objDB.Create
	End Function
	
	
	'+
	'Update()
	
	'Updates the database with this object's properties
	'
	'lbailey
	'8 june 2002
	'-
	Public Sub Update()
		m_objDB.Update()
	End Sub
	
	
	
	'+
	'Delete()
	'
	'deletes the object from the db
	'
	'lbailey
	'8 june 2002
	'-
	Public Sub Delete(Optional ByRef fAsk As Boolean = False)
		
		If fAsk = True Then
			'Ask the user to confirm deletion
			If MsgBox("Are you sure you want to delete brand, " & Me.GetName() & "?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "DELETE???") = MsgBoxResult.No Then
				Exit Sub
			End If
		End If
		
		'Check to see if object is used in various places if so mark as inactive.
		If BrandinUse() = True Then
			'The brand is presently in use in One or more Quantifications.
			If MsgBox("The Brand is currently in use in one or more Quantifications. " & "It can not be deleted at this time.  It can be marked as inactive, " & "and be hidden from Quantifications that do not already contain it." & vbCrLf & vbCrLf & "Mark as Inactive?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "BRAND IN USE") = MsgBoxResult.Yes Then
				m_objDB.setInactive(True)
				Update()
			End If
		Else
			' Delete the object
			m_objDB.Delete()
		End If
	End Sub
	
	Public Function BrandinUse() As Boolean
		
		Dim objBrandCol As New ProQdb.BrandCollectionDB
		Dim mCol As Collection
		
		Dim aParams(1, 1) As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(m_objDB.GetID())
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid
		
		mCol = objBrandCol.GetCollection(SP_QSELBRANDINUSE, aParams)
		
		If mCol.Count() > 0 Then
			BrandinUse = True
		Else
			BrandinUse = False
		End If
		
		'UPGRADE_NOTE: Object objBrandCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objBrandCol = Nothing
		'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mCol = Nothing
		
	End Function
	
	
	'+
	'standard accessor functions
	'
	'lbailey
	'1 june 2002
	'-
	Public Function GetID() As String
        GetID = m_objDB.GetID()
	End Function
	
	Public Function GetName() As String
		GetName = m_objDB.GetName()
	End Function
	
	Public Function GetSupplier() As String
		GetSupplier = m_objDB.GetSupplier()
	End Function
	
	'UPGRADE_NOTE: GetType was upgraded to GetType_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetType_Renamed() As Integer
        GetType_Renamed = m_objDB.GetType_Renamed
	End Function
	
	Public Function GetShelfLife() As Short
		GetShelfLife = m_objDB.GetShelfLife()
	End Function
	
	Public Function GetStorageTemp() As String
		GetStorageTemp = m_objDB.GetStorageTemp()
	End Function
	
	Public Function GetColdStorage() As Boolean
		GetColdStorage = m_objDB.GetColdStorage()
	End Function
	
	Public Function GetMaterialsRequired() As String
		GetMaterialsRequired = m_objDB.GetMaterialsRequired()
	End Function
	
	Public Function GetNotes() As String
		GetNotes = m_objDB.GetNotes()
	End Function
	
	Public Function IsLab() As Boolean
		IsLab = m_objDB.IsLab()
	End Function
	
	
	
	'+
	'standard manipulator procedures
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub SetID(ByRef strID As String)
		m_objDB.SetID(strID)
	End Sub
	
	Public Sub SetName(ByRef strName As String)
		m_objDB.SetName(strName)
	End Sub
	
	Public Sub SetSupplier(ByRef strSupplier As String)
		m_objDB.SetSupplier(strSupplier)
	End Sub
	
	Public Sub SetType(ByRef lType As Integer)
		m_objDB.SetType(lType)
	End Sub
	
	Public Sub SetShelfLife(ByRef intShelfLife As Short)
		m_objDB.SetShelfLife(intShelfLife)
	End Sub
	
	Public Sub SetStorageTemp(ByRef dblStorageTemp As String)
		m_objDB.SetStorageTemp(dblStorageTemp)
	End Sub
	
	Public Sub SetColdStorage(ByRef fColdStorage As Boolean)
		m_objDB.SetColdStorage(fColdStorage)
	End Sub
	
	Public Sub SetMaterialsRequired(ByRef strMaterialsRequired As String)
		m_objDB.SetMaterialsRequired(strMaterialsRequired)
	End Sub
	
	Public Sub SetNotes(ByRef strNotes As String)
		m_objDB.SetNotes(strNotes)
	End Sub
	
	Public Sub SetLab(ByRef fIsLab As Boolean)
		m_objDB.SetLab(fIsLab)
	End Sub
End Class