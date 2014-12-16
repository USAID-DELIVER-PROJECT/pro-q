Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Carton_NET.Carton")> Public Class Carton
	
	'+
	'Carton.cls
	'
	'represents a carton of kits
	'
	'lbailey
	'26 may 2002
	'-
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	Private m_objDB As New ProQdb.CartonDB
	
	
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
        'If (IsReference(varValue) = True) Then
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
        'If (IsReference(varValue) = True) Then
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
	Public Sub Delete()
		m_objDB.Delete()
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
	
	Public Function GetKitID() As String
		GetKitID = m_objDB.GetKitID
	End Function
	
	Public Function GetNumKits() As Short
		GetNumKits = m_objDB.GetKitsPerCarton()
	End Function
	
	Public Function GetWeight() As Double
		GetWeight = m_objDB.GetWeight()
	End Function
	
	Public Function GetHeight() As Double
		GetHeight = m_objDB.GetHeight()
	End Function
	
	Public Function GetLength() As Double
		GetLength = m_objDB.GetLength()
	End Function
	
	Public Function GetWidth() As Double
		GetWidth = m_objDB.GetWidth()
	End Function
	
	
	
	'+
	'standard manipulator procedures
	'
	'lbailey
	'1 june 2002
	'-
	
	Public Sub SetKitID(ByRef strKitID As String)
		m_objDB.SetKitID(strKitID)
	End Sub
	
	Public Sub SetNumKits(ByRef nNumKits As Short)
		m_objDB.SetKitsPerCarton(nNumKits)
	End Sub
	
	Public Sub SetWeight(ByRef dblWeight As Double)
		m_objDB.SetWeight(dblWeight)
	End Sub
	
	Public Sub SetHeight(ByRef dblHeight As Double)
		m_objDB.SetHeight(dblHeight)
	End Sub
	
	Public Sub SetWidth(ByRef dblWidth As Double)
		m_objDB.SetWidth(dblWidth)
	End Sub
	
	Public Sub SetLength(ByRef dblLength As Double)
		m_objDB.SetLength(dblLength)
	End Sub
End Class