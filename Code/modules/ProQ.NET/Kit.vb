Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Kit_NET.Kit")> Public Class Kit
	'+
	'Kit.cls
	'
	'represents a test kit.  contains both the lookup info as well
	'as the information pertaining to this particular use of the kitS
	'(plural).
	'
	'lbailey
	'8 june 2002
	'-
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	Private m_objDB As New ProQdb.KitDB
	
	
	
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
		m_objDB.Load(strID)
		
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
    'Public Sub Load(ByRef varValue As Object)
    'Dim strTest As String
    '    strTest = TypeName(varValue)
    '    If (strTest <> "String") Then
    'load from db object
    '        LoadFromObject(varValue)
    '    Else
    'load from id
    'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '       LoadFromID(CStr(varValue))
    '    End If
    'End Sub
    Public Sub Load(ByRef varValue As Object)
        LoadFromObject(varValue)
    End Sub

    Public Sub Load(ByRef varValue As String)
        LoadFromID(CStr(varValue))
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
            If MsgBox("Are you sure you want to delete kit, '" & Me.GetTestsPerKit() & "'?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "DELETE???") = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        'Check to see if object is used in various places if so mark as inactive.
        If KitinUse() = True Then
            'The kit is presently in use in One or more Quantifications.
            If MsgBox("The kit is currently in use in one or more Quantifications. " & "It can not be deleted at this time.  It can be marked as inactive, " & "and be hidden from Quantifications that do not already contain it." & vbCrLf & vbCrLf & "Mark as Inactive?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "KIT IN USE") = MsgBoxResult.Yes Then
                m_objDB.setInactive(True)
                Update()
            End If
        Else
            ' Delete the object
            m_objDB.Delete()
        End If
    End Sub

    Public Function KitinUse() As Boolean

        Dim objKitCol As New ProQdb.KitDBCollection
        Dim mCol As Collection

        Dim aParams(1, 1) As Object

        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(m_objDB.GetID())
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid

        mCol = objKitCol.GetCollection(SP_QSELKITINUSE, aParams)

        If mCol.Count() > 0 Then
            KitinUse = True
        Else
            KitinUse = False
        End If

        'UPGRADE_NOTE: Object objKitCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objKitCol = Nothing
        'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mCol = Nothing

    End Function


    '+
    'standard accessor functions
    '
    'lbailey
    '1 june 2002
    '-
    'look-up info
    'Private m_nTestsPerKit As Integer
    'Private m_nQA As Integer
    'Private m_cCost As Double



    Public Function GetID() As String
        'GetID = m_objDB.GetID()
        If m_objDB Is Nothing Then
            GetID = ""
        Else
            GetID = m_objDB.GetID()
        End If
    End Function

    Public Function GetTestsPerKit() As Short
        'GetTestsPerKit = m_objDB.GetTestsPerKit()
        If m_objDB Is Nothing Then
            GetTestsPerKit = 0
        Else
            GetTestsPerKit = m_objDB.GetTestsPerKit()
        End If
    End Function

    Public Function GetKitDimension() As String
        GetKitDimension = m_objDB.GetKitDimension()
    End Function

    Public Function GetWeight() As Double
        GetWeight = m_objDB.GetWeight()
    End Function

    Public Function GetCost() As Double
        GetCost = m_objDB.GetKitCost()
    End Function
    Public Function GetHeight() As Double
        GetHeight = m_objDB.GetHeight()
    End Function
    Public Function GetWidth() As Double
        GetWidth = m_objDB.GetWidth()
    End Function
    Public Function GetLength() As Double
        GetLength = m_objDB.GetLength()
    End Function
    Public Function GetVolume() As Double
        With m_objDB
            GetVolume = .GetHeight() * .GetWidth() * .GetLength()
        End With
    End Function


    '+
    'standard manipulator procedures
    '
    'lbailey
    '1 june 2002
    '-

    Public Sub SetTestsPerKit(ByRef intTestsPerKit As Short)
        m_objDB.SetTestsPerKit(intTestsPerKit)
    End Sub

    Public Sub SetCost(ByRef cCost As Double)
        m_objDB.SetKitCost(CDbl(cCost))
    End Sub
    Public Sub SetWeight(ByRef dblWeight As Double)
        m_objDB.SetWeight(dblWeight)
    End Sub
    Public Sub SetBrandID(ByRef strBrandID As String)
        m_objDB.SetBrandID(strBrandID)
    End Sub

    Public Sub SetKitDimensions(ByRef strKitDimension As String)
        m_objDB.SetKitDimension(strKitDimension)
    End Sub

    Public Sub SetHeight(ByRef dblHeight As Double)
        m_objDB.SetHeight(dblHeight)
    End Sub
    Public Sub SetLength(ByRef dblLength As Double)
        m_objDB.SetLength(dblLength)
    End Sub
    Public Sub SetWidth(ByRef dblWidth As Double)
        m_objDB.SetWidth(dblWidth)
    End Sub
End Class