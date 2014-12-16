Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("UseFunding_NET.UseFunding")> Public Class UseFunding
	'UseFunding.cls
	'
	'This class manages the available funding per use by a
	'Funding Source
	'
	'Jleiner
	'07 June 2002
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	Dim g_objFC As New FundingCalculator
	
	'related db object
	Private m_objDB As New ProQdb.UseFundingDB
	Public m_objUse As New Use
	
	Private m_fNewRecord As Boolean
	Private m_fIsDirty As Boolean
	Private m_fAllUses As Boolean
	Private m_curAllocated As Double
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'public properties
	
	'none
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	'+
	'LoadFromID()
	'
	'creates and loads the db object referenced by the given id, and then
	'constructs child objects
	'
	'lbailey
	'1 june 2002
	'-
	Private Sub LoadFromID(ByRef varValue As Object)
		
		'load the object
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_objDB.Load(CStr(varValue))
		
		'Load the Use Object
		m_objUse.Load(m_objDB.GetUseID)
		
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
		
		'Load the Use Object
		m_objUse.Load(m_objDB.GetUseID)
		
	End Sub
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	
	
	'+
	'Create()
	'
	'Creates a new Use Funding Record
	'
	'jleiner
	'7 june 2002
	'-
	Public Sub Create(ByRef strFundingSourceID As String)
		
		m_fNewRecord = True
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUID.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_objDB.SetID(g_objGUID.GetGUID())
		m_objDB.SetFundingSourceID(strFundingSourceID)
		
	End Sub
	
	Public Sub Delete()
		
		If m_fNewRecord = False Then
			m_objDB.Delete()
		End If
		
	End Sub
	
	Public Sub Save()
		
		If m_fAllUses = False Then
			If m_fNewRecord = True Then
				m_objDB.Create()
			ElseIf m_fIsDirty = True Then 
				m_objDB.Update()
			End If
		End If
		m_fNewRecord = False
		m_fIsDirty = False
		
	End Sub
	
	
	'+
	'Load()
	'
	'loads up The use Funding Object from the UseFundingDB Object
	'
	'lbailey
	'1 june 2002
	'-
    'Public Sub Load(ByRef varValue As Object)
    'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '	If (IsReference(varValue) = True) Then
    'load from db object
    '		LoadFromObject((varValue))
    '	Else
    'load from id
    '		LoadFromID((varValue))
    '	End If
    'End Sub
    Public Sub Load(ByRef varValue As Object)
        LoadFromObject(varValue)
    End Sub

    Public Sub Load(ByRef varValue As String)
        LoadFromID(CStr(varValue))
    End Sub
    '+
    'standard manipulator procedures
    '
    'lbailey
    '1 june 2002
    '-
    Public Function SetID(ByRef strID As String) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetID, strID)
        m_objDB.SetID(strID)
        SetID = S_OK
    End Function
    Public Function SetFundingSourceID(ByRef strID As String) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetFundingSourceID, strID)
        m_objDB.SetFundingSourceID(strID)
        SetFundingSourceID = S_OK
    End Function
    Public Function SetUseID(ByRef strID As String) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetUseID, strID)
        m_objDB.SetUseID(strID)

        If strID <> "*" Then
            'Load the Use
            m_objUse.Load(strID)
        End If

        SetUseID = S_OK
    End Function

    Public Function SetPercent(ByRef sngPercent As Single) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetPercent, sngPercent)
        m_objDB.setPercent(sngPercent)
        SetPercent = S_OK
    End Function

    Public Function SetValue(ByRef dblValue As Double) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetValue, dblValue)
        m_objDB.SetValue(dblValue)
        SetValue = S_OK
    End Function

    Public Function SetNotes(ByRef strNotes As String) As Short
        m_objDB.SetNotes(strNotes)
        SetNotes = S_OK
    End Function
    Public Sub setForAllUses(ByRef fAllUses As Boolean)
        m_fAllUses = fAllUses
    End Sub
    Public Sub SetAllocated(ByRef curAllocated As Double)
        m_curAllocated = curAllocated
    End Sub
    Public Function SetIsDirty(ByRef fIsDirty As Object) As Boolean
        'UPGRADE_WARNING: Couldn't resolve default property of object fIsDirty. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_fIsDirty = fIsDirty
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

    Public Function GetFundingSourceID() As String
        GetFundingSourceID = m_objDB.GetFundingSourceID
    End Function

    Public Function GetUseID() As String
        GetUseID = m_objDB.GetUseID()
    End Function

    Public Function GetPercent() As Single
        GetPercent = m_objDB.GetPercent()
    End Function

    Public Function GetValue() As Double
        GetValue = m_objDB.GetValue()
    End Function

    Public Function GetNotes() As String
        GetNotes = m_objDB.GetNotes()
    End Function

    Public Function GetForAllUses() As Boolean
        GetForAllUses = m_fAllUses
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get these values from the Aggregation
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Public Function getAllocated() As Double
    '  'getAllocated = m_curAllocated
    '  getAllocated = g_objFC.GetAllocated_UseFunding(Me)
    'End Function

    'Public Function getAvailable() As Double
    'getAvailable = m_objDB.GetValue - m_curAllocated
    '  getAvailable = m_objDB.GetValue - getAllocated()
    'End Function
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get ABOVE values from the Aggregation
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Function GetIsNew() As Boolean
        GetIsNew = m_fNewRecord
    End Function

    Public Function GetIsDirty() As Boolean
        GetIsDirty = m_fIsDirty
    End Function

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
        m_fIsDirty = False
        m_fNewRecord = False
    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Terminate_Renamed()
        'UPGRADE_NOTE: Object m_objDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objDB = Nothing
    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub
End Class