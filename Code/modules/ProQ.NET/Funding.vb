Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Funding_NET.Funding")> Public Class Funding
	'Funding.cls
	'
	'The Funds Allocated For a Quantification
	'
	'Jleiner
	'07 June 2002
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	Private Const TYPE_KIT As Short = 1
	Private Const TYPE_CUSTOMS As Short = 2
	Private Const TYPE_STORAGE As Short = 3
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'related db object
	Private m_objDB As New ProQdb.FundingDB
	
	Private m_fNewRecord As Boolean
	Private m_fIsDirty As Boolean
	
	
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
    Private Sub LoadFromID(ByRef strValue As String)

        'load the object
        m_objDB.Load(strValue)

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
	'Create()
	'
	'Creates a new Use Funding Record
	'
	'jleiner
	'7 june 2002
	'-
	Public Sub Create(ByRef strQuantificationID As String)
		
		m_fNewRecord = True
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUID.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_objDB.SetID(g_objGUID.GetGUID())
		m_objDB.SetQuantificationID(strQuantificationID)
		
	End Sub
	
	Public Sub Delete()
		
		If m_fNewRecord = False Then
			m_objDB.Delete()
		End If
		
	End Sub
	
	Public Sub Save()
		
		If m_fNewRecord = True Then
			m_objDB.Create()
		Else
			m_objDB.Update()
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
    Public Function SetQuantificationID(ByRef strID As String) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.getQuantificationID, strID)
        m_objDB.setQuantificationID(strID)
        SetQuantificationID = S_OK
    End Function

    Public Function SetValue(ByRef dblValue As Double) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetValue, dblValue)
        m_objDB.SetValue(dblValue)
        SetValue = S_OK
    End Function
    Public Function SetCategory(ByRef lngCategory As Integer) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetCategory, lngCategory)
        m_objDB.SetCategory(lngCategory)
        SetCategory = S_OK
    End Function
    Public Function SetNotes(ByRef strNotes As String) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetNotes, strNotes)
        m_objDB.SetNotes(strNotes)
        SetNotes = S_OK
    End Function
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

    Public Function GetQuantificationID() As String
        GetQuantificationID = m_objDB.getQuantificationID()
    End Function

    Public Function GetCategory() As Integer
        GetCategory = m_objDB.GetCategory()
    End Function

    Public Function GetValue() As Double
        GetValue = m_objDB.GetValue()
    End Function

    Public Function GetNotes() As String
        GetNotes = m_objDB.GetNotes()
    End Function

    Public Function GetIsNew() As Boolean
        GetIsNew = m_fNewRecord
    End Function

    Public Function GetIsDirty() As Boolean
        GetIsDirty = m_fIsDirty
    End Function

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
        m_fNewRecord = False
        m_fIsDirty = False
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