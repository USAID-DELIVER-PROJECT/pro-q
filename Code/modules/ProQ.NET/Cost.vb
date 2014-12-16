Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Cost_NET.Cost")> Public Class Cost
	'private properties
	
	'brand costs
	'logistics costs
	'customs costs
	
	'Cost.cls
	'
	'The Costs associated with a Quantification
	'
	'Jleiner
	'07 June 2002
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	Private Const TYPE_KIT As Short = 1
	Private Const TYPE_CUSTOMS As Short = 2
	Private Const TYPE_STORAGE As Short = 3
	
	Public Enum Cost_Method
		cmNone = 0
		cmPercent = 1
		cmMethod1 = 2
		cmMethod2 = 3
		cmOverride = 4
	End Enum
	
	Private Const SP_FUNDING_BY_QUANTIFICATION_COST As String = "qselFundingbyQuantificationCostType"
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'related db object
	Private m_objDB As New ProQdb.CostDB
	Private m_fNewRecord As Boolean
	Private m_fIsDirty As Boolean
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'public properties
	Public m_colFundings As New Fundings
	Public m_objCurrentFunding As New Funding
	
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
        m_objDB.Load(varValue)
		
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
	'Creates a new Use Cost Record
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
		
		m_objDB.Delete()
		
	End Sub
	
	Public Sub Save()
		
		Dim objFunding As Funding
		
		If m_fNewRecord = True Then
			m_objDB.Create()
		Else
			m_objDB.Update()
		End If
		
		m_fIsDirty = False
		m_fNewRecord = False
		
		For	Each objFunding In m_colFundings
			objFunding.Save()
		Next objFunding
		
	End Sub
	
	'+
	'Load()
	'
	'loads up The use Cost Object from the CostDB Object
	'
	'lbailey
	'1 june 2002
	'-
    'Public Sub Load(ByRef varValue As Object)
    'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '	If (IsReference(varValue) = True) Then
    'load from db object
    '			LoadFromObject((varValue))
    '		Else
    'load from id
    '			LoadFromID((varValue))
    '		End If

    'Load the Funding Collection
    '		LoadFundings()

    '	End Sub
    Public Sub Load(ByRef varValue As Object)
        LoadFromObject(varValue)
        LoadFundings()
    End Sub

    Public Sub Load(ByRef varValue As String)
        LoadFromID(CStr(varValue))
        LoadFundings()
    End Sub
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''Funding Collection Stuff
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '---------------------------------------------------------
    'LoadFundings()
    ' Comments    : Find all the Fundings that exist
    '             : for the Cost Record and load
    '             : into collection of Fundings
    ' Parameters  : None
    ' Returns     : Integer S_OK
    ' Created     : 07-June-02 jleiner
    '---------------------------------------------------------
    Function LoadFundings() As Short

        Dim objCol As New ProQdb.FundingDBCollection
        Dim objFundingDB As New ProQdb.FundingDB
        Dim mCol As Collection

        Dim strStoredProc As String
        Dim aParams(2, 1) As Object

        strStoredProc = SP_FUNDING_BY_QUANTIFICATION_COST

        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(m_objDB.getQuantificationID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid

        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(1, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(1, 0) = m_objDB.GetCategory()
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(1, 1) = DbType.Int16

        mCol = objCol.GetCollection(strStoredProc, aParams)

        For Each objFundingDB In mCol
            m_objCurrentFunding = m_colFundings.Add(objFundingDB.GetID)
            m_objCurrentFunding.Load(objFundingDB.GetID)
        Next objFundingDB

        'UPGRADE_NOTE: Object objCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objCol = Nothing
        'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mCol = Nothing

    End Function
    ' Create a new use (Funding) for the Aggregate
    Public Function AddFunding() As Short

        m_objCurrentFunding = m_colFundings.Add()
        m_objCurrentFunding.SetQuantificationID(m_objDB.getQuantificationID())
        m_objCurrentFunding.SetCategory(m_objDB.GetCategory())

        'Prompt User for Configuration and redo tree.
        AddFunding = S_OK

    End Function

    ' Delete a use (Funding) from the Cost
    Public Function DeleteFunding() As Short

        m_colFundings.Remove(m_objCurrentFunding.GetID)
        m_objCurrentFunding.Delete()

        'UPGRADE_NOTE: Object m_objCurrentFunding may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objCurrentFunding = Nothing

    End Function
    Public Function GetCurrentFunding() As Funding
        GetCurrentFunding = m_objCurrentFunding
    End Function
    'SetCurrentFunding m_objCurrentFunding
    Public Function SetCurrentFunding(ByRef objFunding As Funding) As Short
        m_objCurrentFunding = objFunding
        SetCurrentFunding = S_OK
    End Function

    'SetCurrentFunding m_strCurrentFunding
    Public Function SetCurrentFundingbyID(ByRef strID As String) As Short
        m_objCurrentFunding = m_colFundings(strID)
        SetCurrentFundingbyID = S_OK
    End Function
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Function must be run from Quantificition level as value is not known
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Public Function GetUnallocatedAmount() As Double
    ' Comments  : Returns the amount required - the amount allocated
    ' Parameters:
    ' Returns   : Double - Amount still needed
    ' Created   : 08-June-2002
    '---------------------------------------------------
    '  GetUnallocatedAmount = m_objDB.GetValue() - _
    ''        GetAllocatedAmount()
    'End Function

    Public Function GetAllocatedAmount() As Double
        ' Comments  : Polls the funding collection for the cost
        '           : and returns the total amount allocated
        ' Parameters:
        ' Returns   : Double - Total Amount Allocated to Cost
        ' Created   : 08-June-2002
        '---------------------------------------------------
        Dim objFunding As Funding
        Dim curAllocated As Double

        For Each objFunding In m_colFundings
            curAllocated = curAllocated + objFunding.GetValue
        Next objFunding

        GetAllocatedAmount = curAllocated

    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    Public Function SetMethod(ByRef lngMethod As Cost_Method) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetMethod, lngMethod)
        m_objDB.setMethod(lngMethod)
        SetMethod = S_OK
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

    Public Function GetQuantificationID() As String
        GetQuantificationID = m_objDB.getQuantificationID()
    End Function

    Public Function GetCategory() As Integer
        GetCategory = m_objDB.GetCategory()
    End Function

    Public Function GetMethod() As Integer
        GetMethod = m_objDB.GetMethod()
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
        m_fIsDirty = False
        m_fNewRecord = False
    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Terminate_Renamed()
        'UPGRADE_NOTE: Object m_colFundings may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_colFundings = Nothing
        'UPGRADE_NOTE: Object m_objCurrentFunding may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objCurrentFunding = Nothing
        'UPGRADE_NOTE: Object m_objDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objDB = Nothing
    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub
End Class