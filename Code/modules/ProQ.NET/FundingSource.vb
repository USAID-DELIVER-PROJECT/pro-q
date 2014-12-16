Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("FundingSource_NET.FundingSource")> Public Class FundingSource
	'FundingSource.cls
	'
	'This class manages the Funding Sources
	'
	'Jleiner
	'07 June 2002
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	Private Const SP_USEFUNDING_BY_FUNDINGSOURCE As String = "qlkpUseFundingbyFundingSource"
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'related db object
	Private m_objDB As New ProQdb.FundingSourceDB
	
	Private m_fNewRecord As Boolean
	Private m_fIsDirty As Boolean
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'public properties
	Public colUseFundings As New UseFundings
	Public m_objCurrentUseFunding As UseFunding
	
	
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
	Private Sub LoadFromID(ByRef strValue As string)
		
		'load the object
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'm_objDB.Load(CType((varValue), Guid).ToString()) 'CStr(varValue))
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
	Public Sub Create(ByRef strAggregationID As String)
		
		m_fNewRecord = True
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUID.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_objDB.SetID(g_objGUID.GetGUID())
		m_objDB.SetAggregationID(strAggregationID)
		
		'set all cost components can be funded for a pot of money (only for new funds)
		m_objDB.SetAllowCustomsCosts(True)
		m_objDB.SetAllowKitCosts(True)
		m_objDB.SetAllowStorageCosts(True)
		
	End Sub
	
	'+
	'Save()
	'
	'Save the Funding Source and use Fundings
	'
	'jleiner
	'7 june 2002
	'-
	Public Sub Save()
		
		Dim objUseFunding As UseFunding
		
		If m_fNewRecord = True Then
			m_objDB.Create()
		Else
			m_objDB.Update()
		End If
		
		m_fNewRecord = False
		m_fIsDirty = False
		
		' Save the use fundings too
		For	Each objUseFunding In colUseFundings
			objUseFunding.Save()
		Next objUseFunding
		
	End Sub
	Public Sub Delete()
		
		If m_fNewRecord = False Then
			m_objDB.Delete()
		End If
		
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

    'Load the use Fundings for the Funding Source
    '	LoadUseFundings()

    'End Sub
    Public Sub Load(ByRef varValue As Object)
        LoadFromObject(varValue)
        LoadUseFundings()
    End Sub

    Public Sub Load(ByRef strValue As String)
        LoadFromID(strValue)
        LoadUseFundings()
    End Sub
    '---------------------------------------------------------
    'LoadUseFundings()
    ' Comments    : Find all the UseFundings that exist
    '             : for the loaded Funding Source and Load them
    '             : into collection UseFundings
    ' Parameters  : None
    ' Returns     : Integer S_OK
    ' Created     : 07-June-02 jleiner
    '---------------------------------------------------------
    Function LoadUseFundings() As Short

        Dim objCol As New ProQdb.UseFundingDBCollection

        Dim objUseFundingDB As ProQdb.UseFundingDB
        Dim mCol As Collection

        Dim strStoredProc As String
        Dim aParams(1, 1) As Object

        strStoredProc = SP_USEFUNDING_BY_FUNDINGSOURCE

        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(m_objDB.GetID())
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid

        mCol = objCol.GetCollection(strStoredProc, aParams)
        For Each objUseFundingDB In mCol
            m_objCurrentUseFunding = colUseFundings.Add(objUseFundingDB.GetID)
            m_objCurrentUseFunding.Load(objUseFundingDB.GetID)
        Next objUseFundingDB

        'If colUseFundings.Count = 0 Then
        '  'Add Dummy Record
        '  Set m_objCurrentUseFunding = colUseFundings.Add()
        '  With m_objCurrentUseFunding
        '    .SetFundingSourceID m_objDB.GetID()
        '    .setForAllUses True
        ''    .SetUseID "*"
        '    .SetValue m_objDB.GetValue
        '  End With
        'End If

        'UPGRADE_NOTE: Object objCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objCol = Nothing
        'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mCol = Nothing

    End Function
    ' Create a new use (UseFunding) for the Aggregate
    Public Function AddUseFunding() As Short

        'Dim objUseFunding As UseFunding

        m_objCurrentUseFunding = colUseFundings.Add()
        m_objCurrentUseFunding.SetFundingSourceID(m_objDB.GetID())

        'Prompt User for Configuration and redo tree.
        AddUseFunding = S_OK

        'Check for and Remove Dummy record if there is more than one
        'If colUseFundings.Count > 1 Then
        '
        '  For Each objUseFunding In colUseFundings
        '    If objUseFunding.GetForAllUses = True Then
        '      colUseFundings.Remove objUseFunding.GetID
        '      objUseFunding.Delete
        '      Exit For
        '    End If
        '  Next
        'End If

    End Function

    ' Delete a new use (UseFunding) from the Aggregate
    Public Function DeleteUseFunding() As Short

        colUseFundings.Remove(m_objCurrentUseFunding.GetID)
        m_objCurrentUseFunding.Delete()

        'UPGRADE_NOTE: Object m_objCurrentUseFunding may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objCurrentUseFunding = Nothing

    End Function
    Public Function GetCurrentUseFunding() As UseFunding
        GetCurrentUseFunding = m_objCurrentUseFunding
    End Function
    'SetCurrentUseFunding m_objCurrentUseFunding
    Public Function SetCurrentUseFunding(ByRef objUseFunding As UseFunding) As Short
        m_objCurrentUseFunding = objUseFunding
        SetCurrentUseFunding = S_OK
    End Function

    'SetCurrentUseFundingByID m_strCurrentUseFunding
    Public Function SetCurrentUseFundingbyID(ByRef strID As String) As Short
        m_objCurrentUseFunding = colUseFundings(strID)
        SetCurrentUseFundingbyID = S_OK
    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' USE AGGREGATE FUNCTIONS INSTEAD TO RETURN THIS STUFF
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Public Function GetAvailableFunds() As Double
    ' Comments  : Returns the total amount of Funds for the Funding Source
    '           : less the amount allocated to each quantification
    ' Parameters: -
    ' Returns   : Double - The amount still Available
    ' Created   : 06-June-2002
    '---------------------------------------------------------------
    '    GetAvailableFunds = m_objDB.GetValue - GetAllocatedFunds()
    '
    'End Function


    'Public Function GetAllocatedFunds() As Double
    ' Comments  : Returns the total amount of Funds allocated
    '           : to each quantification.
    ' Parameters: -
    ' Returns   : Double - The Allocated to Costs
    ' Created   : 14-June-2002
    '---------------------------------------------------------------
    '    Dim objFC As New FundingCalculator
    '    GetAllocatedFunds = objFC.GetAllocated_FundingSource_old(Me)
    '    Set objFC = Nothing
    '
    'End Function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' USE AGGREGATE FUNCTIONS INSTEAD TO RETURN THE ABOVE
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Function GetDesignatedFunds() As Double
        ' Comments  : Returns the total amount of Funds designated
        '           : to Use Fundings (if Not By Use, then Return Full Value)
        ' Parameters: -
        ' Returns   : Double - The amount still Available
        ' Created   : 14-June-2002
        '---------------------------------------------------------------

        Dim objUF As UseFunding
        Dim dblFunds As Double

        dblFunds = 0

        If m_objDB.GetFundByUse = True Then
            For Each objUF In colUseFundings
                If objUF.GetUseID <> "*" Then
                    dblFunds = dblFunds + objUF.GetValue()
                End If
            Next objUF
        Else
            dblFunds = m_objDB.GetValue
        End If

        GetDesignatedFunds = dblFunds

    End Function

    Public Function GetUndesignatedFunds() As Double
        ' Comments  : Returns the total amount of Funds Not Specified
        '           : to a Use Funding
        ' Parameters: -
        ' Returns   : Double - The amount still Available
        ' Created   : 14-June-2002
        '---------------------------------------------------------------
        GetUndesignatedFunds = m_objDB.GetValue - GetDesignatedFunds()

    End Function

    Public Function GetIsNew() As Boolean
        GetIsNew = m_fNewRecord
    End Function

    Public Function GetIsDirty() As Boolean
        GetIsDirty = m_fIsDirty
    End Function
    Public Function SetIsDirty(ByRef fIsDirty As Object) As Boolean
        'UPGRADE_WARNING: Couldn't resolve default property of object fIsDirty. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_fIsDirty = fIsDirty
    End Function

    '+
    'standard manipulator procedures
    '
    'lbailey
    '1 june 2002
    '-
    Public Function SetID(ByRef strID As String) As Short
        m_objDB.SetID(strID)
        SetID = S_OK
    End Function
    Public Function SetAggregationID(ByRef strID As String) As Short
        m_objDB.SetAggregationID(strID)
        SetAggregationID = S_OK
    End Function
    Public Function SetName(ByRef strName As String) As Short
        m_objDB.SetName(strName)
        SetName = S_OK
    End Function
    Public Function SetValue(ByRef dblValue As Double) As Short
        m_objDB.SetValue(dblValue)
        SetValue = S_OK
    End Function

    Public Function SetAbbreviation(ByRef strAbbreviation As String) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetAbbreviation, strAbbreviation)
        m_objDB.SetAbbreviation(strAbbreviation)
        SetAbbreviation = CStr(S_OK)
    End Function

    Public Function SetAllowKitCosts(ByRef fAllow As Boolean) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetAllowKitCosts, fAllow)
        m_objDB.SetAllowKitCosts(fAllow)
        SetAllowKitCosts = S_OK
    End Function

    Public Function SetAllowStorageCosts(ByRef fAllow As Boolean) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetAllowStorageCosts, fAllow)
        m_objDB.SetAllowStorageCosts(fAllow)
        SetAllowStorageCosts = S_OK
    End Function

    Public Function SetAllowCustomsCosts(ByRef fAllow As Boolean) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetAllowCustomsCosts, fAllow)
        m_objDB.SetAllowCustomsCosts(fAllow)
        SetAllowCustomsCosts = S_OK
    End Function

    Public Function SetNotes(ByRef strNotes As String) As Short
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetNotes, strNotes)
        m_objDB.SetNotes(strNotes)
        SetNotes = S_OK
    End Function
    Public Function SetFundByUse(ByRef fFundByUse As Boolean) As Object
        m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objDB.GetFundByUse, fFundByUse)
        m_objDB.setFundByUse(fFundByUse)
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

    Public Function GetAggregationID() As String
        GetAggregationID = m_objDB.GetAggregationID()
    End Function

    Public Function GetName() As String
        GetName = m_objDB.GetName()
    End Function

    Public Function GetAbbreviation() As String
        GetAbbreviation = m_objDB.GetAbbreviation()
    End Function

    Public Function GetValue() As Double
        GetValue = m_objDB.GetValue()
    End Function

    Public Function getAllowKitCosts() As Boolean
        getAllowKitCosts = m_objDB.GetAllowKitCosts()
    End Function

    Public Function getAllowStorageCosts() As Boolean
        getAllowStorageCosts = m_objDB.GetAllowStorageCosts()
    End Function

    Public Function getAllowCustomsCosts() As Boolean
        getAllowCustomsCosts = m_objDB.GetAllowCustomsCosts()
    End Function

    Public Function GetNotes() As String
        GetNotes = m_objDB.GetNotes()
    End Function

    Public Function GetFundByUse() As Boolean
        GetFundByUse = m_objDB.GetFundByUse
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
        'UPGRADE_NOTE: Object colUseFundings may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        colUseFundings = Nothing
        'UPGRADE_NOTE: Object m_objCurrentUseFunding may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objCurrentUseFunding = Nothing

    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub
End Class