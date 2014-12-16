Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Aggregation_NET.Aggregation")> Public Class Aggregation
	'Aggregation.cls
	'
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	
	'Cost Calculator
	Dim m_objFundCalc As New FundingCalculator
	
	'connection to db
    Private m_objAggregationDB As ProQdb.AggregationDB
	
    Private m_strGuidID As String 'local copy
    Private m_strName As String 'local copy
	Private m_strPreparedBy As String 'local copy
	Private m_strNotes As String 'local copy
	Private m_dtmCreated As Date 'local copy
	Private m_strModifiedBy As String 'local copy
	Private m_dtmModified As Date 'local copy
	Private m_strCurrentQuantification As String 'local copy
	
	Public m_objCurrentQuantification As New Quantification 'local copy
	Public m_objCurrentFundingSource As New FundingSource
	
	Public m_objCountry As New Country
	Public m_objProgram As New Program
	
	Public m_colQuantifications As New Quantifications
	Public colFundingSources As New FundingSources
	
	'local variable(s) to hold property value(s)
	Private m_fNewRecord As Boolean 'local copy
	Private m_fIsDirty As Boolean 'Is the record Dirty or Not
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	Public Sub Create()
		'Create the and call the guidgen to create.
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUID.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_strGuidID = g_objGUID.GetGUID()
        m_dtmCreated = Now
		m_fNewRecord = True
		m_fIsDirty = True
		
        AddQuantification()

        '/ JSL - 24SEP2008 - Load the empty funding source, to stop the Aggregate form from failing.
        LoadFundingSources()
	End Sub
	
	
	
	Public Sub Delete(Optional ByRef fAsk As Boolean = False)
		' Call the DELETE Function from the DB object to get this one
		
		If fAsk = True Then
			If MsgBox("Are you sure you want to delete the selected aggregation? ", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "DELETE?") = MsgBoxResult.No Then
				Exit Sub
			End If
		End If
		
		Dim clsdbAgg As ProQdb.AggregationDB
		clsdbAgg = New ProQdb.AggregationDB
		
		If m_fNewRecord = False Then
			clsdbAgg.Delete(m_strGuidID)
		End If
		
		'UPGRADE_NOTE: Object clsdbAgg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		clsdbAgg = Nothing
		
	End Sub
	
	'+
	'Load()
	'
	'loads the indicated aggregation and all of its children (the
	'children being the quantification and the funding sources
	'
	'jleiner
	'11 june 2002
	'-
	Public Function Load(ByRef strID As String) As Short
		
		Dim clsdbAgg As ProQdb.AggregationDB
        clsdbAgg = New ProQdb.AggregationDB

		With clsdbAgg
			.Load(strID)
            m_strGuidID = .GetID()
            m_strName = .GetName()
			m_strNotes = .GetNotes()
			m_strPreparedBy = .GetCreator()
			m_dtmCreated = CDate(.GetCreationDate())
			m_dtmModified = CDate(.GetModifiedDate())
			m_strModifiedBy = .GetModifiedBy()
			
			'Set the Properties that are Objects
            m_objCountry.Load(.GetCountry())
            m_objProgram.Load(.GetProgram())
			
        End With

        'Load the Quantifications too
		LoadQuantifications()
		
		'Load the Funding Sources too
		LoadFundingSources()
		
		Load = S_OK
		
	End Function
	
	
	Public Function Save() As Short
		
		' Save the Quantifications First
		Dim objQ As Quantification
        Dim objFS As New FundingSource
		
		Dim clsdbAgg As New ProQdb.AggregationDB
		'Set clsdbAgg = New ProQdb.AggregationDB
		
		With clsdbAgg
			
			If m_fNewRecord = True Then
				.SetID(m_strGuidID)
				.SetName(m_strName)
				.SetCountry(m_objCountry.GetID())
				.SetProgram(m_objProgram.GetID())
				.SetNotes(m_strNotes)
				.SetCreator(m_strPreparedBy)
				.SetCreationDate(m_dtmCreated)
				
				'Add The New Record
				.Create()
				
			Else
				'Update Record
				clsdbAgg.Load(m_strGuidID)
				.SetID(m_strGuidID)
				.SetName(m_strName)
				.SetCountry(m_objCountry.GetID())
				.SetProgram(m_objProgram.GetID())
				.SetNotes(m_strNotes)
				.SetCreator(m_strPreparedBy)
				.SetCreationDate(m_dtmCreated)
				
				'Save the Record
				.Update()
			End If
			
		End With
		
		m_fNewRecord = False
		m_fIsDirty = False
		
		' Save the Aggregations
		For	Each objQ In m_colQuantifications
			objQ.Save()
		Next objQ
		
        ' Save the FundingSources
        If Not colFundingSources Is Nothing Then
            For Each objFS In colFundingSources
                objFS.Save()
            Next objFS
        End If

        Save = S_OK

    End Function
	' Create a new use (Quantification) for the Aggregate
	Function AddQuantification() As Short
		
		m_objCurrentQuantification = m_colQuantifications.Add()
		m_objCurrentQuantification.SetAggregationID(m_strGuidID)
		
		'Prompt User for Configuration and redo tree.
		
		AddQuantification = S_OK
		
	End Function
	
	
	' Delete a new use (Quantification) from the Aggregate
	Function DeleteQuantification() As Short
		
		m_colQuantifications.Remove(GetCurrentQuantificationID)
		m_objCurrentQuantification.Delete()
		
	End Function
	
	
	'---------------------------------------------------------
	'LoadQuantifications()
	' Comments    : Find all the Quantifications that exist
	'             : for the loaded Aggregate and Load them
	'             : into Array(?) collection?
	' Parameters  : None
	' Returns     : Integer S_OK
	' Created     : 29-May-02 jleiner
	'---------------------------------------------------------
	Function LoadQuantifications() As Short
		
        Dim objCol As New ProQdb.QuantificationDBCollection
        Dim objdb As New ProQdb.QuantificationDB
		
        Dim mCol As New Collection

		Dim strStoredProc As String
        Dim aParams(1, 1) As Object

		strStoredProc = SP_GET_QUANTIFICATIONS_BY_AGGREGATION
		
        aParams(0, 0) = New Guid(m_strGuidID)
        aParams(0, 1) = DbType.Guid 'DbType.Guid

        mCol = objCol.GetCollection(strStoredProc, aParams)

        For Each objdb In mCol
            'or i = 1 To mCol.Count
            m_objCurrentQuantification = m_colQuantifications.Add(objdb.GetID)
            m_objCurrentQuantification.Load(objdb.GetID)
        Next

        objCol = Nothing
        mCol = Nothing


    End Function


    '---------------------------------------------------------
    'LoadFundingSources()
    ' Comments    : Find all the FundingSources that exist
    '             : for the loaded Funding Source and Load them
    '             : into collection FundingSources
    ' Parameters  : None
    ' Returns     : Integer S_OK
    ' Created     : 07-June-02 jleiner
    '---------------------------------------------------------
    Function LoadFundingSources() As Short

        Dim objCol As New ProQdb.FundingSourceDBCollection
        Dim objFundingSourceDB As New ProQdb.FundingSourceDB

        Dim mCol As New Collection

        Dim strStoredProc As String
        Dim aParams(1, 1) As Object

        strStoredProc = SP_FUNDINGSOURCE_BY_AGGREGATION

        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(m_strGuidID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid

        mCol = objCol.GetCollection(strStoredProc, aParams)
        colFundingSources = New FundingSources

        For Each objFundingSourceDB In mCol
            m_objCurrentFundingSource = colFundingSources.Add(objFundingSourceDB.GetID)
            m_objCurrentFundingSource.Load(objFundingSourceDB.GetID)
        Next 'objFundingSourceDB

        'UPGRADE_NOTE: Object objCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objCol = Nothing
        'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mCol = Nothing

    End Function
	
	
	
	' Create a new use (FundingSource) for the Aggregate
	Public Function AddFundingSource() As Short
		
		m_objCurrentFundingSource = colFundingSources.Add()
		m_objCurrentFundingSource.SetAggregationID(m_strGuidID)
		
		'Prompt User for Configuration and redo tree.
		AddFundingSource = S_OK
		
	End Function
	
	
	' Delete a new use (FundingSource) from the Aggregate
	Public Function DeleteFundingSource() As Short
		
		colFundingSources.Remove(m_objCurrentFundingSource.GetID)
		m_objCurrentFundingSource.Delete()
		
		'UPGRADE_NOTE: Object m_objCurrentFundingSource may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objCurrentFundingSource = Nothing
		
	End Function
	
	
	Public Function GetCurrentFundingSource() As FundingSource
		GetCurrentFundingSource = m_objCurrentFundingSource
	End Function
	
	
	'SetCurrentFundingSource m_objCurrentFundingSource
	Public Function SetCurrentFundingSource(ByRef objFundingSource As FundingSource) As Short
		m_objCurrentFundingSource = objFundingSource
		SetCurrentFundingSource = S_OK
	End Function
	Public Function SetCurrentFundingSourcebyID(ByRef strID As String) As Short
		m_objCurrentFundingSource = colFundingSources(strID)
		SetCurrentFundingSourcebyID = S_OK
	End Function
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'accessors & manipulators
	
	'SetName(value)
	Public Function SetName(ByRef strValue As String) As Short
		
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_strName, strValue)
		
		'set the member
		m_strName = strValue
		
		'set the return code
		SetName = S_OK
	End Function
	
	'GetName()
	Public Function GetName() As String
		GetName = m_strName
	End Function
	
	'SetPreparedBy(value)
	Public Function SetPreparedBy(ByRef strValue As String) As Short
		
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_strPreparedBy, strValue)
		
		'set the member
		m_strPreparedBy = strValue
		'set the return code
		SetPreparedBy = S_OK
	End Function
	
	'GetPreparedBy()
	Public Function GetPreparedBy() As String
		GetPreparedBy = m_strPreparedBy
	End Function
	
	'GetdtmCreated()
	Public Function GetDateCreated() As Date
		GetDateCreated = m_dtmCreated
	End Function
	
	Public Function SetDateCreated(ByRef dtmValue As Date) As Short
		
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_dtmCreated, dtmValue)
		m_dtmCreated = dtmValue
		SetDateCreated = S_OK
	End Function
	
	'SetCountry(value)
	Public Function SetCountry(ByRef strValue As String) As Short
		
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objCountry.GetID, strValue)
		
		'set the member
		m_objCountry.Load(strValue)
		'set the return code
		SetCountry = S_OK
	End Function
	
	'GetCountryID()
	Public Function GetCountryID() As String
		GetCountryID = m_objCountry.GetID()
	End Function
	
	Public Function GetCountryName() As String
		GetCountryName = m_objCountry.GetName()
	End Function
	
	'SetProgram(value)
	Public Function SetProgram(ByRef strValue As String) As Short
		
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objProgram.GetID, strValue)
		
		'set the member
		m_objProgram.Load(strValue)
		
		'set the return code
		SetProgram = S_OK
	End Function
	
	'GetProgramID()
	Public Function GetProgramID() As String
		GetProgramID = m_objProgram.GetID()
	End Function
	Public Function GetProgramName() As String
		GetProgramName = m_objProgram.GetName()
	End Function
	
	
	'SetNotes(value)
	Public Function SetNotes(ByRef strValue As String) As Short
		
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_strNotes, strValue)
		
		'set the member
		m_strNotes = strValue
		'set the return code
		SetNotes = S_OK
	End Function
	
	'GetNotes()
	Public Function GetNotes() As String
		GetNotes = m_strNotes
	End Function
	
	'SetModifiedBy(value)
	Public Function SetModifiedBy(ByRef strValue As String) As Short
		
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_strModifiedBy, strValue)
		
		'set the member
		m_strModifiedBy = strValue
		'set the return code
		SetModifiedBy = S_OK
	End Function
	
	'GetModifiedBy()
	Public Function GetModifiedBy() As String
		GetModifiedBy = m_strModifiedBy
	End Function
	Public Function GetDateModified() As Date
		GetDateModified = m_dtmModified
	End Function
	
	'GetCurrentQuantification m_strCurrentQuantification
	Public Function GetCurrentQuantificationID() As String
		'    GetCurrentQuantificationID = m_strCurrentQuantification
		GetCurrentQuantificationID = m_objCurrentQuantification.GetID()
	End Function
	
	Public Function GetCurrentQuantification() As Quantification
		GetCurrentQuantification = m_objCurrentQuantification
	End Function
	
	'SetCurrentQuantification m_objCurrentQuantification
	Public Function SetCurrentQuantification(ByRef objQuantification As Quantification) As Short
		m_objCurrentQuantification = objQuantification
		m_strCurrentQuantification = objQuantification.GetID
		SetCurrentQuantification = S_OK
	End Function
	
	'SetCurrentQuantification m_strCurrentQuantification
	Public Function SetCurrentQuantificationbyID(ByRef strID As String) As Short
		m_objCurrentQuantification = m_colQuantifications(strID)
		m_strCurrentQuantification = strID
		SetCurrentQuantificationbyID = S_OK
	End Function
	'SetID(value)
	Public Function SetID(ByRef strValue As String) As Short
		'set the member
		m_strGuidID = strValue
		'set the return code
		SetID = S_OK
	End Function
	
	'getID()
	Public Function GetID() As String
		GetID = m_strGuidID
	End Function
	
	'SetIsDirty(value)
	Public Function SetIsDirty(ByRef fValue As Boolean) As Short
		'set the member
		m_fIsDirty = fValue
		'set the return code
		SetIsDirty = S_OK
	End Function
	
	'getIsDirty()
	Public Function GetIsDirty() As Boolean
		GetIsDirty = m_fIsDirty
	End Function
	
	'GetIsNew()
	Public Function GetIsNew() As Boolean
		GetIsNew = m_fNewRecord
	End Function
	
	'Calculate()
	
	'PrepareReport(id)
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'UPGRADE_NOTE: Object colFundingSources may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		colFundingSources = Nothing
		m_fNewRecord = False
		m_fIsDirty = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object m_objAggregationDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objAggregationDB = Nothing
		'UPGRADE_NOTE: Object m_objCurrentQuantification may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objCurrentQuantification = Nothing
		'UPGRADE_NOTE: Object m_objCurrentFundingSource may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objCurrentFundingSource = Nothing
		'UPGRADE_NOTE: Object m_objCountry may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objCountry = Nothing
		'UPGRADE_NOTE: Object m_objProgram may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objProgram = Nothing
		'UPGRADE_NOTE: Object m_colQuantifications may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_colQuantifications = Nothing
		'UPGRADE_NOTE: Object colFundingSources may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		colFundingSources = Nothing
		'UPGRADE_NOTE: Object m_objFundCalc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objFundCalc = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Funding Source ALLOCATED and AVAILABLE
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Function GetAllocated_FundingSource(ByRef objFS As FundingSource) As Double
		
		GetAllocated_FundingSource = m_objFundCalc.GetAllocated_FundingSource(Me, objFS)
		
	End Function
	
	Function GetAvailable_FundingSource(ByRef objFS As FundingSource) As Double
		
		GetAvailable_FundingSource = objFS.GetValue - GetAllocated_FundingSource(objFS)
		
	End Function
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' ALLOCATED and AVAILABLE Amounts for a Use Funding
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Function GetAllocated_UseFunding(ByRef objUF As UseFunding) As Double
		
		GetAllocated_UseFunding = m_objFundCalc.GetAllocated_UseFunding(Me, objUF)
		
	End Function
	
	Function GetAvailable_UseFunding(ByRef objUF As UseFunding) As Double
		
		GetAvailable_UseFunding = objUF.GetValue - GetAllocated_UseFunding(objUF)
		
	End Function
	
	Sub UndoFundingSources(ByRef objFS As FundingSource)
		' Comments  : Undo changes to a Funding Source Object
		' Parameters: objFS the funding Source in Question
		' Created   : 10-SEP-2002 JSL
		'---------------------------------------------------------------------
		
		Dim strID As String
		
		strID = objFS.GetID
		
		'UPGRADE_NOTE: Object Me.colFundingSources may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Me.colFundingSources = Nothing
		LoadFundingSources()
		
		m_objCurrentFundingSource = colFundingSources(strID)
		
		
		
	End Sub
End Class