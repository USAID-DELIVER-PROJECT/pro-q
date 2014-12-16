Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Quantification_NET.Quantification")> Public Class Quantification
	'+
	'Quantification.cls
	'
	'Represents one of the uses within an aggregation.  Besides all of
	'the normal properties, uses have a "protocol" and a list of
	'"methodologies".  Uses are the same as quantifications.
	'
	'lbailey
	'26 may 2002
	'-
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                           constants
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'Private Const SP_COSTS_BY_QUANTIFICATION = "qselCostsbyQuantification"
	'Private Const SP_QUANTIFICATION_METHODOLOGIES = "qselQuantificationMethodology"
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                        private properties
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	Private m_strGuidID As String
	Private m_strUse As String 'use
	Private m_strProtocolID As String 'StrProtocol
	Private m_strAggregationID As String 'StrAggregation
	Private m_dtmCreated As Date
	Private m_dtmModified As Date
	Private m_lngKitsToOrderCategory As Integer
	Private m_strNotes As String 'StrNotes
	Private m_fUseAverageMethod As Boolean
	Private m_sngDiscordancy As Single
	Private m_sngPrevalence As Single
	
	Public m_objUse As New Use
	
	Public Enum Cost_Type
		ctKits = 1
		ctCustoms = 2
		ctStorage = 3
	End Enum
	
	Public Enum AdjDemand_Categories
		acWASTAGE = 1
		acQC = 2
		acBoth = 3
	End Enum
	
	'protocol
	Private m_objProtocol As New Protocol
	
	'date
	Private m_dDate As Date
	
	Private m_fNewRecord As Boolean
	Private m_fIsDirty As Boolean
	
	Public colMethodologies As QuantMethods
	Public m_objCurrentMethod As QuantMethod
	Public m_objSelectedMethod As QuantMethod
	
	Private m_objDemand As New Section
	Private m_objAdjustedDemand As New Section
	Private m_objRequirements As New Section
	Private m_objCosts As New Section
	Private m_objFunds As New Section
	
	Private m_objCurrentCost As Cost 'Holder for the Current Cost Object
	Public m_objCostsKit As New Cost
	Public m_objCostsCustoms As New Cost
	Public m_objCostsStorage As New Cost
	
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public properties
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'none
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                          private methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'constructor
	'
	'by default, the class should be initialized such that it's pointing
	'to the top(head) of the script.
	'
	'lbailey
	'17 may 2002
	'
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_fIsDirty = False
		m_fNewRecord = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'destructor
	'
	'clean up
	'
	'lbailey
	'17 may 2002
	'
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		On Error Resume Next
		'UPGRADE_NOTE: Object m_objDemand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objDemand = Nothing
		'UPGRADE_NOTE: Object m_objAdjustedDemand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objAdjustedDemand = Nothing
		'UPGRADE_NOTE: Object m_objRequirements may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objRequirements = Nothing
		'UPGRADE_NOTE: Object m_objCosts may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objCosts = Nothing
		'UPGRADE_NOTE: Object m_objFunds may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objFunds = Nothing
		
		'UPGRADE_NOTE: Object m_objCostsKit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objCostsKit = Nothing
		'UPGRADE_NOTE: Object m_objCostsCustoms may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objCostsCustoms = Nothing
		'UPGRADE_NOTE: Object m_objCostsStorage may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objCostsStorage = Nothing
		'UPGRADE_NOTE: Object m_objCurrentCost may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objCurrentCost = Nothing
		
		'UPGRADE_NOTE: Object m_objUse may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objUse = Nothing
		'UPGRADE_NOTE: Object m_objProtocol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objProtocol = Nothing
		
		'UPGRADE_NOTE: Object colMethodologies may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		colMethodologies = Nothing
		'UPGRADE_NOTE: Object m_objCurrentMethod may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objCurrentMethod = Nothing
		'UPGRADE_NOTE: Object m_objSelectedMethod may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objSelectedMethod = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'Create
	Public Sub Create(ByRef strAggID As String)
		
		'ensure that any existing protocol is discarded
		'Set m_objProtocol = Nothing
		'Set m_objProtocol = New Protocol
		
		'Create the and call the guidgen to create.
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUID.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_strGuidID = g_objGUID.GetGUID()
		m_strAggregationID = strAggID
		
		m_sngDiscordancy = PCT_DISCORDANCE
		m_sngPrevalence = PCT_AIDS_PREVALENCE
		m_lngKitsToOrderCategory = 1
		m_dtmCreated = Now
		m_fNewRecord = True
		m_fIsDirty = True
		
		'Create The methodologies Collection
		LoadMethodologies()
		
		'Create The Cost Objects
		CreateCosts()
		
	End Sub
	
	
	'Delete
	Public Sub Delete(Optional ByRef fAsk As Boolean = False)
		
		' Call the DELETE Function from the DB object to get this one
		Dim objQuantification As New ProQdb.QuantificationDB
		
		If fAsk = True Then
			If MsgBox("Are you sure you want to delete the selected quantification? ", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "DELETE?") = MsgBoxResult.No Then
				Exit Sub
			End If
		End If
		
		If m_fNewRecord = False Then
			'Delete the Protocol First then the Quantification
			GetProtocol.Delete()
			objQuantification.Delete(m_strGuidID)
		End If
		
		'UPGRADE_NOTE: Object objQuantification may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objQuantification = Nothing
		
	End Sub
	
	
	
	'Load
	Public Function Load(ByRef strID As String) As Short
		
		Dim objQuantification As New ProQdb.QuantificationDB
		
		With objQuantification
			
			
			.Load(strID)
			m_strGuidID = .GetID()
			m_strAggregationID = .GetAggregationID()
			m_strNotes = .GetNotes()
			m_dtmCreated = CDate(.GetCreationDate())
			m_dtmModified = CDate(.GetModifiedDate())
			m_lngKitsToOrderCategory = .GetlngKitsToOrderCategory()
			m_fUseAverageMethod = .getUseAverageMethod()
			m_sngDiscordancy = .getDiscordancy()
			m_sngPrevalence = .getPrevalence()
			
			'Set the Properties that are Objects
			m_objUse.Load(.GetUseID)
			m_strProtocolID = .GetProtocolID
			
		End With
		
		m_fIsDirty = False
		m_fNewRecord = False
		
		'Load Methodologies chosen for the Quantification
		LoadMethodologies()
		
		'Load the 3 Cost Objects (Kits,Customs,Storage)
		LoadCosts()
		
		'load the one and only protocol
		LoadProtocol()
		
		Load = S_OK
		
		'UPGRADE_NOTE: Object objQuantification may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objQuantification = Nothing
		
	End Function
	
	Public Function LoadProtocol() As Integer
		m_objProtocol.SetQuantificationID(m_strGuidID)
		m_objProtocol.LoadByID(m_strProtocolID)
	End Function
	
	Public Function GetProtocol() As Protocol
		GetProtocol = m_objProtocol
	End Function
	
	Public Sub SetProtocol(ByRef objProtocol As Protocol)
		m_objProtocol = objProtocol
	End Sub
	
	Public Function Save() As Short
		
		Dim objQuantification As New ProQdb.QuantificationDB
		Dim objQM As QuantMethod
		
		With objQuantification
			If m_fNewRecord = True Then
				.SetID(m_strGuidID)
				.SetAggregationID(m_strAggregationID)
				.SetNotes(m_strNotes)
				.SetCreationDate(m_dtmCreated)
				.SetModifiedDate(Now)
				.SetlngKitsToOrderCategory(m_lngKitsToOrderCategory)
				.SetProtocolID(m_strProtocolID)
				.SetUseID(m_objUse.GetID)
				.SetUseAverageMethod(m_fUseAverageMethod)
				.SetDiscordancy(m_sngDiscordancy)
				.SetPrevalence(m_sngPrevalence)
				'Add The New Record
				.Create()
				
			ElseIf m_fIsDirty = True Then 
				'Update Record
				.Load(m_strGuidID)
				.SetID(m_strGuidID)
				.SetAggregationID(m_strAggregationID)
				.SetNotes(m_strNotes)
				.SetCreationDate(m_dtmCreated)
				.SetModifiedDate(Now)
				.SetlngKitsToOrderCategory(m_lngKitsToOrderCategory)
				.SetProtocolID(m_strProtocolID)
				.SetUseID(m_objUse.GetID)
				.SetUseAverageMethod(m_fUseAverageMethod)
				.SetDiscordancy(m_sngDiscordancy)
				.SetPrevalence(m_sngPrevalence)
				
				'Save the Record
				.Update()
			End If
			
		End With
		
		' Reset the variables
		'---------------------------------
		m_fNewRecord = False
		m_fIsDirty = False
		
		' Save the Methodologies
		'---------------------------------
		For	Each objQM In colMethodologies
			objQM.Save()
		Next objQM
		
		' Save the Cost Objects
		'---------------------------------
		m_objCostsKit.Save()
		m_objCostsCustoms.Save()
		m_objCostsStorage.Save()
		
		' Save the Protocol
		'---------------------------------
		m_objProtocol.Save()
		
		Save = S_OK
		
		'UPGRADE_NOTE: Object objQuantification may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objQuantification = Nothing
		
	End Function
	
	'---------------------------------------------------------
	'LoadMethodologies()
	' Comments    : Find all the Quantifications that exist
	'             : for the loaded Aggregate and Load them
	'             : into Array(?) collection?
	' Parameters  : None
	' Returns     : Integer S_OK
	' Created     : 29-May-02 jleiner
	'---------------------------------------------------------
	Public Function LoadMethodologies() As Short
		
		Dim objQuantMethodDB As New ProQdb.QuantMethodDB
		Dim objQuantMethod As New QuantMethod
		
        Dim rst As DataSet 'ADODB.Recordset()
		Dim strStoredProc As String
        Dim aParams(1, 1) As Object
        Dim i As Integer

		colMethodologies = New QuantMethods
		
		strStoredProc = SP_QUANTIFICATION_METHODOLOGIES
		
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(m_strGuidID)
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid 'DbType.Guid

        rst = objQuantMethodDB.ReturnRS(strStoredProc, aParams)

        'For	Each objDB In mCol
        For i = 0 To rst.Tables(strStoredProc).Rows.Count - 1
            objQuantMethod = colMethodologies.Add(CType((rst.Tables(strStoredProc).Rows(i).Item("GuidID")), Guid).ToString()) 'colMethodologies.Add(.Fields("guidID").Value)
            objQuantMethod.Load(CType((rst.Tables(strStoredProc).Rows(i).Item("GuidID")), Guid).ToString()) '.Fields("guidID").Value)
            If objQuantMethod.GetSelectedMethod = True Then
                SetSelectedMethod(objQuantMethod.GetID)
            End If
        Next
        'UPGRADE_NOTE: Object objQuantMethodDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objQuantMethodDB = Nothing

    End Function


    'AddMethodology()
    Public Function AddMethodology() As Short

        m_objCurrentMethod = colMethodologies.Add()
        m_objCurrentMethod.SetQuantificationID(m_strGuidID)

        'Prompt User for Configuration and redo tree.

        AddMethodology = S_OK

    End Function


    'RemoveMethodology()
    Public Function DeleteMethodology() As Short

        colMethodologies.Remove(m_objCurrentMethod.GetID)
        m_objCurrentMethod.Delete()

        DeleteMethodology = S_OK

    End Function



    '---------------------------------------------------------
    'LoadCosts()
    ' Comments    : Load the 3 cost records for the Quantification
    ' Parameters  : None
    ' Returns     : Integer S_OK
    ' Created     : 07-June-02 jleiner
    '---------------------------------------------------------
    Public Function LoadCosts() As Short

        Dim objCol As New ProQdb.CostDBCollection
        Dim objCostDB As New ProQdb.CostDB

        Dim mCol As Collection
        'Dim i As Integer
        'Dim sGuid As String

        Dim strStoredProc As String
        Dim aParams(1, 1) As Object

        strStoredProc = SP_COSTS_BY_QUANTIFICATION

        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(m_strGuidID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid

        mCol = objCol.GetCollection(strStoredProc, aParams)
        For Each objCostDB In mCol
            Select Case objCostDB.GetCategory
                Case TYPE_KITS
                    m_objCostsKit.Load((objCostDB.GetID))
                Case TYPE_CUSTOMS
                    m_objCostsCustoms.Load((objCostDB.GetID))
                Case TYPE_STORAGE
                    m_objCostsStorage.Load((objCostDB.GetID))
            End Select
        Next objCostDB

        SetCurrentCost(TYPE_KITS)

        'UPGRADE_NOTE: Object objCostDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objCostDB = Nothing
        'UPGRADE_NOTE: Object objCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objCol = Nothing
        'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mCol = Nothing

    End Function
	
	Function CreateCosts() As Short
		' Comments  :  For a new quantification create the cost objects (kit,
		'           :  Customs, Storage).
		' Parameters:
		' Returns   : Succesful or not
		' Created   : 08-June-2002
		'---------------------------------------------------------------------
		
		' Kit Cost Object
		With m_objCostsKit
			.Create(m_strGuidID)
			.SetCategory(TYPE_KITS)
		End With
		
		' Customs Cost Object
		With m_objCostsCustoms
			.Create(m_strGuidID)
			.SetCategory(TYPE_CUSTOMS)
		End With
		
		' Storage Costs Object
		With m_objCostsStorage
			.Create(m_strGuidID)
			.SetCategory(TYPE_STORAGE)
		End With
		
		SetCurrentCost(TYPE_KITS)
		
		CreateCosts = S_OK
	End Function
	Public Function getCurrentCost() As Cost
		getCurrentCost = m_objCurrentCost
	End Function
	
	Public Sub SetCurrentCost(ByRef intType As Short)
		Select Case intType
			Case TYPE_KITS
				m_objCurrentCost = m_objCostsKit
			Case TYPE_CUSTOMS
				m_objCurrentCost = m_objCostsCustoms
			Case TYPE_STORAGE
				m_objCurrentCost = m_objCostsStorage
		End Select
	End Sub
	
	Public Sub UndoCost(ByRef intType As Short)
		' Comments  :  Undo changes to a cost object, by reloading cost object
		' Parameters:
        ' Created   : 05-SEP-2002 JSL
        ' Modified  : 21-AUG-2008 JSL - resfresh objects with = New Cost, to clear error
		'---------------------------------------------------------------------

        Dim strID As String

        Try
            Select Case intType
                Case TYPE_KITS
                    strID = m_objCostsKit.GetID
                    '/ jleiner Added 'New Object' Since instance was cleared by the nothing statement
                    m_objCostsKit = Nothing
                    m_objCostsKit = New Cost
                    m_objCostsKit.Load(strID)
                Case TYPE_CUSTOMS
                    strID = m_objCostsCustoms.GetID
                    '/ jleiner Added 'New Object' Since instance was cleared by the nothing statement
                    m_objCostsCustoms = Nothing
                    m_objCostsCustoms = New Cost
                    m_objCostsCustoms.Load(strID)
                Case TYPE_STORAGE
                    strID = m_objCostsStorage.GetID
                    '/ jleiner Added 'New Object' Since instance was cleared by the nothing statement
                    m_objCostsStorage = Nothing
                    m_objCostsStorage = New Cost
                    m_objCostsStorage.Load(strID)
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            Exit Sub
        End Try

    End Sub
	
	
	Function InitializeSections() As Short
		' Comments  :  For a quantification Initialize the sections (Demand, Adj Demand,
		'           :  Qty Required, Costs, Funds)
		' Parameters:
		' Returns   : Succesful or not
		' Created   : 08-June-2002
		'---------------------------------------------------------------------
		
		m_objDemand.SetSectionType(Globals_Renamed.QuantityCategory.Demand)
		m_objAdjustedDemand.SetSectionType(Globals_Renamed.QuantityCategory.AdjustedDemand)
		m_objRequirements.SetSectionType(Globals_Renamed.QuantityCategory.QuantityRequired)
		m_objCosts.SetSectionType(Globals_Renamed.QuantityCategory.Costs)
		m_objFunds.SetSectionType(Globals_Renamed.QuantityCategory.Funds)
		
		InitializeSections = S_OK
	End Function
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'accessors & manipulators
	
	
	'SetNotes()
	Public Function SetNotes(ByVal strValue As String) As Short
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_strNotes, strValue)
		m_strNotes = strValue
		SetNotes = S_OK
	End Function
	
	'GetNotes
	Public Function GetNotes() As String
		GetNotes = m_strNotes
	End Function
	
	'SetUse
	Public Function SetUse(ByVal strValue As String) As Short
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_objUse.GetID, strValue)
		m_objUse.Load(strValue)
		SetUse = S_OK
	End Function
	
	'GetUse
	Public Function GetUse() As Use
		GetUse = m_objUse
	End Function
	
	'SetKitstoOrderCategory
	Public Sub SetKitsToOrderCategory(ByRef lngOrderCategory As Integer)
		m_lngKitsToOrderCategory = lngOrderCategory
	End Sub
	
	'GetKitstoOrderCategogy
	Public Function getKitsToOrderCategory() As Integer
		getKitsToOrderCategory = m_lngKitsToOrderCategory
	End Function
	
	'SetProtocol
	Public Function SetProtocolID(ByVal strValue As String) As Short
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_strProtocolID, strValue)
		m_strProtocolID = strValue
	End Function
	
	'GetProtocol
	Public Function GetProtocolID() As String
		GetProtocolID = m_objProtocol.GetID
	End Function
	'SetCreationDate
	Public Sub SetCreationDate(ByVal dtmValue As Date)
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_dtmCreated, dtmValue)
		m_dtmCreated = dtmValue
	End Sub
	
	'GetCreationDate
	Public Function GetCreationDate() As Date
		GetCreationDate = m_dtmCreated
	End Function
	
	'SetAggregationID m_strAggregationID
	Public Function SetAggregationID(ByVal strValue As String) As Short
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_strAggregationID, strValue)
		m_strAggregationID = strValue
	End Function
	
	'GetAggregationID
	Public Function GetAggregationID() As String
		GetAggregationID = m_strAggregationID
	End Function
	
	'SetID m_strGuidID
	Public Function SetID(ByVal strValue As String) As Short
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_strGuidID, strValue)
		m_strGuidID = strValue
	End Function
	
	'GetID
	Public Function GetID() As String
		GetID = m_strGuidID
	End Function
	Public Function GetIsNew() As Boolean
		GetIsNew = m_fNewRecord
	End Function
	Public Sub SetDiscordancy(ByVal strValue As Single)
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_sngDiscordancy, strValue)
		m_sngDiscordancy = strValue
	End Sub
	Public Function getDiscordancy() As Single
		getDiscordancy = m_sngDiscordancy
	End Function
	Public Sub SetPrevalence(ByVal strValue As Single)
		m_fIsDirty = g_objEditor.UpdateDirtyFlag(m_fIsDirty, m_sngPrevalence, strValue)
		m_sngPrevalence = strValue
	End Sub
	Public Function getPrevalence() As Single
		getPrevalence = m_sngPrevalence
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
	Public Sub SetUseAverageMethod(ByRef fUse As Boolean)
		m_fUseAverageMethod = fUse
	End Sub
	Public Function getUseAverageMethod() As Boolean
		getUseAverageMethod = m_fUseAverageMethod
	End Function
	Public Function SetCurrentMethod(ByRef strID As String) As Short
		
		m_objCurrentMethod = colMethodologies(strID)
		SetCurrentMethod = S_OK
		
	End Function
	
	Public Function GetCurrentMethod() As QuantMethod
		
		GetCurrentMethod = m_objCurrentMethod
		
	End Function
	
	Public Function SetSelectedMethod(ByRef strID As Object) As Short
		
        Dim i As Integer 'objQuantMethod As QuantMethod
        For i = 1 To colMethodologies.Count
            colMethodologies.Item(i).SetSelectedMethod(False)
        Next
		
		colMethodologies(strID).SetSelectedMethod(True)
		m_objSelectedMethod = colMethodologies(strID)
		
		' Save the Methodologies
		'---------------------------------
		If GetIsNew = False Then
            For i = 1 To colMethodologies.Count  'Each objQuantMethod In colMethodologies
                colMethodologies.Item(i).Save() 'objQuantMethod.Save()
            Next
		End If
		
	End Function
	
	Public Function GetSelectedMethod() As QuantMethod
		
		If m_objSelectedMethod Is Nothing Then
			If colMethodologies.Count > 0 Then
				SetSelectedMethod(colMethodologies(1).GetID)
				'Set m_objSelectedMethod = colMethodologies(1)
			End If
		End If
		
		GetSelectedMethod = m_objSelectedMethod
		
	End Function
	
	
	Function getSection(ByRef qcSection As Integer) As Section
		
		Select Case qcSection
			Case Globals_Renamed.QuantityCategory.Demand
				getSection = m_objDemand
			Case Globals_Renamed.QuantityCategory.AdjustedDemand
				getSection = m_objAdjustedDemand
			Case Globals_Renamed.QuantityCategory.QuantityRequired
				getSection = m_objRequirements
			Case Globals_Renamed.QuantityCategory.Costs
				getSection = m_objCosts
			Case Globals_Renamed.QuantityCategory.Funds
				getSection = m_objFunds
		End Select
	End Function
	
	Function GetTotalQuantificationCost() As Double
		GetTotalQuantificationCost = m_objCostsKit.GetValue + GetCost_Amount(Cost_Type.ctStorage) + GetCost_Amount(Cost_Type.ctCustoms)
	End Function
	
	Function GetCost_Amount(ByRef intCostType As Cost_Type) As Double
		' Comments  : Returns the cost value for the Cost object
		' Parameters: intType - The cost object variable (TYPE_KITS,TYPE_STORAGE,TYPE_CUSTOMS)
		' Returns   : Double - the amount
		' Created   : 16-June-2002 jleiner
		'-------------------------------------------------------------------------------
		Dim curCost As Double
		
		Select Case intCostType
			Case TYPE_KITS
				'curCost = m_objCostsKit.GetValue
				curCost = CalculateKitCost()
				m_objCostsKit.SetValue(CDbl(curCost))
				
			Case TYPE_STORAGE
				Select Case m_objCostsStorage.GetMethod
					Case 1
						'% of Kit of cost
						'curCost = m_objCostsKit.GetValue * m_objCostsStorage.GetValue / 100
						curCost = CalculateKitCost() * m_objCostsStorage.GetValue / 100
					Case 2, 3, 4
						curCost = m_objCostsStorage.GetValue
				End Select
			Case TYPE_CUSTOMS
				Select Case m_objCostsCustoms.GetMethod
					Case 1
						'% of Kit of cost
						'curCost = m_objCostsKit.GetValue * m_objCostsCustoms.GetValue / 100
						curCost = CalculateKitCost() * m_objCostsCustoms.GetValue / 100
						
					Case 2, 3, 4
						curCost = m_objCostsCustoms.GetValue
				End Select
		End Select
		
		GetCost_Amount = curCost
		
	End Function
	
	Function GetCost_Unallocated(ByRef intType As Cost_Type) As Double
		' Comments  : Returns the unallocated amount for the specified Cost object
		' Parameters: intType - The cost object variable (TYPE_KITS,TYPE_STORAGE,TYPE_CUSTOMS)
		' Returns   : Double - the amount
		' Created   : 16-June-2002 jleiner
		'-------------------------------------------------------------------------------
		Dim curCost As Double
		
		Select Case intType
			Case TYPE_KITS
				curCost = GetCost_Amount(intType) - m_objCostsKit.GetAllocatedAmount
			Case TYPE_STORAGE
				curCost = GetCost_Amount(intType) - m_objCostsStorage.GetAllocatedAmount
			Case TYPE_CUSTOMS
				curCost = GetCost_Amount(intType) - m_objCostsCustoms.GetAllocatedAmount
		End Select
		
		GetCost_Unallocated = curCost
		
	End Function
	
	Function GetCost_AllocatedRatio(ByRef intType As Cost_Type) As Single
		' Comments  : Returns the unallocated amount for the specified Cost object
		' Parameters: intType - The cost object variable (TYPE_KITS,TYPE_STORAGE,TYPE_CUSTOMS)
		' Returns   : Double - the amount
		' Created   : 16-June-2002 jleiner
		'-------------------------------------------------------------------------------
		Dim curCost As Double
		Dim curAllocated As Double
		
		Select Case intType
			Case TYPE_KITS
				curCost = GetCost_Amount(intType)
				curAllocated = m_objCostsKit.GetAllocatedAmount
			Case TYPE_STORAGE
				curCost = GetCost_Amount(intType)
				curAllocated = m_objCostsStorage.GetAllocatedAmount
			Case TYPE_CUSTOMS
				curCost = GetCost_Amount(intType)
				curAllocated = m_objCostsCustoms.GetAllocatedAmount
		End Select
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NulltoValue(curCost, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If NulltoValue(curCost, 0) = 0 Then
			GetCost_AllocatedRatio = 0
		Else
            GetCost_AllocatedRatio = System.Math.Round(NulltoValue(curAllocated, 0) / NulltoValue(curCost, 0), 4)
		End If
	End Function
	
	
	Function GetDemandCollection(ByRef objQM As QuantMethod) As Collection
		' Comments  : Returns a Collection of DemandDB objects for the Use/Method
		' Parameters: ObjQM - the QuantMethod in Question
		' Returns   : Collection of DemandDB objects
		' Created   : 19-June-2002 jleiner
		'-------------------------------------------------------------------------------
		
		Dim strSQL As String
		Dim aValues(1) As Object
		'Dim mCol As Collection
		Dim objDBCol As New ProQdb.DemandDBCollection
		'Dim objDB As DemandDB
		
		'UPGRADE_WARNING: Couldn't resolve default property of object aValues(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		aValues(0) = m_strGuidID
		
		If Not (objQM Is Nothing) Then
			strSQL = g_objDM.GetNeedSQLString(m_objUse.GetID, objQM.GetMethodologyID)
			strSQL = g_objDM.ReplaceSQLString(strSQL, aValues)
			GetDemandCollection = objDBCol.GetCollection(strSQL)
		Else
			'UPGRADE_NOTE: Object GetDemandCollection may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			GetDemandCollection = Nothing
		End If
		
		'UPGRADE_NOTE: Object objDBCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objDBCol = Nothing
	End Function
	Function GetInitialDemand(ByRef objQM As QuantMethod, ByRef fDemand2 As Boolean) As Object
		
		If m_fUseAverageMethod = False Then
			If objQM Is Nothing Then
				If m_objSelectedMethod Is Nothing Then
					If Me.colMethodologies.Count > 0 Then
						objQM = colMethodologies(1)
					Else
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetInitialDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetInitialDemand = System.DBNull.Value
						Exit Function
					End If
				Else
					objQM = m_objSelectedMethod
				End If
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object GetInitialDemand_Method(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetInitialDemand = GetInitialDemand_Method(objQM, fDemand2)
		Else
			If Not (objQM Is Nothing) Then
				'Average but Method Specified
				'UPGRADE_WARNING: Couldn't resolve default property of object GetInitialDemand_Method(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetInitialDemand = GetInitialDemand_Method(objQM, fDemand2)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object GetAverageDemand(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetInitialDemand = GetAverageDemand(fDemand2)
			End If
		End If
		
	End Function
	
	
	Function GetInitialDemand_Method(ByRef objQM As QuantMethod, ByRef fDemand2 As Boolean) As Object
		' Comments  : Returns the Initial Demand Calc for the Methodology.
		'           : If the Methodology is 'Logisitcs' then return the sum
		'           : of the Demands for the 1st Level Tests.
		' Parameters: objQM - The QuantMethod in Question
		'           : fDemand2 - Return Demand(getResult) or Test2 demand (GetDemand2)
		' Returns   : Variant - if fDemand2 is false then return the Demand
		'           :           if fDemand2 is True then return the 2nd Level Demand
		' Created   : 19-June-2002 jleiner
		'-------------------------------------------------------------------------------
		
        Dim mCol As New Collection
		'Dim colTests As Collection
		'Dim objPT As ProtocolTest
		'Dim objPB As ProtocolBrand
		
		Dim objDemandDB As ProQdb.DemandDB
		Dim varDemand As Object
        varDemand = Nothing
        '
		mCol = GetDemandCollection(objQM)
		
		If mCol.Count() = 0 Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetInitialDemand_Method. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetInitialDemand_Method = System.DBNull.Value
			Exit Function
		End If
		
        If objQM.GetMethodologyID.Equals("3C97F284-1AA3-402C-A8A0-68393066E19D", StringComparison.InvariantCultureIgnoreCase) Then
            'Logistics Demand Loop through and get all the 1st Level Demands
            'Set colTests = m_objProtocol.GetLevel(1).GetTests
            'For Each objPT In colTests
            '    For Each objPB In objPT.GetBrands
            '        For Each objDemandDB In mCol
            '            Debug.Print objPB.GetBrand.GetID, objDemandDB.GetBrandID
            '            If objDemandDB.GetBrandID = objPB.GetBrand.GetID Then
            '                varDemand = varDemand + objDemandDB.getResult() * objPB.GetBrandPercent()
            '            End If
            '        Next
            '    Next
            'Next
            For Each objDemandDB In mCol
                If fDemand2 = False Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object objDemandDB.getResult(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If Not IsDBNull(objDemandDB.getResult()) Then
                        varDemand = varDemand + objDemandDB.getResult()
                    End If
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object objDemandDB.getDemand2(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ' 4-Aug-08 - jleiner - Keep Demand Value to be null for passing back to calling equation
                    If Not IsDBNull(objDemandDB.getDemand2()) Then
                        varDemand = varDemand + objDemandDB.getDemand2()
                    End If
                End If
            Next objDemandDB

            'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

            If Not IsNothing(varDemand) Then
                varDemand = System.Math.Round(NulltoValue(varDemand, 0) / NulltoValue(m_objProtocol.GetWeight(), 0), 0)
            Else
                varDemand = DBNull.Value
            End If

        Else
            'Return the Demand figure
            objDemandDB = mCol.Item(1)
            If fDemand2 = False Then
                'UPGRADE_WARNING: Couldn't resolve default property of object objDemandDB.getResult(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                varDemand = objDemandDB.getResult()
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object objDemandDB.getDemand2(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                varDemand = objDemandDB.getDemand2()
            End If
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object GetInitialDemand_Method. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetInitialDemand_Method = varDemand

    End Function
	
	Function GetInitialDemand_Brand(ByRef objQM As QuantMethod, ByRef objBrand As Brand) As Object
		' Comments  : Returns the Initial Demand Calc for the Methodology.
		'           : If the Methodology is 'Logisitcs' then return the sum
		'           : of the Demands for the 1st Level Tests.
		' Parameters: objQM - The QuantMethod in Question
		'           : fDemand2 - Return Demand(getResult) or Test2 demand (GetDemand2)
		' Returns   : Variant - if fDemand2 is false then return the Demand
		'           :           if fDemand2 is True then return the 2nd Level Demand
		' Created   : 19-June-2002 jleiner
		'-------------------------------------------------------------------------------
		
		Dim mCol As Collection
        'Dim colTests As Collection
		
		Dim objDemandDB As ProQdb.DemandDB
		Dim varDemand As Object
        varDemand = Nothing
		mCol = GetDemandCollection(objQM)
		
		If mCol.Count() = 0 Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetInitialDemand_Brand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetInitialDemand_Brand = System.DBNull.Value
			Exit Function
		End If
		
        If objQM.GetMethodologyID.Equals("3C97F284-1AA3-402C-A8A0-68393066E19D", StringComparison.InvariantCultureIgnoreCase) Then
            For Each objDemandDB In mCol
                If objDemandDB.getBrandID = objBrand.GetID Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object objDemandDB.getResult(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    varDemand = objDemandDB.getResult()
                    Exit For
                End If
            Next objDemandDB

            'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varDemand = varDemand / m_objProtocol.GetWeight()

        Else
            'Wrong Function
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varDemand = System.DBNull.Value
        End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetInitialDemand_Brand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetInitialDemand_Brand = varDemand
		
	End Function
	
	Function GetServiceCapacity(ByRef fHasLabTest As Boolean) As ProQdb.DemandDB
		' Comments  : Returns the the Service Capacity of the System.
		' Parameters: None
		' Returns   : Variant - The Service Capacity
		' Created   : 21-June-2002 jleiner
		'-------------------------------------------------------------------------------
		
		Dim mCol As Collection
		
		Dim objDemandDB As ProQdb.DemandDB
        'Dim varDemand As Object
		
		Dim strSQL As String
        Dim aValues(1) As String
		Dim objDBCol As New ProQdb.DemandDBCollection
		
		'UPGRADE_WARNING: Couldn't resolve default property of object aValues(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aValues(0) = m_strGuidID
		
		strSQL = g_objDM.GetFormulaSQLString(m_objUse.GetID, 2, fHasLabTest) 'Need test for elisa
		strSQL = g_objDM.ReplaceSQLString(strSQL, aValues)
		
		mCol = objDBCol.GetCollection(strSQL)
		'UPGRADE_NOTE: Object objDBCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objDBCol = Nothing
		
		If mCol.Count() = 0 Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetServiceCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GetServiceCapacity = Nothing
			Exit Function
		End If
		
		'Return the Demand figure
		objDemandDB = mCol.Item(1)
		'varDemand = objDemandDB.GetResult()
		
		GetServiceCapacity = objDemandDB
		'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mCol = Nothing
		
	End Function
	
	'Function GetDemand_SCFiltered() As Variant
	' Comments  : Returns the SC Filtered Demand for the Select Methodology. Which
	'           : is the Lower value of Demand and Service Capacity
	' Parameters: None
	' Returns   : Variant - The Service Capacity Filtered Demand
	' Created   : 21-June-2002 jleiner
	'-------------------------------------------------------------------------------
	'    Dim varInitialDemand As Variant
	'    Dim varServiceCapacity As Variant
	'
	'    varInitialDemand = GetInitialDemand(GetSelectedMethod, False)
	'    varServiceCapacity = GetServiceCapacity.getResult
	
	'    'Return Lessor of the 2 values
	'    If IsNull(varInitialDemand) Or IsNull(varServiceCapacity) Then
	'        GetDemand_SCFiltered = Null
	'    Else
	'        GetDemand_SCFiltered = IIf(varInitialDemand < varServiceCapacity, _
	''                    varInitialDemand, varServiceCapacity)
	'    End If
	'End Function
	
	'Function GetDemand_adjusted(objSBrand As SelectedBrand) As Variant
	'' Comments  : Returns the Adjusted Demand for the passed in Brand
	'' Parameters: objBrand the Brand to test
	''           :
	'' Returns   : Variant - The AdjustedDemand
	'' Created   : 21-June-2002 jleiner
	''-------------------------------------------------------------------------------'
	'
	'GetDemand_adjusted = Me.GetDemand_SCFiltered * _
	''            (1 + Me.GetDemand_WastageAndQC(objSBrand, acBoth))'
	'
	'End Function
	
	Function GetDemand_WastageAndQC(ByRef objSBrand As SelectedBrand, ByRef lngReturn As AdjDemand_Categories) As Object
		' Comments  : Returns the QC and Wastage for the passed in Brand.  Can return
		'           : one or Sum of both (as percents (in decimal form)
		' Parameters: objBrand the Brand to test
		'           : lngReturn a AdjDemand_Categories enum value 1 for Wastage, 2 for QC 3 for both (summed)
		' Returns   : Variant - The Service Capacity
		' Created   : 21-June-2002 jleiner
		'-------------------------------------------------------------------------------
		Dim mCol As Collection
		
		Dim objDemandDB As ProQdb.DemandDB
		Dim varValue As Object
		
		Dim strSQL As String
		Dim aValues(1) As Object
		Dim objDBCol As New ProQdb.DemandDBCollection
		
		'UPGRADE_WARNING: Couldn't resolve default property of object aValues(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		aValues(0) = m_strGuidID
		
		strSQL = g_objDM.GetFormulaSQLString(m_objUse.GetID, 3, False)
		strSQL = g_objDM.ReplaceSQLString(strSQL, aValues)
		
		mCol = objDBCol.GetCollection(strSQL)
		'UPGRADE_NOTE: Object objDBCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objDBCol = Nothing
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		varValue = System.DBNull.Value
		
		If mCol.Count() = 0 Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetDemand_WastageAndQC. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetDemand_WastageAndQC = System.DBNull.Value
			Exit Function
		End If
		
		For	Each objDemandDB In mCol
			If objSBrand.GetBrandID = objDemandDB.GetBrandID Then
				'Return the desired values
				Select Case lngReturn
					Case AdjDemand_Categories.acWASTAGE
						'UPGRADE_WARNING: Couldn't resolve default property of object objDemandDB.getDemand2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						varValue = objDemandDB.getDemand2
					Case AdjDemand_Categories.acQC
						'UPGRADE_WARNING: Couldn't resolve default property of object objDemandDB.getResult. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						varValue = objDemandDB.getResult
					Case AdjDemand_Categories.acBoth
						'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						varValue = objDemandDB.getResult + objDemandDB.getDemand2
				End Select
				Exit For
			End If
		Next objDemandDB
		
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetDemand_WastageAndQC. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetDemand_WastageAndQC = varValue
		'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mCol = Nothing
	End Function
	
	Function GenerateBrandQTYCalcs(Optional ByRef objDemandMethod As QuantMethod = Nothing) As BrandQtyCalculatorCol
		'+
		' Comments  : Returns a collection of BrandQtyCalculators (one for each selected
		'           : brand
		' Parameters: -
		' Returns   : Collection
		' Created   : 24-June-2002
		' Modified  : 07-May-2003 jleiner - Service Capacity needs to be checked for
		'           : Units (Samples vs. Tests) if in units, convert to Tests.
		' -
		'---------------------------------------------------------------
		
		Dim colBQC As New Collection
		Dim objBQC As New BrandQtyCalculator
		Dim objBQCcol As New BrandQtyCalculatorCol
		Dim colDemand As Collection 'The Collection containing the demand Figures
        'Dim objDemand As ProQdb.DemandDB 'The Demand Calculations
		Dim varInitialDemand As Object
		Dim varPrevRate As Object
		Dim varServiceCapacity_NoLab As Object
		Dim varServiceCapacity_Lab As Object
		Dim colQC As Collection
		Dim objQC As ProQdb.DemandDB
		Dim objServiceCapacity As ProQdb.DemandDB
		Dim objLogConsDB As New ProQdb.LogisticsConsDBCollection
		Dim colLogistics As Collection
		Dim objLog As ProQdb.LogisticsConsDB
		Dim colStorage As Collection
		
		Dim objStorageDB As New ProQdb.DemandDBCollection
		Dim objStorage As ProQdb.DemandDB
		Dim objSelBrand As SelectedBrand
		Dim objInitDemand As ProQdb.DemandDB
		
		Dim sngTypePercent As Single
		
		Dim strSQL As String
		Dim aValues(2) As Object
		Dim objDBCol As New ProQdb.DemandDBCollection
		Dim fIsLabBased As Boolean
		Dim fColdStorage As Boolean
		Dim strBrandID As String
		Dim fBrandFound As Boolean
		Dim fIsLogisticsMethod As Boolean
		
        'Dim i As Short
		
        'Dim varQC As Object
        'Dim varWastage As Object
		
		If objDemandMethod Is Nothing And m_fUseAverageMethod = False Then
			objDemandMethod = GetSelectedMethod
		End If
		
		If Not objDemandMethod Is Nothing Then
            fIsLogisticsMethod = (objDemandMethod.GetMethodologyID.Equals("3C97F284-1AA3-402C-A8A0-68393066E19D", StringComparison.InvariantCultureIgnoreCase))
		End If
		'------------------------------------------------------
		'Get the Initial Demand Figures for the collection
		'------------------------------------------------------
        varInitialDemand = GetInitialDemand(objDemandMethod, False)
        varPrevRate = GetInitialDemand(objDemandMethod, True)
		
		colDemand = GetDemandCollection(objDemandMethod)
		
		'------------------------------------------------------
		'Get the QC as Wastage figures for each Brand
        '------------------------------------------------------------
		aValues(0) = m_strGuidID
		strSQL = g_objDM.GetFormulaSQLString(m_objUse.GetID, 3, False)
		strSQL = g_objDM.ReplaceSQLString(strSQL, aValues)
		
		colQC = objDBCol.GetCollection(strSQL)
		
		
		'------------------------------------------------------
		'Update the SelectedBrands
		'------------------------------------------------------
		If GetProtocol.GetLevels.Count() >= 2 Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(varPrevRate) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object varPrevRate. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetProtocol.GetLevel(2).SetPercent(CSng(varPrevRate), getDiscordancy)
			Else
				GetProtocol.GetLevel(2).SetPercent(getPrevalence, getDiscordancy)
			End If
		End If
		
		If GetProtocol.GetLevels.Count() >= 3 Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(varPrevRate) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object varPrevRate. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetProtocol.GetLevel(3).SetPercent(CSng(varPrevRate), getDiscordancy)
			Else
				GetProtocol.GetLevel(3).SetPercent(getPrevalence, getDiscordancy)
			End If
		End If
		
		'GetProtocol.GetLevel(3).SetPercent
		'UPGRADE_WARNING: Couldn't resolve default property of object NulltoValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetProtocol.SetNumberSamples(CInt(NulltoValue(varInitialDemand, 0)))
		
		'------------------------------------------------------
		'Get the Service Capacity Values
		'------------------------------------------------------
		'Check for an Lab-based Test in Collection
		fIsLabBased = False
		For	Each objSelBrand In GetProtocol.GetSelectedBrands.GetSelectedBrands
			If objSelBrand.GetBrand.IsLab = True Then
				fIsLabBased = True
				Exit For
			End If
		Next objSelBrand
		
		objServiceCapacity = GetServiceCapacity(fIsLabBased)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NulltoValue(objServiceCapacity.getValue1, SC_SAMPLES). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If NulltoValue(objServiceCapacity.getValue1, SC_SAMPLES) = SC_SAMPLES Then
            'Values in Samples, Set into Tests.
            'UPGRADE_WARNING: Couldn't resolve default property of object varServiceCapacity_NoLab. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varServiceCapacity_NoLab = objServiceCapacity.getResult * GetProtocol.GetWeight
            'UPGRADE_WARNING: Couldn't resolve default property of object varServiceCapacity_Lab. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varServiceCapacity_Lab = objServiceCapacity.getDemand2 * GetProtocol.GetWeight
        Else
            'Values in Tests
            'UPGRADE_WARNING: Couldn't resolve default property of object objServiceCapacity.getResult. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object varServiceCapacity_NoLab. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varServiceCapacity_NoLab = objServiceCapacity.getResult
            'UPGRADE_WARNING: Couldn't resolve default property of object objServiceCapacity.getDemand2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object varServiceCapacity_Lab. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varServiceCapacity_Lab = objServiceCapacity.getDemand2
        End If

        '------------------------------------------------------
        'Get the Storage Capacity Numbers
        ' - Need to determine if Cold Storage is required or not
        '------------------------------------------------------------
        fColdStorage = False
        For Each objSelBrand In GetProtocol.GetSelectedBrands.GetSelectedBrands
            If objSelBrand.GetBrand.GetColdStorage = True Then
                fColdStorage = True
                Exit For
            End If
        Next objSelBrand

        'UPGRADE_WARNING: Couldn't resolve default property of object aValues(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aValues(0) = m_strGuidID
        strSQL = g_objDM.GetFormulaSQLString(m_objUse.GetID, 5, fColdStorage)
        strSQL = g_objDM.ReplaceSQLString(strSQL, aValues)

        colStorage = objStorageDB.GetCollection(strSQL)
        If colStorage.Count() > 0 Then
            objStorage = colStorage.Item(1)
            objBQCcol.SetColdStorage(objStorage.getDemand2)
            objBQCcol.SetStorageCapacity(objStorage.getResult)
        End If
        'UPGRADE_NOTE: Object objStorageDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objStorageDB = Nothing

        '----------------------------------------------------------------------------
        '----------------------------------------------------------------------------
        ' For Each Selected Brand in Protocol Create a BrandQTYCalc and
        ' Populate the known values
        '----------------------------------------------------------------------------
        '----------------------------------------------------------------------------
        For Each objSelBrand In GetProtocol.GetSelectedBrands.GetSelectedBrands

            'strBrandID = objSelBrand.GetBrandID
            strBrandID = objSelBrand.GetID

            '---------------------------------------------------------
            'Determine if the product is lab based or not
            '---------------------------------------------------------
            fIsLabBased = objSelBrand.GetBrand.IsLab

            'Set objBQC = New BrandQtyCalculator
            objBQC = objBQCcol.Add(strBrandID)
            objBQC.SetSelectedBrand(objSelBrand)

            'Set objQC = colQC.Item(strBrandID)

            '----------------------------------------------------------
            ' Find the Brand in the Quality Control object
            '----------------------------------------------------------
            fBrandFound = False
            For Each objQC In colQC
                If objQC.getBrandID = strBrandID Then
                    fBrandFound = True
                    Exit For
                End If
            Next objQC

            If fBrandFound = False Then
                GoTo Skip_brand
                'Brand Not found in collection
                'varQC = Null
                'varWastage = Null
            Else
                'varQC = objQC.getResult
                'varWastage = objQC.getDemand2
            End If

            '-----------------------------------------------------------
            'Get the Logistics Considerations figures for the brand
            '-----------------------------------------------------------
            'UPGRADE_WARNING: Couldn't resolve default property of object aValues(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            aValues(0) = m_strGuidID
            'UPGRADE_WARNING: Couldn't resolve default property of object aValues(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            aValues(1) = strBrandID
            strSQL = g_objDM.GetFormulaSQLString(m_objUse.GetID, 4, fIsLabBased)
            strSQL = g_objDM.ReplaceSQLString(strSQL, aValues)

            colLogistics = objLogConsDB.GetCollection(strSQL)

            'Set objLog = colLogistics.Item(strBrandID)

            For Each objLog In colLogistics
                If objLog.getBrandID = strBrandID Then
                    Exit For
                End If
            Next objLog

            '-----------------------------------------------------------
            ' Set the values into the brand Calculator
            '-----------------------------------------------------------
            objBQC.SetSelectedBrand(objSelBrand)

            ' Check for Logistics Demand or Not
            '-----------------------------------------------------------
            If fIsLogisticsMethod = False Then
                objBQC.SetDemand_Initial(objSelBrand.GetCount)
            Else
                For Each objInitDemand In colDemand
                    If objInitDemand.getBrandID = objSelBrand.GetID Then
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object NulltoValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        objBQC.SetDemand_Initial(IIf(NulltoValue(objInitDemand.getResult, 0) = 0, System.DBNull.Value, objInitDemand.getResult))
                        Exit For
                    End If
                Next objInitDemand
            End If

            ' Set the Service Capacity for the product
            '-----------------------------------------------------------
            If fIsLabBased = False Then
                'UPGRADE_WARNING: Couldn't resolve default property of object GetProtocol.GetSelectedBrands.GetTypeBrands()().GetPercent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sngTypePercent = GetProtocol.GetSelectedBrands.GetTypeBrands(False).Item(strBrandID).GetPercent
                If sngTypePercent = 0 Then
                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                    objBQC.SetServiceCapacity(System.DBNull.Value)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object varServiceCapacity_NoLab. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    objBQC.SetServiceCapacity(System.Math.Round(NulltoValue(varServiceCapacity_NoLab, 0) * sngTypePercent, 0))
                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object GetProtocol.GetSelectedBrands.GetTypeBrands()().GetPercent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sngTypePercent = GetProtocol.GetSelectedBrands.GetTypeBrands(True).Item(strBrandID).GetPercent
                If sngTypePercent = 0 Then
                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                    objBQC.SetServiceCapacity(System.DBNull.Value)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object varServiceCapacity_Lab. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    objBQC.SetServiceCapacity(System.Math.Round(NulltoValue(varServiceCapacity_Lab, 0) * sngTypePercent, 0))
                End If
            End If

            objBQC.SetQC(objQC.getResult)
            objBQC.SetWastage(objQC.getDemand2)
            objBQC.SetLeadTime(objLog.getLeadTime)
            objBQC.SetBufferStock(objLog.getBufferStock)
            objBQC.SetShipments(objLog.getShipments)

            '-------------------------------------------------------------------------------
            ' Storage Capacity:
            ' Get the volume of the product and the percent of the space required
            '-------------------------------------------------------------------------------
            'lblanken 7/9/08 - Set values to zero if null so calc will still work.
            If objSelBrand.GetBrand.GetColdStorage = True Then
                objBQC.SetStorageCapacity(System.Math.Round(NulltoValue(objStorage.getDemand2, 0) * sngTypePercent, 0))
                objBQC.SetColdStorage(System.Math.Round(NulltoValue(objStorage.getDemand2, 0) * sngTypePercent, 0))
            Else
                objBQC.SetStorageCapacity(System.Math.Round(NulltoValue(objStorage.getResult, 0) * sngTypePercent, 0))
                objBQC.SetColdStorage(System.Math.Round(NulltoValue(objStorage.getDemand2, 0) * sngTypePercent, 0))
            End If

            objBQC.SetQtyOnHand(objLog.getQtyOnHand)
            objBQC.SetQtyOnOrder(objLog.getQtyOnOrder)

            'colBQC.Add(objBQC, strBrandID)

Skip_brand:
            'Set objBQC = Nothing
        Next objSelBrand

        'Set GenerateBrandQTYCalcs = colBQC
        GenerateBrandQTYCalcs = objBQCcol

        'UPGRADE_NOTE: Object colBQC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        colBQC = Nothing
        'UPGRADE_NOTE: Object objDBCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objDBCol = Nothing
		
	End Function
	
	Function CalculateKitCost() As Double
		
		
		Dim objBQCcol As BrandQtyCalculatorCol
		Dim objBQC As BrandQtyCalculator
		Dim dblCost As Double
		
		
		objBQCcol = GenerateBrandQTYCalcs()
		
		For	Each objBQC In objBQCcol
			'UPGRADE_WARNING: Couldn't resolve default property of object NulltoValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dblCost = dblCost + NulltoValue(objBQC.getKitTotalCost, 0)
		Next objBQC
		
		CalculateKitCost = dblCost
		
		'UPGRADE_NOTE: Object objBQCcol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objBQCcol = Nothing
		'UPGRADE_NOTE: Object objBQC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objBQC = Nothing
		
	End Function
	
	Function GetAverageDemand(ByRef fDemand2 As Boolean) As Object
		' Comments  : Returns the Average of Initial Demand Calc for the Methodologies.
		' Parameters: fDemand2 - Return Demand(getResult) or Test2 demand (GetDemand2)
		' Returns   : Variant - if fDemand2 is false then return the Demand
		'           :           if fDemand2 is True then return the 2nd Level Demand
		' Created   : 14-Aug-2002 jleiner
		'-------------------------------------------------------------------------------
		
        'Dim objDemandDB As ProQdb.DemandDB
		Dim varDemand As Object
		Dim varDemandTemp As Object
		Dim intMethodCount As Short
		'Dim objMethodology As Methodology
		Dim objQM As QuantMethod
        varDemand = Nothing
		For	Each objQM In colMethodologies
            If objQM.GetMethodologyID <> "00000000-0000-0000-0000-000000000001" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object GetInitialDemand_Method(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object varDemandTemp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                varDemandTemp = GetInitialDemand_Method(objQM, fDemand2)
                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                If Not IsDBNull(varDemandTemp) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object varDemandTemp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    varDemand = varDemand + varDemandTemp
                    intMethodCount = intMethodCount + 1
                End If
            End If
		Next objQM
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(varDemand) And intMethodCount <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object varDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetAverageDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetAverageDemand = varDemand / intMethodCount
		Else
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetAverageDemand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetAverageDemand = System.DBNull.Value
		End If
		
	End Function
	
	Function GetDemandPartsCollection(ByRef objQM As QuantMethod) As Collection
		' Comments  : Returns a Collection of DemandPartsDB objects for the Use/Method
		'           : For displaying the subparts of the equations (A,B,C,D)
		' Parameters: ObjQM - the QuantMethod in Question
		' Returns   : Collection of DemandDB objects
		' Created   : 19-June-2002 jleiner
		'-------------------------------------------------------------------------------
		
		Dim strSQL As String
		
		Dim aValues(1) As Object
		'Dim mCol As Collection
		Dim objDBCol As New ProQdb.DemandPartsDBCollection
		'Dim objDB As DemandDB
		
		'UPGRADE_WARNING: Couldn't resolve default property of object aValues(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		aValues(0) = m_strGuidID
		
		
		strSQL = g_objDM.GetNeedSQLString(m_objUse.GetID, objQM.GetMethodologyID)
		strSQL = g_objDM.ReplaceSQLString(strSQL, aValues)
		MsgBox(strSQL)
		GetDemandPartsCollection = objDBCol.GetCollection(strSQL)
		'UPGRADE_NOTE: Object objDBCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objDBCol = Nothing
	End Function
	'SaveKitstoOrderCategory
	Function SaveKitsToOrderCategory(ByRef lngOrderCategory As Integer) As Object
		Dim objQuantification As New ProQdb.QuantificationDB
        'Dim objQM As QuantMethod
		
		SetKitsToOrderCategory(lngOrderCategory)
		
		With objQuantification
			.Load(m_strGuidID)
			.SetlngKitsToOrderCategory(m_lngKitsToOrderCategory)
			'Save the Record
			.Update()
		End With
		'UPGRADE_NOTE: Object objQuantification may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objQuantification = Nothing
	End Function
End Class