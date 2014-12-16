Option Strict On
Option Explicit On
Imports VB = Microsoft.VisualBasic
<System.Runtime.InteropServices.ProgId("Script_NET.Script")> Public Class Script
	
	'Script.cls
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	
	Const S_OK As Short = 0
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	Private m_lngID As Integer 'local Copy
	Private m_lngParentID As Integer 'local Copy
	Private m_lngQuestionID As Integer 'local Copy
	Private m_strDescription As String 'local Copy
	Private m_lngType As Integer 'local Copy
	Private m_lngGroupType As Boolean 'local Copy
	Private m_lngTrueGroup As Boolean 'local Copy
	Private m_lngFalseGroup As Boolean 'local Copy
	Private m_strTreeType As String 'local Copy
	Private m_strDefaultAction As String 'local Copy
	Private m_strMethodologyID As String 'local Copy
	Private m_strUseID As String 'local Copy
	Private m_fNewRecord As Boolean 'local copy
	Private m_fEnabled As Boolean 'Local
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' private methods
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Private Sub SaveResponse(ByRef strAggregationID As String, ByRef strQuantificationID As String, ByRef lngScriptRelationshipID As Integer, ByRef strTypeID As String, Optional ByRef lngSubActionID As Integer = 0)
		' Comments   : Saves the values in the scriptrelationship table to the response table
		' Parameters : strQuantificationID - current quantification Guid
		'            : lngScriptRelationship - ID of the current script relationship record
		'            : strActionID - field needed for treeview
		'            : strTypeID - Guid of the specifi
		' Returns    :  -
		' Modified   : 03-June-2002 LKB
		' --------------------------------------------------------
		Dim clsResponseDB As ProQdb.ResponseDB
		
		'Link to the response table via proq
		clsResponseDB = New ProQdb.ResponseDB
		
		
		'set values for the response table
		With clsResponseDB
			.SetQuantificationID(strQuantificationID)
			.SetScriptRelationshipID(lngScriptRelationshipID)
			.SetTypeID(strTypeID)
			.SetAggregationID(strAggregationID)
			.SetEnabled(m_fEnabled)
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(lngSubActionID) Then
				.SetSubActionID(lngSubActionID)
			End If
		End With
		'create record in response table
		clsResponseDB.Create()
		
		'Save the Response to the table
		clsResponseDB.Update()
		
		'UPGRADE_NOTE: Object clsResponseDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		clsResponseDB = Nothing
		
	End Sub
	
	Private Function LoadBrand(ByRef strAggregationID As String, ByRef strQuantificationID As String, ByRef lngScriptRelationshipID As Integer, Optional ByRef strBrandID As String = "", Optional ByRef lngBrandCount As Integer = 0) As Object
		' Comments    : Find all the Brands that exist
		'             : for the loaded Quantification and copy the
		'             : ID into the response table when creating the script
		' Parameters  : strQuantificationID - current quantification Guid
		'             : lngScriptRelationship - ID of the current script relationship record
		' Returns     : None
		' Created     : 04-Jun-02 LKB
		'---------------------------------------------------------
		
		Dim objCol As New ProQdb.SelectedBrandDBCollection
		Dim objDB As New ProQdb.SelectedBrandDB
		Dim mCol As Collection
		Dim strStoredProc As String
		Dim aParams(1, 1) As Object
		Dim i As Integer
        'Dim strDefaultAction As String
		
		'run stored procedure
		strStoredProc = "qselBrandsByQuantification"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(strQuantificationID)
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid

        Select Case strBrandID
            Case ""
                i = 1
            Case Else
                i = lngBrandCount + 1
        End Select

        mCol = objCol.GetCollection(strStoredProc, aParams)
        For Each objDB In mCol
            Select Case strBrandID
                Case objDB.GetID()
                    SaveResponse(strAggregationID, strQuantificationID, m_lngID, objDB.GetID(), i)
                    i = i + 1
                Case ""
                    SaveResponse(strAggregationID, strQuantificationID, m_lngID, objDB.GetID(), i)
                    i = i + 1
            End Select
        Next objDB

        'UPGRADE_NOTE: Object objCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objCol = Nothing
        'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mCol = Nothing
        'UPGRADE_NOTE: Object objDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objDB = Nothing

    End Function

    Private Function CheckMethodology(ByRef strQuantificationID As String, ByRef strMethodology As String) As Boolean
        ' Comments    : Checks if methodology is one of the selected methodologies for this quantification
        ' Parameters  : strQuantificationID - current quantification Guid
        '             : strMethodology - ID of the methodology in scriptrelationship table
        ' Returns     : True is is selected, false otherwise
        ' Created     : 04-Jun-02 LKB
        '---------------------------------------------------------

        Dim objQuantMethod As New ProQdb.QuantMethodDB
        Dim rst As DataSet 'ADODB.Recordset
        Dim strStoredProc As String
        Dim aParams(1, 1) As Object
        Dim i As Integer
        'Dim strDefaultAction As String

        CheckMethodology = False
        'run stored procedure
        strStoredProc = "qselMethodolgyByQuantification"

        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(strQuantificationID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid

        rst = objQuantMethod.ReturnRS(strStoredProc, aParams)
        If rst.Tables(strStoredProc).Rows.Count <> 0 Then
            With rst
                For i = 0 To .Tables(strStoredProc).Rows.Count - 1  'Do Until .EOF
                    'check if methodologies match
                    If .Tables(strStoredProc).Rows(i).Item("guidID").ToString = strMethodology Then
                        'set return value to true
                        CheckMethodology = True
                    End If
                    '.MoveNext()
                Next i 'Loop
                '.Close()
            End With
        End If
        'UPGRADE_NOTE: Object objQuantMethod may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objQuantMethod = Nothing
    End Function

    Private Function CheckElisaBlot(ByRef strQuantificationID As String) As Boolean
        ' Comments    : Checks if brand selected are of type elisa or blot for this quantification
        ' Parameters  : strQuantificationID - current quantification Guid
        ' Returns     : True is is selected, false otherwise
        ' Created     : 04-Jun-02 LKB
        '---------------------------------------------------------

        Dim objCol As New ProQdb.SelectedBrandDBCollection
        Dim objDB As New ProQdb.SelectedBrandDB
        Dim mCol As Collection
        Dim objBrand As New Brand

        Dim strStoredProc As String
        Dim aParams(1, 1) As Object
        'Dim strDefaultAction As String

        CheckElisaBlot = False

        'run stored procedure
        strStoredProc = "qselBrandsByQuantification"

        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(strQuantificationID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid


        mCol = objCol.GetCollection(strStoredProc, aParams)
        For Each objDB In mCol
            'check if type is not rapid for current selected brand
            objBrand.Load(objDB.GetBrandID())
            If objBrand.GetType_Renamed() <> 2 Then
                'set return value to true
                CheckElisaBlot = True
            End If
        Next objDB

        'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mCol = Nothing
        'UPGRADE_NOTE: Object objCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objCol = Nothing
        'UPGRADE_NOTE: Object objDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objDB = Nothing
        'UPGRADE_NOTE: Object objBrand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objBrand = Nothing

    End Function

    Private Function CheckColdStorage(ByRef strQuantificationID As String) As Boolean
        ' Comments    : Checks if any brand selected need cold storage for this quantification
        ' Parameters  : strQuantificationID - current quantification Guid
        ' Returns     : True is is selected, false otherwise
        ' Created     : 04-Jun-02 LKB
        '---------------------------------------------------------


        Dim objCol As New ProQdb.SelectedBrandDBCollection
        Dim objDB As New ProQdb.SelectedBrandDB
        Dim mCol As Collection
        Dim objBrand As New Brand

        Dim strStoredProc As String
        Dim aParams(1, 1) As Object
        'Dim strDefaultAction As String

        CheckColdStorage = False

        'run stored procedure
        strStoredProc = "qselBrandsByQuantification"

        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(strQuantificationID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid


        mCol = objCol.GetCollection(strStoredProc, aParams)
        For Each objDB In mCol
            'check if type is not rapid for current selected brand
            objBrand.Load(objDB.GetBrandID())
            If objBrand.GetColdStorage() = True Then
                'set return value to true
                CheckColdStorage = True
            End If
        Next objDB

        'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mCol = Nothing
        'UPGRADE_NOTE: Object objCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objCol = Nothing
        'UPGRADE_NOTE: Object objDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objDB = Nothing
        'UPGRADE_NOTE: Object objBrand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objBrand = Nothing

    End Function
    Private Function GetMaxSubAction(ByRef strQuantificationID As String) As Integer
        Dim BrandCountDB As New ProQdb.BrandCountCollectionDB
        Dim objBrandCountDB As New ProQdb.BrandCountDB
        Dim cCol As Collection
        Dim strStoredProc As String
        Dim aParams(1, 1) As Object

        GetMaxSubAction = 0
        'run stored procedure
        strStoredProc = "qlkpBrandCount"
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(strQuantificationID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid
        cCol = BrandCountDB.GetCollection(strStoredProc, aParams)

        For Each objBrandCountDB In cCol
            GetMaxSubAction = CInt(objBrandCountDB.GetBrandCount)
        Next objBrandCountDB

    End Function
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' public methods
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Function CreateQuantScript(ByRef strAggregationID As String, ByRef strQuantificationID As String, ByRef strUseID As String) As Object
        ' Comments   : Creates the script after the quantification has been configured
        ' Parameters : strQuantificationID - Current Quantification Guid
        '              strUseID - Current Use Guid
        ' Returns    :  -
        ' Modified   : 03-June-2002 LKB
        ' --------------------------------------------------------
        Dim i As Integer
        Dim clsdbScript As ProQdb.ScriptRelationshipDB
        'Link to scripRelationship table via proqdb
        clsdbScript = New ProQdb.ScriptRelationshipDB

        'For each entry on the script relationship table...
        For i = 1 To 414
            'Load the record and get the values needed
            clsdbScript.Load(i)
            m_lngID = clsdbScript.GetID
            m_lngType = clsdbScript.GetType_Renamed
            m_strMethodologyID = clsdbScript.GetMethodologyID
            m_strUseID = clsdbScript.GetUseID
            m_fEnabled = clsdbScript.GetDefaultEnabled

            If i = m_lngID Then
                'check if the use is null or equals the current use
                If m_strUseID = "" Or m_strUseID = strUseID Then
                    'before adding the the table check if methodology is null or equals one of the selected methodologies
                    'If true then copy to script (depending on type)
                    If m_strMethodologyID = "" Or CheckMethodology(strQuantificationID, m_strMethodologyID) Then
                        'check if the record has a type assocated with it Brand, Elisa/Blot, Cold Storage
                        Select Case m_lngType
                            Case 1 'Brand
                                'copy the record for each brand selected
                                LoadBrand(strAggregationID, strQuantificationID, m_lngID)
                            Case 2 'Elisa/Blot
                                If CheckElisaBlot(strQuantificationID) Then
                                    SaveResponse(strAggregationID, strQuantificationID, m_lngID, "")
                                End If
                            Case 3 'Cold Storage
                                If CheckColdStorage(strQuantificationID) Then
                                    SaveResponse(strAggregationID, strQuantificationID, m_lngID, "")
                                End If
                            Case 4 'Aggregation (only needs to be put in once, run code when creating new aggregation)

                            Case Else
                                'copy the record once
                                SaveResponse(strAggregationID, strQuantificationID, m_lngID, "")
                        End Select
                    End If
                End If
            End If
        Next i
        'UPGRADE_NOTE: Object clsdbScript may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsdbScript = Nothing
        MakeTreeViewParentTable()
    End Function
    Public Function CreateAggScript(ByRef strAggregationID As String) As Object
        ' Comments   : Creates the script after the Aggregation has been configured
        ' Parameters : strAggregationID - Current Aggregation Guid
        ' Returns    :  -
        ' Modified   : 08-June-2002 LKB
        ' --------------------------------------------------------
        Dim i As Integer
        Dim clsdbScript As ProQdb.ScriptRelationshipDB
        'Link to scripRelationship table via proqdb
        clsdbScript = New ProQdb.ScriptRelationshipDB

        'For each entry on the script relationship table...
        For i = 1 To 414
            'Load the record and get the values needed
            clsdbScript.Load(i)
            m_lngID = clsdbScript.GetID
            m_lngType = clsdbScript.GetType_Renamed

            'check if the record has type Aggregation
            If m_lngType = 4 Then 'type is aggregation
                SaveResponse(strAggregationID, "", m_lngID, "")
            End If

        Next i

        'UPGRADE_NOTE: Object clsdbScript may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsdbScript = Nothing
        MakeTreeViewParentTable()

    End Function

    Public Function CreateIndividualMethod(ByRef strAggregationID As String, ByRef strQuantificationID As String, ByRef strUseID As String, ByRef strMethodologyID As String) As Object
        ' Comments   : Creates the script after the quantification has been configured
        ' Parameters : strQuantificationID - Current Quantification Guid
        '              strUseID - Current Use Guid
        ' Returns    :  -
        ' Modified   : 03-June-2002 LKB
        ' --------------------------------------------------------
        Dim i As Integer
        Dim clsdbScript As ProQdb.ScriptRelationshipDB
        'Link to scripRelationship table via proqdb
        clsdbScript = New ProQdb.ScriptRelationshipDB

        'For each entry on the script relationship table...
        For i = 1 To 414
            'Load the record and get the values needed
            clsdbScript.Load(i)
            m_lngID = clsdbScript.GetID
            m_lngType = clsdbScript.GetType_Renamed
            m_strMethodologyID = clsdbScript.GetMethodologyID
            m_strUseID = clsdbScript.GetUseID
            m_fEnabled = clsdbScript.GetDefaultEnabled

            If i = m_lngID Then
                'before adding the the table check if methodology is null or equals one of the selected methodologies
                'If true then copy to script (depending on type)
                If m_strMethodologyID = strMethodologyID Then
                    'check if the use is null or equals the current use
                    If m_strUseID = "" Or m_strUseID = strUseID Then
                        'check if the record has a type assocated with it Brand, Elisa/Blot, Cold Storage
                        Select Case m_lngType
                            Case 1 'Brand
                                'copy the record for each brand selected
                                LoadBrand(strAggregationID, strQuantificationID, m_lngID)
                            Case 2 'Elisa/Blot
                                If CheckElisaBlot(strQuantificationID) Then
                                    SaveResponse(strAggregationID, strQuantificationID, m_lngID, "")
                                End If
                            Case 3 'Cold Storage
                                If CheckColdStorage(strQuantificationID) Then
                                    SaveResponse(strAggregationID, strQuantificationID, m_lngID, "")
                                End If
                            Case 4 'Aggregation (only needs to be put in once, run code when creating new aggregation)

                            Case Else
                                'copy the record once
                                SaveResponse(strAggregationID, strQuantificationID, m_lngID, "")
                        End Select
                    End If
                End If
            End If
        Next i
        'UPGRADE_NOTE: Object clsdbScript may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsdbScript = Nothing

        MakeTreeViewParentTable()

    End Function

    Public Function DeleteMethod(ByRef strQuantificationID As String, ByRef strMethodologyID As String) As Object
        ' Delete (From the DB)

        Dim strProc As String
        Dim aParams(2, 1) As Object
        Dim objQuantMethod As New ProQdb.QuantMethodDB

        strProc = SP_DELETE_SCRIPT_BY_QUANTMETHOD

        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(strMethodologyID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(1, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(1, 0) = New Guid(strQuantificationID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(1, 1) = DbType.Guid

        objQuantMethod.ReturnRS(strProc, aParams)
        'UPGRADE_NOTE: Object objQuantMethod may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objQuantMethod = Nothing
        MakeTreeViewParentTable()
    End Function
    Public Function CreateIndividualType(ByRef strAggregationID As String, ByRef strQuantificationID As String, ByRef strUseID As String, ByRef lngType As Integer, Optional ByRef strBrandID As String = "") As Object
        ' Comments   : Creates the script after the quantification has been configured
        ' Parameters : strQuantificationID - Current Quantification Guid
        '              strUseID - Current Use Guid
        ' Returns    :  -
        ' Modified   : 03-June-2002 LKB
        ' --------------------------------------------------------
        Dim i As Integer
        Dim clsdbScript As ProQdb.ScriptRelationshipDB
        Dim lngBrandCount As Integer
        'Link to scripRelationship table via proqdb
        clsdbScript = New ProQdb.ScriptRelationshipDB

        lngBrandCount = GetMaxSubAction(strQuantificationID)

        'For each entry on the script relationship table...
        For i = 1 To 414
            'Load the record and get the values needed
            clsdbScript.Load(i)
            m_lngID = clsdbScript.GetID
            m_lngType = clsdbScript.GetType_Renamed
            m_strMethodologyID = clsdbScript.GetMethodologyID
            m_strUseID = clsdbScript.GetUseID
            m_fEnabled = clsdbScript.GetDefaultEnabled

            If i = m_lngID Then
                If lngType = m_lngType Then
                    'check if the use is null or equals the current use
                    If m_strUseID = "" Or m_strUseID = strUseID Then
                        'before adding the the table check if methodology is null or equals one of the selected methodologies
                        'If true then copy to script (depending on type)
                        If m_strMethodologyID = "" Or CheckMethodology(strQuantificationID, m_strMethodologyID) Then
                            'check if the record has a type assocated with it Brand, Elisa/Blot, Cold Storage
                            Select Case m_lngType
                                Case 1 'Brand
                                    'copy the record for each brand selected
                                    LoadBrand(strAggregationID, strQuantificationID, m_lngID, strBrandID, lngBrandCount)
                                Case 2 'Elisa/Blot
                                    SaveResponse(strAggregationID, strQuantificationID, m_lngID, "")
                                Case 3 'Cold Storage
                                    SaveResponse(strAggregationID, strQuantificationID, m_lngID, "")
                            End Select
                        End If
                    End If
                End If
            End If
        Next i
        'UPGRADE_NOTE: Object clsdbScript may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        clsdbScript = Nothing

        MakeTreeViewParentTable()
    End Function
    Public Function DeleteType(ByRef strQuantificationID As String, ByRef lngType As Integer, Optional ByRef strBrandID As String = "") As Object
        ' Delete (From the DB)

        Dim strProc As String
        Dim objQuantMethod As New ProQdb.QuantMethodDB

        Dim aParams(3, 1) As Object
        Dim bParams(2, 1) As Object
        Select Case lngType
            Case 1
                strProc = SP_DELETE_SCRIPT_BY_TYPEID

                'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                aParams(0, 0) = lngType
                'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                aParams(0, 1) = ADODB.DataTypeEnum.adInteger
                'UPGRADE_WARNING: Couldn't resolve default property of object aParams(1, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                aParams(1, 0) = New Guid(strQuantificationID)
                'UPGRADE_WARNING: Couldn't resolve default property of object aParams(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                aParams(1, 1) = DbType.Guid
                'UPGRADE_WARNING: Couldn't resolve default property of object aParams(2, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                aParams(2, 0) = New Guid(strBrandID)
                'UPGRADE_WARNING: Couldn't resolve default property of object aParams(2, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                aParams(2, 1) = DbType.Guid

                objQuantMethod.ReturnRS(strProc, aParams)

            Case 2, 3
                strProc = SP_DELETE_SCRIPT_BY_TYPE

                'UPGRADE_WARNING: Couldn't resolve default property of object bParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                bParams(0, 0) = lngType
                'UPGRADE_WARNING: Couldn't resolve default property of object bParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                bParams(0, 1) = ADODB.DataTypeEnum.adInteger
                'UPGRADE_WARNING: Couldn't resolve default property of object bParams(1, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                bParams(1, 0) = strQuantificationID
                'UPGRADE_WARNING: Couldn't resolve default property of object bParams(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                bParams(1, 1) = DbType.Guid

                objQuantMethod.ReturnRS(strProc, bParams)

        End Select


        'UPGRADE_NOTE: Object objQuantMethod may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objQuantMethod = Nothing
        MakeTreeViewParentTable()
    End Function
    Public Function UpdateType(ByRef strQuantificationID As String, ByRef lngType As Integer, ByRef strNewTypeID As String, ByRef strOldTypeID As String) As Object
        ' Delete (From the DB)

        Dim strProc As String
        Dim objQuantMethod As New ProQdb.QuantMethodDB

        Dim aParams(4, 1) As Object
        strProc = SP_UPDATE_SCRIPT_BY_TYPEID

        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = lngType
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = ADODB.DataTypeEnum.adInteger
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(1, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(1, 0) = New Guid(strQuantificationID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(1, 1) = DbType.Guid
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(2, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(2, 0) = New Guid(strNewTypeID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(2, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(2, 1) = DbType.Guid
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(3, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(3, 0) = New Guid(strOldTypeID)
        'UPGRADE_WARNING: Couldn't resolve default property of object aParams(3, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(3, 1) = DbType.Guid

        objQuantMethod.ReturnRS(strProc, aParams)

        'UPGRADE_NOTE: Object objQuantMethod may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objQuantMethod = Nothing
        MakeTreeViewParentTable()
    End Function
	
	Public Sub MakeTreeViewParentTable()
		
		Dim objExec As New ProQdb.DBExec
		Dim aParams(0, 1) As Object
		Dim strQuery As String
		Dim strDELTable As String
		
		strDELTable = "qdelTreeViewParents"
		strQuery = "qapdTreeViewParents"
		
		objExec.ExecuteActionQuery(strDELTable, aParams)
		objExec.ExecuteActionQuery(strQuery, aParams)
		
        Dim PauseTime, Start As Double
		'UPGRADE_WARNING: Couldn't resolve default property of object PauseTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PauseTime = 5 ' Set duration.
		'UPGRADE_WARNING: Couldn't resolve default property of object Start. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Start = VB.Timer() ' Set start time.
		'UPGRADE_WARNING: Couldn't resolve default property of object PauseTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Start. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Do While VB.Timer() < Start + PauseTime
			System.Windows.Forms.Application.DoEvents() ' Yield to other processes.
		Loop 
		
		'UPGRADE_NOTE: Object objExec may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objExec = Nothing
		
	End Sub
End Class