Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("QuantificationDB_NET.QuantificationDB")> Public Class QuantificationDB
	' TODO: Declare local ADO Recordset object. For example:
    'Private WithEvents rs as dataset
    'Private Const m_strDSN As String = "DSN=JSI_ProQ;UID=;PWD=;"
    'the name of the table or view used by this class
    Private Const m_strTable As String = DB_TABLE_QUANTIFICATION
    Private Const m_strSQL As String = "Select * From " & DB_TABLE_QUANTIFICATION
    Private Const S_OK As Short = 0
    'whether or not the object has been constructed
    Private m_objConn As OleDbConnection 'OleDbConnection
    Private m_objAdapter As OleDbDataAdapter
    'Private m_objRS As DataSet 

    Private m_strID As String
    Private m_strAggregationID As String
    Private m_strUseID As String
    Private m_strProtocolID As String
    Private m_lngKitsToOrderCategory As Integer
    Private m_dtmCreated As Date
    Private m_dtmModified As Date
    Private m_strNotes As String
    Private m_fGetUseAverageMethod As Boolean
    Private m_sngDiscordancy As Single
    Private m_sngPrevalence As Single


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                         private methods
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Class_Initialize() (constructor)
    '
    'connect to the db and load the data into memory
    '
    'lbailey
    '16 may 2002
    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
        'get a connection to the db
        'Set m_objConn = GetConnection(m_strDSN)

        'load the data into memory
        'm_objRS.Open strTable, m_objConn, adOpenDynamic, adLockPessimistic, adCmdTable

    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub
    'Class_Terminate() (destructor)
    '
    'cleans up whatever needs to be cleaned up when this object is
    'released.
    'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Terminate_Renamed()
        'release the connections to the db
        'm_objRS.Close
        'm_objConn.Close
        'Set m_objConn = Nothing

    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub


    'Load()
    '
    'gets the specified record and populates this object with its values
    '
    Public Function Load(ByRef strID As String) As Object

        Dim rst As DataSet
        Dim i As Integer
        Dim strGuid As String

        m_objConn = GetConnection(DB_DSN)
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL & " WHERE GuidID = {" & strID & "}", m_objConn)
        rst = New DataSet
        m_objAdapter.Fill(rst, m_strTable)

        'select the desired record
        For i = 0 To rst.Tables(m_strTable).Rows.Count - 1
            '.Filter = DB_TABLE_PK & " = '" & strID & "'"
            strGuid = CType((rst.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK)), Guid).ToString()
            If strGuid = strID Then
                With rst.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() '.Fields("GuidID").Value
                    m_strAggregationID = CType((.Item("guidID_Aggregation")), Guid).ToString() '.Fields("guidID_Aggregation").Value
                    m_strUseID = CType((.Item("guidID_Use")), Guid).ToString() '.Fields("guidID_Use").Value
                    m_strProtocolID = CType((.Item("guidID_Protocol")), Guid).ToString() '.Fields("guidID_Protocol").Value
                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                    m_strNotes = IIf(IsDBNull(.Item("memNotes")), "", .Item("memNotes")) 'IIf(IsDBNull(.Fields("memNotes").Value), "", .Fields("memNotes").Value)
                    m_dtmCreated = .Item("dtmCreated") '.Fields("dtmCreated").Value
                    m_dtmModified = .Item("dtmModified") '.Fields("dtmModified").Value
                    m_lngKitsToOrderCategory = .Item("lngKitsToOrderCategory") '.Fields("lngKitsToOrderCategory").Value
                    m_fGetUseAverageMethod = .Item("fUseAverageMethod") '.Fields("fUseAverageMethod").Value
                    m_sngDiscordancy = .Item("sngDiscordancy") '.Fields("sngDiscordancy").Value
                    m_sngPrevalence = .Item("sngPrevalence") '.Fields("sngPrevalence").Value
                End With

                CloseDB(m_objConn, rst)
                Exit For
            End If
        Next

        'return result code
        'UPGRADE_WARNING: Couldn't resolve default property of object Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Load = S_OK

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function

    Public Sub Update()

        Dim i As Integer
        Dim cb As OleDb.OleDbCommandBuilder
        Dim rst As New DataSet
        'connect to the db and get a recordset
        m_objConn = GetConnection(DB_DSN)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        m_objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(m_objAdapter)

        With rst
            For i = 0 To .Tables(m_strTable).Rows.Count - 1
                Dim strGuid As String
                strGuid = CType(.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK), Guid).ToString
                If strGuid = m_strID Then
                    'stuff all of the record's properties
                    .Tables(m_strTable).Rows(i).Item("guidID_Aggregation") = New Guid(m_strAggregationID)
                    .Tables(m_strTable).Rows(i).Item("guidID_Protocol") = New Guid(m_strProtocolID)
                    .Tables(m_strTable).Rows(i).Item("guidID_Use") = New Guid(m_strUseID)
                    .Tables(m_strTable).Rows(i).Item("memNotes") = m_strNotes
                    .Tables(m_strTable).Rows(i).Item("dtmCreated") = m_dtmCreated
                    .Tables(m_strTable).Rows(i).Item("dtmModified") = m_dtmModified
                    .Tables(m_strTable).Rows(i).Item("lngKitsToOrderCategory") = m_lngKitsToOrderCategory
                    .Tables(m_strTable).Rows(i).Item("fUseAverageMethod") = m_fGetUseAverageMethod
                    .Tables(m_strTable).Rows(i).Item("sngDiscordancy") = m_sngDiscordancy
                    .Tables(m_strTable).Rows(i).Item("sngPrevalence") = m_sngPrevalence                    'write the record
                    m_objAdapter.Update(rst, m_strTable) '.Update()

                    Exit For
                End If
            Next
        End With

        'clean up
        CloseDB(m_objConn, rst)

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Sub


    '+
    'Create()
    '
    'Adds the New Record to the db using the info that we've got
    '
    '-
    Public Function Create() As String

        'get a connection to the db
        m_objConn = GetConnection(DB_DSN)
        Dim cb As OleDb.OleDbCommandBuilder
        Dim dsNewRow As DataRow
        Dim rst As New DataSet

        'connect to the db and get a recordset
        m_objConn = GetConnection(DB_DSN)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        m_objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(m_objAdapter)

        dsNewRow = rst.Tables(m_strTable).NewRow()

        'if we don't already have an id for this, then create one
        If (Len(m_strID) < 1) Then
            'use the guid generator to create an id
            'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            m_strID = g_objGUIDGenerator.GetGUID()
        End If
        'stuff all of the record's properties
        dsNewRow.Item("guidID") = New Guid(m_strID) 'dsNewRow.Item("guidID").Value = m_strID
        dsNewRow.Item("guidID_Aggregation") = New Guid(m_strAggregationID)
        dsNewRow.Item("guidID_Use") = IIf(Len(m_strUseID) > 0, New Guid(m_strUseID), System.DBNull.Value)
        dsNewRow.Item("guidID_Protocol") = IIf(Len(m_strProtocolID) > 0, New Guid(m_strProtocolID), System.DBNull.Value)
        dsNewRow.Item("memNotes") = IIf(Len(m_strNotes) > 0, m_strNotes, System.DBNull.Value)
        dsNewRow.Item("dtmCreated") = m_dtmCreated
        dsNewRow.Item("dtmModified") = m_dtmModified
        dsNewRow.Item("lngKitsToOrderCategory") = m_lngKitsToOrderCategory
        dsNewRow.Item("fUseAverageMethod") = m_fGetUseAverageMethod
        dsNewRow.Item("sngDiscordancy") = m_sngDiscordancy
        dsNewRow.Item("sngPrevalence") = m_sngPrevalence
        rst.Tables(m_strTable).Rows.Add(dsNewRow)

        m_objAdapter.Update(rst, m_strTable)

        'clean up
        CloseDB(m_objConn, rst)

        'return the id of the created object
        Create = m_strID

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function

    Public Function Delete(ByVal strID As String) As Object

        Dim rst As New DataSet

        DeleteRecord(m_objConn, rst, m_strTable, m_strSQL, strID)
        CloseDB(m_objConn, rst)
        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function

    Public Function ReturnRS(ByRef strStoredProc As String, ByRef aParams(,) As Object) As dataset
        Dim objDbExec As New DBExec
        m_objConn = GetConnection(DB_DSN)
        ReturnRS = objDbExec.ReturnRS(m_objConn, strStoredProc, aParams)
        'UPGRADE_NOTE: Object objDbExec may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objDbExec = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function
    
	
	Private Sub Class_GetDataMember(ByRef DataMember As String, ByRef Data As Object)
		' TODO:  Return the appropriate recordset based on DataMember. For example:
		
		'Select Case DataMember
		'Case ""             ' Default
		'    Set Data = Nothing
		'Case Else           ' Default
		'    Set Data = rs
		'End Select
	End Sub
	
	'+
	'standard manipulator functions
	'
	'lbailey
	'16 may 2002
	'-
	Public Function SetID(ByRef strID As String) As Short
		'set the member
		m_strID = strID
		'set the return code
		SetID = S_OK
	End Function
	Public Function SetAggregationID(ByRef strAggregateID As String) As Short
		'set the member
		m_strAggregationID = strAggregateID
		'set the return code
		SetAggregationID = S_OK
	End Function
	Public Function SetUseID(ByRef strUseID As String) As Short
		'set the member
		m_strUseID = strUseID
		'set the return code
		SetUseID = S_OK
	End Function
	Public Function SetProtocolID(ByRef strProtocolID As String) As Short
		'set the member
		m_strProtocolID = strProtocolID
		'set the return code
		SetProtocolID = S_OK
	End Function
	Public Function SetlngKitsToOrderCategory(ByRef lngKitsToOrderCategory As Integer) As Short
		'set the member
		m_lngKitsToOrderCategory = lngKitsToOrderCategory
		'set the return code
		SetlngKitsToOrderCategory = S_OK
	End Function
	
	Public Function SetNotes(ByRef strNotes As String) As Short
		'set the member
		m_strNotes = strNotes
		'set the return code
		SetNotes = S_OK
	End Function
	
	Public Function SetCreationDate(ByRef dCreationDate As Date) As Short
		'set the member
		m_dtmCreated = dCreationDate
		'set the return code
		SetCreationDate = S_OK
	End Function
	
	Public Function SetModifiedDate(ByRef dModifiedDate As Date) As Short
		'set the member
		m_dtmModified = dModifiedDate
		'set the return code
		SetModifiedDate = S_OK
	End Function
	
	Public Sub setUseAverageMethod(ByRef fUse As Boolean)
		m_fGetUseAverageMethod = fUse
	End Sub
	Public Sub setDiscordancy(ByRef sngDiscordancy As Single)
		m_sngDiscordancy = sngDiscordancy
	End Sub
	Public Sub setPrevalence(ByRef sngPrevalence As Single)
		m_sngPrevalence = sngPrevalence
	End Sub
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'standard accessor functions
	
	Public Function GetID() As String
		GetID = m_strID
	End Function
	
	Public Function GetAggregationID() As String
		GetAggregationID = m_strAggregationID
	End Function
	
	Public Function GetProtocolID() As String
		GetProtocolID = m_strProtocolID
	End Function
	
	Public Function GetUseID() As String
		GetUseID = m_strUseID
	End Function
	
	Public Function GetlngKitsToOrderCategory() As Integer
		GetlngKitsToOrderCategory = m_lngKitsToOrderCategory
	End Function
	
	Public Function GetCreationDate() As String
		GetCreationDate = CStr(m_dtmCreated)
	End Function
	
	Public Function GetModifiedDate() As String
		GetModifiedDate = CStr(m_dtmModified)
	End Function
	
	Public Function GetNotes() As String
		GetNotes = m_strNotes
	End Function
	
	Public Function getUseAverageMethod() As Boolean
		getUseAverageMethod = m_fGetUseAverageMethod
	End Function
	Public Function getDiscordancy() As Single
		getDiscordancy = m_sngDiscordancy
	End Function
	Public Function getPrevalence() As Single
		getPrevalence = m_sngPrevalence
	End Function
End Class